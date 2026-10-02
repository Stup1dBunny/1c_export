"""Folder-scoped XLSX publication through the official Google Drive API.

OAuth may authorize the user's Drive, but this publisher only searches and
writes in the configured report folder and its direct daily folders. It never
changes sharing, deletes files, or exports business data to another destination.
"""
from __future__ import annotations

import hashlib
import io
import json
from pathlib import Path
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from report_tools import MIME, inspect_xlsx, write_json, upload_lock, report_folder

FOLDER = 'application/vnd.google-apps.folder'
FIELDS = 'id,name,mimeType,parents,trashed,sha256Checksum,size,capabilities(canAddChildren)'


def query_literal(value):
    return "'" + value.replace('\\', '\\\\').replace("'", "\\'") + "'"


class DriveAPI:
    def __init__(self, service, folder_id, state, folder_name='1c-reports', root_id=None, parent_id=None):
        self.service = service
        self.folder_id, self.folder_name = folder_id, folder_name
        self.root_id = root_id or folder_id
        self.parent_id = parent_id or (root_id if root_id and root_id != folder_id else None)
        self.state = Path(state)

    def metadata(self, file_id):
        try:
            return self.service.files().get(fileId=file_id, fields=FIELDS,
                supportsAllDrives=True).execute(num_retries=3)
        except HttpError as error:
            if error.resp.status == 404:
                return None
            raise

    def verify_folder(self):
        folder = self.metadata(self.folder_id)
        if (not folder or folder.get('trashed') or folder.get('mimeType') != FOLDER
                or not folder.get('capabilities', {}).get('canAddChildren')):
            raise PermissionError('Cannot add reports to the configured Drive folder')
        if self.folder_id != self.root_id:
            if self.parent_id not in folder.get('parents', []):
                raise ValueError('The daily folder was moved outside the configured report folder')
            parent = self.parent_id
            for _ in range(4):
                if parent == self.root_id:
                    break
                ancestor = self.metadata(parent)
                if (not ancestor or ancestor.get('trashed') or ancestor.get('mimeType') != FOLDER
                        or len(ancestor.get('parents', [])) != 1):
                    raise ValueError('The report folder hierarchy is outside the configured root')
                parent = ancestor['parents'][0]
            else:
                raise ValueError('The report folder hierarchy is outside the configured root')

    def find(self, name):
        query = f'{query_literal(self.folder_id)} in parents and trashed = false and name = {query_literal(name)}'
        found, token = [], None
        while True:
            response = self.service.files().list(q=query, fields=f'nextPageToken,files({FIELDS})',
                pageToken=token, pageSize=1000, supportsAllDrives=True,
                includeItemsFromAllDrives=True).execute(num_retries=3)
            found.extend(response.get('files', []))
            token = response.get('nextPageToken')
            if not token:
                break
        if len(found) > 1:
            raise ValueError('More than one Drive item has the requested name in this folder')
        return found[0] if found else None

    def journal_path(self, name, folders=False):
        key = hashlib.sha256((self.folder_id + ':' + name).encode()).hexdigest()
        return (self.state / 'folders' if folders else self.state) / (key + '.json')

    def reserved_id(self):
        return self.service.files().generateIds(count=1, space='drive', type='files').execute(num_retries=3)['ids'][0]

    def create_reserved(self, metadata, source=None):
        options = dict(body=metadata, fields=FIELDS, supportsAllDrives=True)
        if source is not None:
            options['media_body'] = MediaFileUpload(str(source), mimetype=MIME, resumable=True)
        try:
            return self.service.files().create(**options).execute(num_retries=3)
        except HttpError as error:
            if error.resp.status != 409:
                raise
            # A lost successful response is recovered by its pre-reserved ID.
            existing = self.metadata(metadata['id'])
            if not existing:
                raise ValueError('The reserved Drive ID cannot yet be resolved') from error
            return existing

    def ensure_subfolder(self, name):
        self.verify_folder()
        journal = self.journal_path(name, folders=True)
        with upload_lock(self.state / 'folders'):
            saved = json.loads(journal.read_text(encoding='utf-8')) if journal.exists() else {}
            found = self.find(name)
            known_id = saved.get('folder_id')
            if not found and known_id:
                found = self.metadata(known_id)
                if not found and saved.get('status') != 'pending':
                    raise ValueError('The previously verified daily folder is missing')
            if not found:
                if saved and not known_id and saved.get('status') != 'prepared':
                    raise ValueError('An earlier browser folder creation is uncertain; inspect it before retrying')
                folder_id = known_id or self.reserved_id()
                write_json(journal, dict(name=name, parent_id=self.folder_id, folder_id=folder_id, status='pending'))
                found = self.create_reserved(dict(id=folder_id, name=name, mimeType=FOLDER, parents=[self.folder_id]))
            if found.get('trashed'):
                raise ValueError('The daily report folder is in Drive trash')
            if (found.get('name') != name or found.get('mimeType') != FOLDER
                    or self.folder_id not in found.get('parents', [])
                    or known_id and found['id'] != known_id):
                raise ValueError('The daily folder identity or location changed')
            child = DriveAPI(self.service, found['id'], self.state, name, root_id=self.root_id, parent_id=self.folder_id)
            child.verify_folder()
            write_json(journal, dict(name=name, parent_id=self.folder_id, folder_id=found['id'], status='verified'))
            return child

    def for_day(self, cfg, day, end_day=None):
        if not cfg.get('drive_daily_folders'):
            return self
        name = report_folder(cfg, day, end_day)
        if not cfg.get('drive_month_folders'):
            return self.ensure_subfolder(name)
        month = self.ensure_subfolder(day.strftime('%Y.%m_отчёты'))
        if end_day is None:
            self.move_legacy_day(month, name)
        return month.ensure_subfolder(name)

    def move_legacy_day(self, month, name):
        """Move only the matching daily folder within this report root, retaining IDs.

        The destination journal records the ID before the API call. A lost
        response is recovered whether the folder is still at root or has moved.
        """
        self.verify_folder()
        month.verify_folder()
        legacy_journal = self.journal_path(name, folders=True)
        target_journal = month.journal_path(name, folders=True)
        with upload_lock(self.state / 'folders'):
            old = json.loads(legacy_journal.read_text(encoding='utf-8')) if legacy_journal.exists() else {}
            target = json.loads(target_journal.read_text(encoding='utf-8')) if target_journal.exists() else {}
            found = self.find(name)
            destination = month.find(name)
            known_id = old.get('folder_id') or (target.get('folder_id') if target.get('status') == 'moving' else None)
            if not found and known_id:
                found = self.metadata(known_id)
                if not found:
                    raise ValueError('The previously verified daily folder is missing')
            if found is None:
                return
            if found.get('trashed'):
                raise ValueError('The daily report folder is in Drive trash')
            if (found.get('mimeType') != FOLDER or found.get('name') != name
                    or found.get('parents') not in ([self.folder_id], [month.folder_id])
                    or known_id and found['id'] != known_id
                    or destination and destination['id'] != found['id']
                    or target.get('folder_id') and target['folder_id'] != found['id']):
                raise ValueError('The daily folder identity or location changed')
            evidence = dict(name=name, parent_id=month.folder_id, folder_id=found['id'])
            if found.get('parents') == [self.folder_id]:
                write_json(target_journal, dict(evidence, status='moving', previous_parent_id=self.folder_id))
                self.service.files().update(fileId=found['id'], addParents=month.folder_id,
                    removeParents=self.folder_id, fields=FIELDS, supportsAllDrives=True).execute(num_retries=3)
                found = self.metadata(found['id'])
                if (not found or found.get('trashed') or found.get('parents') != [month.folder_id]
                        or found.get('name') != name or found.get('mimeType') != FOLDER):
                    raise ValueError('The daily folder move was not verified')
            write_json(target_journal, dict(evidence, status='verified'))
            if old:
                write_json(legacy_journal, dict(old, status='relocated', relocated_parent_id=month.folder_id))

    def content_sha(self, file):
        if file.get('sha256Checksum'):
            return file['sha256Checksum']
        stream = io.BytesIO()
        download = MediaIoBaseDownload(stream, self.service.files().get_media(fileId=file['id'], supportsAllDrives=True))
        done = False
        while not done:
            _, done = download.next_chunk(num_retries=3)
        return hashlib.sha256(stream.getvalue()).hexdigest()

    def validate_file(self, file, name):
        if (file.get('name') != name or file.get('trashed')
                or self.folder_id not in file.get('parents', []) or file.get('mimeType') != MIME):
            raise ValueError('The Drive file identity, type or parent does not match the report')

    def publish(self, source, *, replace_previous_sha=None):
        source = Path(source)
        checked = inspect_xlsx(source)
        self.verify_folder()
        journal = self.journal_path(source.name)
        with upload_lock(self.state):
            saved = json.loads(journal.read_text(encoding='utf-8')) if journal.exists() else {}
            if saved and saved.get('sha256') != checked['sha256'] and replace_previous_sha != saved.get('sha256'):
                raise ValueError('Changed reports require the reviewed previous SHA-256')
            found = self.find(source.name)
            known_id = saved.get('file_id')
            if not found and known_id:
                found = self.metadata(known_id)
                if not found and saved.get('status') != 'pending':
                    raise ValueError('The previously published file is missing; refusing a duplicate')
            if found:
                self.validate_file(found, source.name)
                if known_id and known_id != found['id']:
                    raise ValueError('The published Drive file identity changed')
                remote_sha = self.content_sha(found)
                if remote_sha == checked['sha256']:
                    result = self.record_verified(journal, source.name, found['id'], checked['sha256'])
                    return dict(result, status='unchanged')
                if (replace_previous_sha is None or not saved or saved.get('status') != 'verified'
                        or saved.get('sha256') != replace_previous_sha or remote_sha != replace_previous_sha):
                    raise ValueError('The existing Drive file differs; no overwrite was performed')
                # Same ID is retained. On an interrupted update, the next run
                # verifies whether the old or reviewed new content is present.
                found = self.service.files().update(fileId=found['id'],
                    media_body=MediaFileUpload(str(source), mimetype=MIME, resumable=True),
                    fields=FIELDS, supportsAllDrives=True).execute(num_retries=3)
            else:
                if replace_previous_sha is not None:
                    raise ValueError('Cannot replace a file that does not exist')
                if saved and not known_id and saved.get('status') != 'prepared':
                    raise ValueError('An earlier browser upload is uncertain; inspect it before retrying')
                file_id = known_id or self.reserved_id()
                write_json(journal, dict(filename=source.name, file_id=file_id, folder_id=self.folder_id,
                                        sha256=checked['sha256'], status='pending'))
                found = self.create_reserved(dict(id=file_id, name=source.name, mimeType=MIME,
                                                 parents=[self.folder_id]), source)
            self.validate_file(found, source.name)
            if self.content_sha(found) != checked['sha256']:
                raise ValueError('The uploaded Drive content does not match the local report')
            return self.record_verified(journal, source.name, found['id'], checked['sha256'])

    def record_verified(self, journal, name, file_id, sha):
        result = dict(filename=name, file_id=file_id, folder_id=self.folder_id,
                      sha256=sha, status='verified', verified_by='Drive API content SHA-256')
        write_json(journal, result)
        return result
