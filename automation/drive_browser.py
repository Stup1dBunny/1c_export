"""Publish through the user's Chrome login and verify the downloaded copy.

No cookies/tokens are extracted and no sharing settings are changed. An
uncertain earlier upload must be resolved before attempting another upload.
"""
from __future__ import annotations

import hashlib
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, urlsplit
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from onec_ui import OneCUI
from report_tools import inspect_xlsx, write_json, upload_lock, copy_checked, report_folder


class DriveAuthorizationRequired(PermissionError):
    """The user must restore Google's session; never automate their password."""


class DriveBrowser:
    def __init__(self, driver, downloads: Path, folder_id: str, state: Path, folder_name: str = '1c-reports'):
        self.driver = driver
        self.ui = OneCUI(driver, downloads)
        self.folder_id = folder_id
        self.folder_name = folder_name
        self.folder_url = 'https://drive.google.com/drive/folders/' + folder_id
        self.state = state

    def open_folder(self):
        self.driver.get(self.folder_url)
        self.require_drive_session()
        self.ui.one('[role="button"][aria-label="' + self.folder_name + '"]', 60)

    def require_drive_session(self):
        if urlsplit(self.driver.current_url).hostname == 'accounts.google.com':
            raise DriveAuthorizationRequired(
                'Google Drive requires sign-in in the robot Chrome profile; prepared reports are retained')

    def click(self, selector):
        # Drive uses ordinary DOM buttons. 1C's custom menus need pointer
        # actions, but Drive toolbar updates are more reliable with WebDriver
        # element clicks after locating the current visible control.
        self.ui.one(selector).click()

    def for_day(self, cfg, day, end_day=None):
        if not cfg.get('drive_daily_folders'):
            return self
        if cfg.get('drive_month_folders'):
            raise ValueError('Monthly folder hierarchy requires the configured Drive API backend')
        return self.ensure_subfolder(report_folder(cfg, day, end_day))

    def ensure_subfolder(self, name):
        """Resolve exactly one child folder; recover a lost creation response.

        A pending creation with no search result is deliberately not repeated.
        Opening the result also proves it is a folder before uploading into it.
        """
        if not name or any(c in name for c in ('"', '\n', '\r', '/', '\\')):
            raise ValueError('Unsupported daily folder name')
        identity = hashlib.sha256((self.folder_id + ':' + name).encode()).hexdigest()
        state = self.state / 'folders'
        journal = state / (identity + '.json')
        with upload_lock(state):
            saved = json.loads(journal.read_text(encoding='utf-8')) if journal.exists() else {}
            found = self.find(name)
            if not found:
                if saved and saved.get('status') != 'prepared':
                    raise ValueError('A previous daily folder creation is uncertain or the folder is missing')
                write_json(journal, dict(name=name, parent_id=self.folder_id, status='prepared'))
                self.open_folder()
                self.click('button[aria-label="Создать"]')
                self.click('[role="menuitem"][data-key="1"]')
                field = self.ui.one('[role="dialog"] input[aria-label="Новая папка"]')
                field.send_keys(Keys.CONTROL, 'a')
                field.send_keys(name)
                if field.get_attribute('value') != name:
                    raise ValueError('Daily folder name was not entered correctly')
                write_json(journal, dict(name=name, parent_id=self.folder_id, status='pending'))
                self.click('[role="dialog"] button[data-mdc-dialog-action="ok"]')
                def created(_):
                    matches = [r for r in self.rows() if r['name'] == name]
                    if len(matches) > 1:
                        raise ValueError('Duplicate daily folders after creation')
                    return matches[0] if matches else False
                found = WebDriverWait(self.driver, 60).until(created)
            if saved.get('folder_id') and saved['folder_id'] != found['id']:
                raise ValueError('The daily folder identity changed')
            child = DriveBrowser(self.driver, self.ui.downloads, found['id'], self.state, name)
            child.open_folder()
            write_json(journal, dict(name=name, parent_id=self.folder_id,
                                    folder_id=found['id'], status='verified'))
            return child

    def rows(self):
        return self.driver.execute_script(r'''
            return [...document.querySelectorAll('[role=row][data-id]')]
                .map(e=>({id:e.getAttribute('data-id'),
                    name:e.querySelector('.i92Sbe[data-tooltip]')?.textContent || e.innerText.split('\n')[0]}))
                .filter(e=>e.name);
        ''')

    def find(self, filename):
        # This query shape was produced by Drive's own advanced-search UI
        # with the requested folder selected in its location picker.
        if any(c in filename for c in ('"', '\n', '\r')):
            raise ValueError('Unsupported Drive search filename')
        query = 'title:"' + filename + '" parent:' + self.folder_id
        self.driver.get('https://drive.google.com/drive/search?q=' + quote(query))
        def ready(_):
            self.require_drive_session()
            if self.rows():
                return True
            return 'По вашему запросу не найдено ни одного объекта' in self.driver.execute_script('return document.body.innerText')
        WebDriverWait(self.driver, 60).until(ready)
        matches = [r for r in self.rows() if r['name'] == filename]
        if len(matches) > 1:
            raise ValueError('More than one remote file has the requested name')
        return matches[0] if matches else None

    def download_copy(self, file_id):
        before = self.ui.download_snapshot()
        # Drive can return a search row before making it visible and clickable.
        # Waiting for that exact known ID is safe and never resubmits an upload.
        self.ui.one('[role="row"][data-id="' + file_id + '"]', 90).click()
        # Drive exposes different numeric menu keys in its UI variants. The
        # visible selection toolbar has the same Download label in both.
        def selected_only_target(_):
            selected = self.driver.execute_script('''
                return [...document.querySelectorAll('[role=row][data-id][aria-selected=true]')]
                    .map(e=>e.getAttribute('data-id'));
            ''')
            return selected == [file_id]
        WebDriverWait(self.driver, 15).until(selected_only_target)
        self.click('[aria-label="Скачать"]')
        # A Drive file can acquire a new version without changing its ID.
        # Keep each distinct downloaded version by content hash so verifying
        # an update never overwrites the earlier evidence or collides with it.
        temporary = self.state / 'verified' / 'downloads' / (uuid.uuid4().hex + '.xlsx')
        copied = self.ui.wait_xlsx(before, temporary)
        archived = copy_checked(temporary, self.state / 'verified' / file_id / (copied['sha256'] + '.xlsx'))
        temporary.unlink()
        return archived

    def verify_copy(self, file_id, expected_sha):
        copied = self.download_copy(file_id)
        if copied['sha256'] != expected_sha:
            raise ValueError('Downloaded Drive copy differs from the local report')
        return copied

    def upload_version(self, source, file_id, before_submit):
        """Use Drive's observed Manage versions dialog, preserving file ID."""
        self.click('[role="row"][data-id="' + file_id + '"]')
        self.click('[role="button"][aria-label="Другие действия"]')
        self.click('[role="menuitem"][aria-label="Информация о файле"]')
        version_menu = self.ui.one('[role="menuitem"][data-key="26"]')
        if version_menu.text.strip() != 'Управление версиями':
            raise ValueError('Drive version menu changed')
        version_menu.click()
        dialog = self.ui.one('[role="dialog"].a-Fb-k')
        if source.name not in dialog.text:
            raise ValueError('The version dialog belongs to a different file')
        existing = WebDriverWait(self.driver, 60).until(lambda _: {
            e.get_attribute('data-revisionid') for e in
            dialog.find_elements(By.CSS_SELECTOR, '[role="listitem"][data-revisionid]')
            if e.is_displayed() and e.get_attribute('data-revisionid')})
        self.driver.execute_cdp_cmd('Page.setInterceptFileChooserDialog', {'enabled': True})
        try:
            self.ui.text('Загрузить новую версию').click()
            inputs = self.driver.find_elements(By.CSS_SELECTOR, 'input[type="file"]:not([multiple])')
            if len(inputs) != 1:
                raise ValueError('The version upload file picker is ambiguous')
            before_submit()
            inputs[0].send_keys(str(source.resolve()))
        finally:
            self.driver.execute_cdp_cmd('Page.setInterceptFileChooserDialog', {'enabled': False})

        def version_arrived(_):
            rows = dialog.find_elements(By.CSS_SELECTOR, '[role="listitem"][data-revisionid]')
            current = [r for r in rows if r.is_displayed() and r.text.startswith('Текущая версия')]
            return (len(current) == 1 and source.name in current[0].text
                    and current[0].get_attribute('data-revisionid') not in existing)
        WebDriverWait(self.driver, 180).until(version_arrived)
        dialog.find_element(By.CSS_SELECTOR, 'button[name="ok"]').click()

    def replace_known(self, source, checked, saved, journal, found, expected_previous_sha):
        if (not found or saved.get('status') != 'verified'
                or saved.get('file_id') != found['id']
                or expected_previous_sha != saved.get('sha256')):
            raise ValueError('Replacement requires the reviewed previous report and its known Drive ID')
        pending = saved.get('replacement')
        if pending and (pending.get('sha256') != checked['sha256']
                        or pending.get('previous_sha256') != expected_previous_sha):
            raise ValueError('Another report replacement is unfinished')
        remote = self.download_copy(found['id'])
        if remote['sha256'] not in {checked['sha256'], expected_previous_sha}:
            raise ValueError('The Drive report changed after review; it was not replaced')
        if remote['sha256'] == expected_previous_sha:
            if pending and pending.get('status') != 'prepared':
                raise ValueError('A submitted replacement is uncertain; inspect Drive before retrying its upload')
            replacement = dict(sha256=checked['sha256'], previous_sha256=expected_previous_sha,
                               source=str(source.resolve()), status='prepared')
            write_json(journal, dict(saved, replacement=replacement))
            def before_submit():
                replacement['status'] = 'pending'
                write_json(journal, dict(saved, replacement=replacement))
            self.upload_version(source, found['id'], before_submit)
            replacement['status'] = 'uploaded'
            write_json(journal, dict(saved, replacement=replacement))
            located = self.find(source.name)
            if not located or located['id'] != found['id']:
                raise ValueError('The replaced report cannot yet be found at its original Drive ID')
            self.verify_copy(found['id'], checked['sha256'])
        # This also recovers a successful upload whose completion response was lost.
        history = list(saved.get('history', [])) + [dict(sha256=expected_previous_sha,
            file_id=found['id'], replaced_at=datetime.now(timezone.utc).isoformat())]
        result = dict(filename=source.name, file_id=found['id'], sha256=checked['sha256'],
                      folder_id=self.folder_id, status='verified', replaced=True,
                      verified_by='downloaded SHA-256', history=history)
        write_json(journal, result)
        return result

    def publish(self, source: Path, *, replace_previous_sha=None):
        checked = inspect_xlsx(source)
        identity = hashlib.sha256((self.folder_id + ':' + source.name).encode()).hexdigest()
        journal = self.state / (identity + '.json')
        with upload_lock(self.state):
            saved = json.loads(journal.read_text(encoding='utf-8')) if journal.exists() else {}
            if replace_previous_sha is not None and saved.get('status') != 'verified':
                raise ValueError('A replacement requires a previously verified publication')
            pending = saved.get('replacement')
            if pending and (checked['sha256'] != pending.get('sha256')
                            or replace_previous_sha != pending.get('previous_sha256')):
                raise ValueError('Resume the unfinished replacement with its reviewed new report')
            changed = saved and saved.get('sha256') != checked['sha256']
            if changed and replace_previous_sha is None:
                raise ValueError('A different local report already has this publication name')
            found = self.find(source.name)
            if changed:
                return self.replace_known(source, checked, saved, journal, found, replace_previous_sha)
            if found:
                if saved.get('status') == 'verified' and saved.get('file_id') == found['id']:
                    # A user can edit a Drive file without changing its ID.
                    # Recheck its contents before treating a repeat as complete.
                    self.verify_copy(found['id'], checked['sha256'])
                    return dict(saved, status='unchanged')
                self.verify_copy(found['id'], checked['sha256'])
                result = dict(filename=source.name, file_id=found['id'], sha256=checked['sha256'],
                              folder_id=self.folder_id, status='verified', verified_by='downloaded SHA-256')
                write_json(journal, result)
                return result
            if saved and saved.get('status') != 'prepared':
                raise ValueError('A previous upload is uncertain or missing; refusing to create a duplicate')
            # Opening a menu cannot submit a file. Keep this phase retryable;
            # mark uncertain immediately BEFORE passing the file to Chrome.
            write_json(journal, dict(filename=source.name, sha256=checked['sha256'], status='prepared'))
            self.driver.get(self.folder_url)
            self.ui.one('[role="button"][aria-label="' + self.folder_name + '"]', 60)
            for attempt in range(2):
                if attempt:
                    # Observed Drive UI occasionally ignored Create until a
                    # refresh. Retrying this navigation cannot duplicate an upload.
                    self.driver.refresh()
                    self.ui.one('[role="button"][aria-label="' + self.folder_name + '"]', 60)
                self.click('button[aria-label="Создать"]')
                try:
                    self.ui.one('[role="menuitem"][data-key="17"]', 30)
                    break
                except TimeoutException:
                    if attempt:
                        raise
            self.driver.execute_cdp_cmd('Page.setInterceptFileChooserDialog', {'enabled': True})
            try:
                self.click('[role="menuitem"][data-key="17"]')
                upload = self.driver.find_element(By.CSS_SELECTOR, 'input[type="file"][multiple]')
                write_json(journal, dict(filename=source.name, sha256=checked['sha256'], status='pending'))
                upload.send_keys(str(source.resolve()))
            finally:
                self.driver.execute_cdp_cmd('Page.setInterceptFileChooserDialog', {'enabled': False})
            def uploaded(_):
                matches = [r for r in self.rows() if r['name'] == source.name]
                if len(matches) > 1:
                    raise ValueError('Duplicate remote filename after upload')
                panels = self.driver.execute_script('''
                    return [...document.querySelectorAll('[role=dialog]')]
                        .filter(e=>e.offsetWidth).map(e=>e.innerText);
                ''')
                done = any(p.startswith(('Загружен ', 'Загружены ')) and source.name in p for p in panels)
                return matches[0] if done and len(matches) == 1 else False
            found = WebDriverWait(self.driver, 180).until(uploaded)
            write_json(journal, dict(filename=source.name, file_id=found['id'], sha256=checked['sha256'], status='uploaded'))
            # In a growing daily archive the uploaded row may be below the
            # visible folder viewport. Search isolates it before downloading.
            visible_copy = self.find(source.name)
            if not visible_copy or visible_copy['id'] != found['id']:
                raise ValueError('Completed upload is not yet locatable for verification; retry later')
            self.verify_copy(found['id'], checked['sha256'])
            result = dict(filename=source.name, file_id=found['id'], sha256=checked['sha256'],
                          folder_id=self.folder_id, status='verified', verified_by='downloaded SHA-256')
            write_json(journal, result)
            return result
