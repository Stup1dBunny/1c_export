"""Resume one base/day after interrupted read-only 1C operations."""
import hashlib
import json
from pathlib import Path

from report_tools import write_json


class SourceCache:
    NAMES = {'receipts.xlsx', 'stock.xlsx', 'alcohol.xlsx', 'checks.tsv'}

    def __init__(self, folder: Path, context: dict):
        self.folder = folder
        self.journal = folder / 'source-progress.json'
        folder.mkdir(parents=True, exist_ok=True)
        if self.journal.exists():
            self.data = json.loads(self.journal.read_text(encoding='utf-8'))
            if self.data.get('context') != context:
                raise ValueError('Incomplete source run belongs to a different date/base/configuration')
        else:
            if any(folder.iterdir()):
                raise ValueError('Untracked files exist in the source recovery directory')
            self.data = {'context': context, 'files': {}, 'contacts': {}}
            self.save()

    def save(self):
        write_json(self.journal, self.data)

    def source(self, name, collect):
        if name not in self.NAMES:
            raise ValueError('Unsupported source cache filename')
        path = self.folder / name
        saved = self.data['files'].get(name, {})
        if saved.get('sha256'):
            if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != saved['sha256']:
                raise ValueError('A saved source was removed or changed: ' + name)
            return path
        if path.exists() and saved.get('status') != 'pending':
            raise ValueError('Untracked source file cannot be reused: ' + name)
        if not path.exists():
            self.data['files'][name] = {'status': 'pending'}
            self.save()
            collect(path)
        # A completed local export can precede a failed UI close. Recover it
        # from its recorded pending operation. The caller always reparses and
        # independently reconciles the contents before preparing final reports.
        if not path.is_file() or not path.stat().st_size:
            raise ValueError('The source operation produced no complete file')
        self.data['files'][name] = {'status': 'complete',
            'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
        self.save()
        return path

    @property
    def contacts(self):
        return self.data['contacts']

    def contact(self, key, value):
        self.data['contacts'][key] = value
        self.save()
