"""Install the validated exporter into a new directory; never replace an installation.

Without --install this prints a read-only installation plan. Close the dedicated
robot Chrome before installation so its 1C profile can be copied consistently.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

from nightly import load_runtime, prepared_files
from report_tools import read_config, inspect_xlsx, write_json, copy_checked, plan

SOURCE = Path(__file__).resolve().parent
CORE = ('nightly.py', 'onec_browser.py', 'onec_ui.py', 'source_extract.py',
        'drive_browser.py', 'drive_api.py', 'google_drive_auth.py', 'report_tools.py', 'source_cache.py', 'card_catalog.py', 'bases.json', 'README.md',
        'run_reports.cmd', 'install_schedule.ps1', 'requirements-lock.txt', 'manual-1c-export.html',
        'run_status.py', 'reconciliation.py', 'period_runner.py', 'program-guide.html', 'program-tests.html')


def require_verified_publications(cfg, runtime, report_date, prepared):
    """Check the publisher's completed receipts, including current API runs.

    This proves the packaged bytes were published successfully. It does not
    overwrite or adopt later manual edits to historical reports on Drive.
    """
    from drive_api import DriveAPI
    publisher = DriveAPI(None, cfg['drive_folder_id'], runtime['drive_state'])
    if cfg.get('drive_daily_folders'):
        if cfg.get('drive_month_folders'):
            month_name = report_date.strftime('%Y.%m_отчёты')
            month = json.loads(publisher.journal_path(month_name, folders=True).read_text(encoding='utf-8'))
            if (month.get('status') != 'verified' or month.get('name') != month_name
                    or month.get('parent_id') != cfg['drive_folder_id'] or not month.get('folder_id')):
                raise ValueError('The month folder has no verified publication receipt')
            publisher = DriveAPI(None, month['folder_id'], runtime['drive_state'], root_id=cfg['drive_folder_id'])
        name = report_date.strftime(cfg.get('drive_daily_folder_pattern', '%d.%m.%Y_отчёты'))
        path = publisher.journal_path(name, folders=True)
        folder = json.loads(path.read_text(encoding='utf-8'))
        if (folder.get('status') != 'verified' or folder.get('name') != name
                or folder.get('parent_id') != publisher.folder_id or not folder.get('folder_id')):
            raise ValueError('The daily folder has no verified publication receipt')
        publisher = DriveAPI(None, folder['folder_id'], runtime['drive_state'], root_id=cfg['drive_folder_id'])
    expected = {job['filename'] for job in plan(cfg, report_date)}
    packaged = [item for files in prepared.values() for item in files]
    if len(packaged) != len(expected) or {Path(item['path']).name for item in packaged} != expected:
        raise ValueError('The packaged report set is incomplete or contains duplicates')
    for item in packaged:
        name = Path(item['path']).name
        receipt = json.loads(publisher.journal_path(name).read_text(encoding='utf-8'))
        if (receipt.get('status') != 'verified' or receipt.get('filename') != name
                or receipt.get('folder_id') != publisher.folder_id or not receipt.get('file_id')
                or receipt.get('sha256') != item['sha256']):
            raise ValueError('The current local report version has no verified publication receipt')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--destination', type=Path, default=Path('C:/1c_export/automation'))
    parser.add_argument('--validated-date', required=True)
    parser.add_argument('--install', action='store_true')
    parser.add_argument('--browser-closed', action='store_true')
    args = parser.parse_args()
    destination = args.destination.resolve()
    if destination != Path('C:/1c_export/automation').resolve():
        raise ValueError('Installation is scoped to C:/1c_export/automation')
    cfg = read_config(SOURCE / 'bases.json')
    runtime = load_runtime(SOURCE / 'runtime.json')
    report_date = date.fromisoformat(args.validated_date)
    bases = [b for b in cfg['bases'] if b.get('enabled', True)]
    expected = plan(cfg, report_date)
    if len(bases) != 9 or len(expected) != 27:
        raise ValueError('Expected nine active bases and 27 daily reports')
    package = runtime['builder'].parent / 'node_modules/@oai/artifact-tool'
    launcher = destination.parent / 'run_reports.cmd'
    print(json.dumps({'destination': str(destination), 'launcher': str(launcher),
        'validated_date': report_date.isoformat(), 'daily_files': len(expected),
        'python': sys._base_executable, 'node_source': str(runtime['node']),
        'artifact_package_source': str(package), 'browser_profile_source': str(runtime['profile']),
        'creates_schedule': False}, ensure_ascii=False, indent=2))
    if not args.install:
        return 0
    if not cfg.get('production_enabled') or not args.browser_closed:
        raise ValueError('Production validation and a closed robot browser are required')
    if destination.exists() or launcher.exists():
        raise FileExistsError('Installation or launcher already exists; inspect it before any update')
    if not runtime['credentials_source'].is_file() or not package.is_dir():
        raise FileNotFoundError('A required local runtime source is missing')
    if cfg.get('drive_backend') == 'api' and not runtime['google_token'].is_file():
        raise FileNotFoundError('The verified Google Drive API connection is missing')
    prepared = {}
    for base in bases:
        prepared[base['id']] = prepared_files(runtime, base, report_date)
        if not prepared[base['id']]:
            raise ValueError('Validated local files are missing for ' + base['id'])
    require_verified_publications(cfg, runtime, report_date, prepared)

    destination.mkdir()
    write_json(destination / 'installation.json', {'status': 'installing'})
    for filename in CORE:
        shutil.copy2(SOURCE / filename, destination / filename)
    shutil.copy2(runtime['builder'], destination / 'build_reports.mjs')
    shutil.copytree(package, destination / 'node_modules/@oai/artifact-tool')
    (destination / 'runtime').mkdir()
    shutil.copy2(runtime['node'], destination / 'runtime/node.exe')
    subprocess.run([sys._base_executable, '-m', 'venv', str(destination / '.venv')], check=True)
    dependencies = Path(sys.prefix) / 'Lib/site-packages'
    shutil.copytree(dependencies, destination / '.venv/Lib/site-packages', dirs_exist_ok=True)
    shutil.copytree(runtime['profile'], destination / 'chrome_profile',
        ignore=shutil.ignore_patterns('Cache', 'Code Cache', 'GPUCache', 'ShaderCache',
                                     'GrShaderCache', 'DawnCache', 'Crashpad', 'SingletonLock',
                                     'SingletonCookie', 'SingletonSocket'))
    shutil.copytree(runtime['selenium_cache'], destination / 'selenium_cache')
    shutil.copytree(runtime['drive_state'], destination / 'drive-state')
    if cfg.get('drive_backend') == 'api':
        (destination / 'google-auth').mkdir()
        shutil.copy2(runtime['google_token'], destination / 'google-auth/token.json')
        client = runtime['google_token'].parent / 'client.json'
        if client.is_file():
            shutil.copy2(client, destination / 'google-auth/client.json')
    for folder in ('downloads', 'runs', 'reports', 'state'):
        (destination / folder).mkdir()
    if (runtime['state'] / 'card-catalog').is_dir():
        shutil.copytree(runtime['state'] / 'card-catalog', destination / 'state/card-catalog',
                        ignore=shutil.ignore_patterns('*.inspect.ndjson', 'builder*.log',
                                                     'catalog.json', 'manifest.json', 'pending.xlsx'))
    # Keep independent validated staging copies. Recovery no longer needs files
    # in the development workspace, even for the initial day's reports.
    for base in bases:
        files = []
        for item in prepared[base['id']]:
            path = Path(item['path'])
            staged = destination / 'prepared' / report_date.isoformat() / base['id'] / path.name
            copy_checked(path, staged)
            final = destination / 'reports' / path.name
            checked = copy_checked(staged, final)
            files.append(dict(checked, staged_path=str(staged)))
        write_json(destination / 'state' / report_date.isoformat() / (base['id'] + '.json'), {
            'date': report_date.isoformat(), 'base_id': base['id'], 'files': files})
    from onec_browser import load_credentials
    credentials = load_credentials(runtime['credentials_source'])
    (destination / 'credentials.py').write_text('LOGIN = ' + repr(credentials.login) + '\nPASSWORD = ' + repr(credentials.password) + '\n', encoding='utf-8')
    write_json(destination / 'runtime.json', {
        'profile': 'chrome_profile', 'downloads': 'downloads', 'selenium_cache': 'selenium_cache',
        'google_token': 'google-auth/token.json',
        'credentials_source': 'credentials.py', 'state': 'state',
        'drive_state': 'drive-state', 'runs': 'runs', 'reports': 'reports',
        'node': 'runtime/node.exe', 'builder': 'build_reports.mjs'})
    python = destination / '.venv/Scripts/python.exe'
    subprocess.run([str(python), '-X', 'utf8', str(destination / 'nightly.py'), '--plan',
                    '--date', report_date.isoformat()], check=True, capture_output=True)
    subprocess.run([str(destination / 'runtime/node.exe'), '--input-type=module', '-e',
                    "import {Workbook} from '@oai/artifact-tool'; Workbook.create();"],
                   cwd=destination, check=True, capture_output=True)
    launcher.write_text('@echo off\ncall "%~dp0automation\\run_reports.cmd" %*\nexit /b %ERRORLEVEL%\n',
                        encoding='ascii')
    write_json(destination / 'installation.json', {'status': 'installed',
        'validated_date': report_date.isoformat(), 'validated_report_count': 27,
        'files': {name: __import__('hashlib').sha256((destination / name).read_bytes()).hexdigest()
                  for name in CORE}})
    print('INSTALLED', str(destination), flush=True)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
