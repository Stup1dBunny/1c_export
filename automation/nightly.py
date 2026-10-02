"""Standalone 1C export runner for the previous Moscow calendar day.

Run --pilot --base vda_vin_retz for an explicit single-base run. A cached complete
local set is reused on retries; pending uploads are reconciled by the configured
Drive publisher (official API by default).
"""
from __future__ import annotations

import argparse
import json
import logging
import subprocess
import sys
import time
from contextlib import contextmanager, nullcontext
from datetime import date, datetime
from pathlib import Path

from onec_browser import OneCBrowser, load_credentials
from onec_ui import OneCUI
from drive_browser import DriveBrowser
from source_cache import SourceCache
from reconciliation import DiscrepancyLedger, reconcile_sources, describe_detail
from card_catalog import CardCatalogCache, publish_catalog
from run_status import RunStatus, CHECKS, error_reason
from period_runner import month_end
from report_tools import (ROOT, MSK, read_config, report_day, plan, inspect_xlsx,
                          write_json, upload_lock, copy_checked, replace_checked, period_key, period_label)
from source_extract import (read_first_sheet, extract_receipts, extract_stock,
                            extract_alcohol, extract_check_list,
                            prepare_sales, json_numbers)

REPORT_SCHEMA_VERSION = 2


@contextmanager
def daily_logging(state, day, end_day=None):
    logger = logging.getLogger()
    previous, level = logger.handlers[:], logger.level
    handler = logging.FileHandler(state / (period_key(day, end_day) + '.log'), encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    for old in previous:
        logger.removeHandler(old)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    try:
        yield
    finally:
        logger.removeHandler(handler)
        handler.close()
        for old in previous:
            logger.addHandler(old)
        logger.setLevel(level)


def current_report_schema(runtime, base, day, end_day=None):
    marker = runtime['state'] / period_key(day, end_day) / (base['id'] + '.json')
    return marker.exists() and json.loads(marker.read_text(encoding='utf-8')).get('report_schema_version') == REPORT_SCHEMA_VERSION


def load_runtime(path):
    return {k: (path.parent / v).resolve() for k, v in
            json.loads(path.read_text(encoding='utf-8-sig')).items()}


def refresh_for_observation(observation, run_id):
    if not observation:
        return None
    if observation.get('publication_pending'):
        folder = Path(observation['generation']).name
        if folder.startswith('recheck-'):
            return folder[len('recheck-'):]
        return None
    if observation.get('open_count'):
        return run_id
    return None


def prepared_files(runtime, base, day, end_day=None):
    marker = runtime['state'] / period_key(day, end_day) / (base['id'] + '.json')
    if not marker.exists():
        return None
    saved = json.loads(marker.read_text(encoding='utf-8'))
    if (saved['date'] != day.isoformat() or saved.get('end_date') != (end_day.isoformat() if end_day else None)
            or saved['base_id'] != base['id'] or len(saved['files']) != 3):
        raise ValueError('Invalid prepared report manifest')
    for item in saved['files']:
        target = Path(item['path'])
        if not target.exists() and item.get('staged_path'):
            staged = Path(item['staged_path'])
            if inspect_xlsx(staged)['sha256'] != item['sha256']:
                raise ValueError('Staged report was changed after validation')
            copy_checked(staged, target)
        checked = inspect_xlsx(target)
        if checked['sha256'] != item['sha256'] and checked['sha256'] in item.get('replace_previous_shas', []):
            staged = Path(item['staged_path'])
            if inspect_xlsx(staged)['sha256'] != item['sha256']:
                raise ValueError('Staged report was changed after validation')
            checked = replace_checked(staged, target, checked['sha256'], runtime['state'] / 'report-history')
        if checked['sha256'] != item['sha256']:
            raise ValueError('A prepared report was changed after validation')
    return saved['files']


def prepare_base(browser, runtime, base, cfg, day, refresh_id=None, progress_event=None, end_day=None):
    notify = progress_event or (lambda stage: None)
    cached = prepared_files(runtime, base, day, end_day)
    marker = runtime['state'] / period_key(day, end_day) / (base['id'] + '.json')
    saved = json.loads(marker.read_text(encoding='utf-8')) if marker.exists() else {}
    if cached and saved.get('report_schema_version') == REPORT_SCHEMA_VERSION and (not refresh_id or saved.get('refresh_id') == refresh_id):
        return cached
    jobs = [j for j in plan(cfg, day, end_day) if j['base_id'] == base['id']]
    if not cached and any((runtime['reports'] / j['filename']).exists() for j in jobs):
        raise ValueError('Existing local reports have no validation manifest; refusing to overwrite')
    links = base['report_links']
    folder = runtime['runs'] / period_key(day, end_day) / base['id'] / (('recheck-' + refresh_id) if refresh_id else 'resumable')
    if cached and not refresh_id and saved.get('source_folder'):
        # Upgrade column layout from the most recent validated source set,
        # including a previous reconciliation, instead of reverting to older sources.
        folder = Path(saved['source_folder'])
    progress = SourceCache(folder, {'schema': 1, 'base_id': base['id'], 'date': day.isoformat(),
        **({'end_date': end_day.isoformat()} if end_day else {}),
        'report_links': links, 'supplier_missing': base.get('stock_supplier_missing') is True,
        'revenue_source': cfg.get('revenue_source', 'strict_match')})
    ui = OneCUI(browser.driver, runtime['downloads'], card_catalog_format=base.get('card_catalog_format', 'text'))
    results = {}
    for kind, extractor in [('receipts', extract_receipts), ('stock', extract_stock), ('alcohol', extract_alcohol)]:
        notify('Получение и проверка: ' + {'receipts': 'сводный отчёт по чекам', 'stock': 'остатки', 'alcohol': 'товарные продажи'}[kind])
        stage_started = time.monotonic()
        logging.info('base=%s date=%s collecting=%s', base['id'], day, kind)
        path = progress.source(kind + '.xlsx',
                               lambda target: ui.collect_report(links[kind + '_variant'], day, target,
                                                                **({'end_day': end_day} if end_day else {})))
        rows, provenance = read_first_sheet(path)
        options = {'supplier_missing': base.get('stock_supplier_missing') is True} if kind == 'stock' else {}
        results[kind] = extractor(rows, **options)
        if kind == 'stock' and results[kind]['supplier_source'] == 'absent_in_base':
            logging.warning('base=%s date=%s supplier field absent in base; output column kept blank',
                            base['id'], day)
        write_json(folder / (kind + '.json'), json_numbers(dict(source=provenance, **results[kind])))
        logging.info('base=%s date=%s stage=%s elapsed_seconds=%.1f',
                     base['id'], day, kind, time.monotonic() - stage_started)
    stage_started = time.monotonic()
    notify('Получение полного журнала чеков и сверка источников')
    checks_path = progress.source('checks.tsv',
        lambda target: ui.collect_check_list(links['checks_list'], day, target, **({'end_day': end_day} if end_day else {})))
    checks = checks_path.read_text(encoding='utf-8')
    ledger = None
    if cfg.get('discrepancy_policy') == 'preliminary':
        sales, bonuses, issues = reconcile_sources(results['receipts'], results['alcohol'], checks, day.isoformat(),
                                                   end_day.isoformat() if end_day else None)
        ledger = DiscrepancyLedger(runtime['state'])
    else:
        sales = prepare_sales(results['receipts'], results['alcohol'],
                              revenue_source=cfg.get('revenue_source', 'strict_match'))
        write_json(folder / 'sales-reconciliation.json', json_numbers(sales['reconciliation']))
        bonuses = extract_check_list(checks, day.isoformat(), results['receipts'], end_date=end_day.isoformat() if end_day else None)
    logging.info('base=%s date=%s stage=check_journal elapsed_seconds=%.1f',
                 base['id'], day, time.monotonic() - stage_started)
    logging.info('base=%s date=%s check list exclusions=%s',
                 base['id'], day, bonuses['excluded_records'])
    if bonuses.get('archived_without_fiscal_fields'):
        logging.warning('base=%s date=%s reconciled_archived_without_fiscal_fields=%s',
                        base['id'], day, bonuses['archived_without_fiscal_fields'])
    stage_started = time.monotonic()
    catalog = CardCatalogCache(runtime['state'] / 'card-catalog', base['id'], links['cards_list'])
    notify('Сопоставление покупок со справочником карт')
    def collect_cards(target):
        notify('1С формирует полный справочник карт; ожидание готового списка')
        return ui.collect_card_list(links['cards_list'], target)
    bonuses = catalog.enrich(bonuses, collect_cards)
    if ledger:
        observation = ledger.observe(base, day, issues + bonuses.get('card_resolution_issues', []),
                                     str(folder), end_day, analysis_version=2)
        write_json(folder / 'sales-reconciliation.json', observation)
    if bonuses.get('card_resolution_issues'):
        logging.warning('base=%s date=%s ambiguous_card_rows=%s; source names and purchases retained without guessing card numbers',
                        base['id'], day, bonuses['ambiguous_card_rows'])
    if bonuses.get('updated_owner_card_count'):
        logging.info('base=%s date=%s owner_names_from_fresh_catalog=%s; joined by card number',
                     base['id'], day, bonuses['updated_owner_card_count'])
    if bonuses.get('unnumbered_card_rows'):
        logging.warning('base=%s date=%s unnumbered_card_rows=%s; named purchases retained with an explicit missing-number label',
                        base['id'], day, bonuses['unnumbered_card_rows'])
    if bonuses.get('ambiguous_owner_card_count'):
        logging.warning('base=%s date=%s ambiguous_owner_cards=%s; purchases retained by card number',
                        base['id'], day, bonuses['ambiguous_owner_card_count'])
    if bonuses.get('missing_card_reference_rows'):
        logging.warning('base=%s date=%s missing_card_reference_rows=%s; purchases retained with blank contacts',
                        base['id'], day, bonuses['missing_card_reference_rows'])
    logging.info('base=%s date=%s stage=bulk_cards statistics=%s elapsed_seconds=%.1f',
                 base['id'], day, bonuses['card_catalog_statistics'], time.monotonic() - stage_started)
    write_json(folder / 'sales.json', json_numbers(sales))
    write_json(folder / 'bonus.json', json_numbers(bonuses))
    staged_reports = folder / 'reports-staged'
    manifest = {'output_dir': str(staged_reports), 'reports': [
        {'type': kind, 'base': base['name'], 'date': day.isoformat(), 'data': source,
         **({'end_date': end_day.isoformat()} if end_day else {})}
        for kind, source in [('Продажи', 'sales.json'), ('Остатки', 'stock.json'), ('Бонусы', 'bonus.json')]]}
    if ledger and observation['issues']:
        control_rows = [{'state': 'Открыто' if i['status'] == 'open' else 'Устранено', 'shop': i['shop'],
                         'message': i['message'], 'detail': describe_detail(i['detail']),
                         'first_seen': i['first_seen'], 'last_checked': i['last_checked']}
                        for i in observation['issues']]
        write_json(folder / 'control.json', {'rows': control_rows})
        manifest['reports'].append({'type': 'Контроль', 'base': base['name'], 'date': day.isoformat(), 'data': 'control.json',
                                   **({'end_date': end_day.isoformat()} if end_day else {})})
    write_json(folder / 'manifest.json', manifest)
    notify('Создание и проверка итоговых XLSX')
    stage_started = time.monotonic()
    completed = subprocess.run([str(runtime['node']), str(runtime['builder']), str(folder / 'manifest.json')],
                               capture_output=True, text=True, encoding='utf-8', timeout=900 if end_day else 300)
    (folder / 'builder.log').write_text(completed.stdout + completed.stderr, encoding='utf-8')
    if completed.returncode:
        raise RuntimeError('Workbook builder failed; see the run log')
    staged_files = [inspect_xlsx(staged_reports / j['filename']) for j in jobs]
    logging.info('base=%s date=%s stage=workbooks elapsed_seconds=%.1f',
                 base['id'], day, time.monotonic() - stage_started)
    files = []
    qualities = {'Продажи': sales.get('quality', 'verified'), 'Остатки': 'verified',
                 'Бонусы': bonuses.get('quality', 'verified')}
    for item, job in zip(staged_files, jobs):
        prior = next((p for p in (cached or []) if Path(p['path']).name == job['filename']), None)
        files.append(dict(item, staged_path=item['path'], path=str(runtime['reports'] / Path(item['path']).name),
                          quality=qualities[job['report_type']],
                          replace_previous_shas=([prior['sha256']] + prior.get('replace_previous_shas', [])) if prior else []))
    # Record all validated source copies before publishing any final local
    # filename. A restart can complete the set without rebuilding ZIP files.
    write_json(runtime['state'] / period_key(day, end_day) / (base['id'] + '.json'), {
        'report_schema_version': REPORT_SCHEMA_VERSION,
        'date': day.isoformat(), 'base_id': base['id'], 'source_folder': str(folder), 'files': files,
        **({'end_date': end_day.isoformat()} if end_day else {}),
        'refresh_id': refresh_id,
        'control': inspect_xlsx(staged_reports / f'{period_label(day, end_day)}_{base["name"]}_Контроль.xlsx')
                   if ledger and observation['issues'] else None})
    return prepared_files(runtime, base, day, end_day)


def transient_failure(error):
    from google.auth.exceptions import TransportError
    from googleapiclient.errors import HttpError
    from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException, InvalidSessionIdException
    from requests.exceptions import ConnectionError as RequestConnectionError, Timeout as RequestTimeout
    from urllib3.exceptions import MaxRetryError, ReadTimeoutError, NewConnectionError
    if isinstance(error, MaxRetryError):
        return isinstance(error.reason, (ReadTimeoutError, NewConnectionError))
    if isinstance(error, UnexpectedAlertPresentException):
        return (error.alert_text or '').strip().startswith('Потеряно соединение.')
    if isinstance(error, HttpError):
        return error.resp.status in {429, 500, 502, 503, 504}
    return isinstance(error, (TimeoutError, ConnectionError, TransportError,
        RequestConnectionError, RequestTimeout, TimeoutException, InvalidSessionIdException, ReadTimeoutError))


def retry_operation(operation, on_failure, attempts=3, sleep=time.sleep):
    """Retry temporary transport/UI failures, never data or authorization errors."""
    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except Exception as error:
            retry = attempt < attempts and transient_failure(error)
            on_failure(error, attempt, retry)
            if not retry:
                raise
            sleep(10 * attempt)


def run_base_attempt(runtime, base, cfg, day, credentials, local_only=False, status=None, refresh_id=None, end_day=None):
    """Each retry gets a clean browser; verified sources/contacts stay on disk."""
    backend = cfg.get('drive_backend', 'browser')
    browser = None
    def event(stage, **fields):
        if status is not None:
            status.event(base['id'], stage, **fields)
    try:
        event('Проверка подключения и сохранённых файлов', status='running', checks=[], files=[])
        root_publisher = None
        publisher = None
        if backend == 'api' and not local_only:
            from google_drive_auth import build_service
            from drive_api import DriveAPI
            root_publisher = DriveAPI(build_service(runtime['google_token']), cfg['drive_folder_id'], runtime['drive_state'])
            root_publisher.verify_folder()
            publisher = root_publisher.for_day(cfg, day, **({'end_day': end_day} if end_day else {}))
        files = prepared_files(runtime, base, day, end_day)
        if refresh_id or (files is not None and 'state' in runtime and not current_report_schema(runtime, base, day, end_day)):
            files = None
        if files is not None:
            event('Используется ранее проверенный комплект', checks=[
                'Три ранее подготовленных XLSX прочитаны; их контрольные суммы совпали. Данные 1С заново не запрашивались.'])
        catalog = None
        catalog_due = False
        if cfg.get('card_catalog_enabled'):
            catalog = CardCatalogCache(runtime['state'] / 'card-catalog', base['id'], base['report_links']['cards_list'])
            catalog_due = catalog.load() is None
        if files is None or catalog_due or (backend == 'browser' and not local_only):
            browser = OneCBrowser(runtime['profile'], runtime['downloads'], runtime['selenium_cache'])
        if files is None or catalog_due:
            event('Вход в 1С; получение отчётов' if files is None else 'Обновление полного справочника карт')
            browser.login(base['url'], credentials)
        if catalog_due and files is not None:
            ui = OneCUI(browser.driver, runtime['downloads'], card_catalog_format=base.get('card_catalog_format', 'text'))
            catalog.refresh(lambda target: ui.collect_card_list(base['report_links']['cards_list'], target))
        if files is None:
            files = prepare_base(browser, runtime, base, cfg, day, refresh_id, event, end_day)
            preliminary = any(i.get('quality') == 'preliminary' for i in files)
            event('Три XLSX готовы; есть предварительные данные' if preliminary else 'Проверки данных пройдены; три XLSX готовы',
                  checks=([ 'Все пробитые чеки включены. Несверенные строки помечены как предварительные; причины записаны на листе «Проверка».'] + CHECKS[1:]) if preliminary else CHECKS.copy())
        if not local_only:
            event('Публикация и проверка файлов на Google Drive')
            started = time.monotonic()
            if backend == 'browser':
                browser.driver.switch_to.new_window('tab')
                root_publisher = DriveBrowser(browser.driver, runtime['downloads'], cfg['drive_folder_id'], runtime['drive_state'])
            if publisher is None:
                publisher = root_publisher.for_day(cfg, day, **({'end_day': end_day} if end_day else {}))
            for item in files:
                options = {}
                if item.get('replace_previous_shas'):
                    if backend != 'api':
                        raise ValueError('Reconciled report replacement requires Drive API')
                    journal = publisher.journal_path(Path(item['path']).name)
                    previous = json.loads(journal.read_text(encoding='utf-8')) if journal.exists() else {}
                    if previous.get('sha256') in item['replace_previous_shas']:
                        options['replace_previous_sha'] = previous['sha256']
                result = publisher.publish(Path(item['path']), **options)
                if status is not None:
                    row = status.row(base['id'])
                    status.event(base['id'], 'Подтверждён файл на Drive: ' + Path(item['path']).name,
                                 files=row['files'] + [{'filename': Path(item['path']).name, 'status': result['status'],
                                                       'quality': item.get('quality', 'verified')}])
                logging.info('base=%s file=%s status=%s', base['id'], Path(item['path']).name, result['status'])
            if cfg.get('discrepancy_policy') == 'preliminary':
                marker = runtime['state'] / period_key(day, end_day) / (base['id'] + '.json')
                control = json.loads(marker.read_text(encoding='utf-8')).get('control')
                if control:
                    if inspect_xlsx(Path(control['path']))['sha256'] != control['sha256']:
                        raise ValueError('Control report changed after validation')
                    controls = root_publisher.ensure_subfolder('Контроль выгрузки')
                    publish_catalog(controls, Path(control['path']))
            if catalog:
                if backend != 'api':
                    raise ValueError('Weekly card catalogs require the configured Drive API backend')
                catalog_file = catalog.prepare_workbook(runtime, base)
                catalogs = root_publisher.ensure_subfolder(cfg.get('card_catalog_folder', 'Справочники карт'))
                result = publish_catalog(catalogs, catalog_file)
                logging.info('base=%s catalog=%s status=%s', base['id'], catalog_file.name, result['status'])
            logging.info('base=%s date=%s stage=drive elapsed_seconds=%.1f', base['id'], day, time.monotonic() - started)
        elif catalog:
            catalog.prepare_workbook(runtime, base)
        if not local_only and cfg.get('discrepancy_policy') == 'preliminary':
            DiscrepancyLedger(runtime['state']).published(base['id'], day, end_day)
        return files
    except Exception:
        logging.exception('base=%s export attempt failed before diagnostics', base['id'])
        if browser is not None:
            try:
                browser.driver.command_executor.client_config.timeout = 3
                diagnostics = runtime['state'] / 'diagnostics' / period_key(day, end_day)
                diagnostics.mkdir(parents=True, exist_ok=True)
                browser.driver.save_screenshot(str(diagnostics / (base['id'] + '.png')))
            except Exception:
                logging.warning('base=%s could not capture the failed browser state', base['id'])
        raise
    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception:
                logging.warning('base=%s browser close failed', base['id'])


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=ROOT / 'bases.json')
    parser.add_argument('--runtime', type=Path, default=ROOT / 'runtime.json')
    parser.add_argument('--date', help='YYYY-MM-DD; default is yesterday in Moscow')
    parser.add_argument('--base', action='append', default=[])
    parser.add_argument('--pilot', action='store_true')
    parser.add_argument('--plan', action='store_true', help='Read-only plan; do not launch Chrome')
    parser.add_argument('--local-only', action='store_true')
    parser.add_argument('--show-report', action='store_true', help='Open the local result after a manual run')
    parser.add_argument('period', nargs='?', choices=['month'], help='One export for a full month starting at --date; three reports per base')
    args = parser.parse_args(argv)
    if args.period and not args.date:
        parser.error('Для month укажите дату начала: --date YYYY-MM-DD month')
    cfg, runtime, day = read_config(args.config), load_runtime(args.runtime), report_day(args.date)
    end_day = month_end(day) if args.period == 'month' else None
    if day >= datetime.now(MSK).date():
        raise ValueError('Only a completed calendar day can be exported')
    selected = [b for b in cfg['bases'] if b.get('enabled', True) and (not args.base or b['id'] in args.base)]
    if not selected:
        raise ValueError('Не выбрано ни одной активной базы')
    if args.base and set(args.base) != {b['id'] for b in selected}:
        raise ValueError('Unknown or disabled base id')
    if args.plan:
        jobs = [j for j in plan(cfg, day, end_day) if j['base_id'] in {b['id'] for b in selected}]
        preview = jobs if not args.period else {
            'mode': 'month', 'start_date': day.isoformat(), 'end_date': end_day.isoformat(),
            'day_count': (end_day - day).days + 1, 'file_count': len(jobs),
            'bases': [{'id': b['id'], 'name': b['name']} for b in selected],
            'reports': jobs}
        print(json.dumps(preview, ensure_ascii=False, indent=2))
        return 0
    if not cfg.get('production_enabled') and not args.pilot:
        raise ValueError('Nightly production is not enabled: all bases must first pass validation')
    if args.pilot and (len(selected) != 1 or not args.base):
        raise ValueError('Pilot mode requires exactly one explicit --base')
    if any(not b.get('report_links') for b in selected):
        raise ValueError('Some selected bases do not yet have verified report navigation links')
    code, _ = run_day(args, cfg, runtime, day, selected, **({'end_day': end_day} if end_day else {}))
    if args.show_report:
        __import__('webbrowser').open((runtime['state'] / ('last-period.html' if end_day else 'last-run.html')).as_uri())
    return code


def run_day(args, cfg, runtime, day, selected, *, lock_held=False, recheck_before=None, end_day=None):
    runtime['state'].mkdir(parents=True, exist_ok=True)
    failures, timings = [], []
    run_started = time.monotonic()
    with (nullcontext() if lock_held else upload_lock(runtime['state'] / 'runner')), daily_logging(runtime['state'], day, end_day):
        status = RunStatus(runtime['state'], day, selected, args.local_only, end_day=end_day)
        print(f'Выгрузка за {period_label(day, end_day)}: {len(selected)} баз, {len(selected)*3} отчётов.', flush=True)
        try:
            credentials = load_credentials(runtime['credentials_source'])
        except Exception as exc:
            logging.exception('Could not load 1C credentials')
            status.finish('failed', [{'base': 'Начальная проверка', 'error_type': type(exc).__name__}], [])
            return 1, status
        backend = cfg.get('drive_backend', 'browser')
        if backend not in {'browser', 'api'}:
            raise ValueError('Unknown Drive publication backend')
        ledger = DiscrepancyLedger(runtime['state'])
        pending = {(s['base_id'], s['date']): s for s in ledger.pending()}
        interrupted = False
        for base in selected:
            base_started = time.monotonic()
            base_status = 'failed'
            try:
                def on_failure(error, attempt, retry):
                    logging.warning('base=%s attempt=%s retry=%s error=%s',
                                    base['id'], attempt, retry, type(error).__name__)
                    if retry:
                        status.event(base['id'], f'Временный сбой; повтор после попытки {attempt}')
                previous = ledger.load(base['id'], day, end_day) if end_day else pending.get((base['id'], day.isoformat()))
                refresh_id = refresh_for_observation(previous, status.stem)
                files = retry_operation(lambda: run_base_attempt(runtime, base, cfg, day, credentials, args.local_only,
                                        status, refresh_id, **({'end_day': end_day} if end_day else {})), on_failure)
                logging.info('base=%s date=%s complete', base['id'], day)
                preliminary = any(f.get('quality') == 'preliminary' for f in files)
                base_status = 'preliminary' if preliminary else 'complete'
                status.event(base['id'], 'Файлы готовы; есть предварительные отчёты, см. контроль расхождений' if preliminary else
                             ('Готово локально' if args.local_only else 'Отчёты и справочник проверены на Drive'),
                             status='preliminary' if preliminary else ('local' if args.local_only else 'complete'))
                print(base['id'], day.isoformat(), 'OK', flush=True)
            except Exception as exc:
                failures.append({'base': base['id'], 'error_type': type(exc).__name__, 'error_reason': error_reason(exc)})
                detail = str(exc) if isinstance(exc, (ValueError, FileExistsError, RuntimeError, ConnectionError)) else ''
                logging.error('base=%s error=%s %s', base['id'], type(exc).__name__, detail, exc_info=True)
                print(base['id'], day.isoformat(), 'FAILED', type(exc).__name__, flush=True)
                status.event(base['id'], 'Обработка базы не завершена', status='failed', error_type=type(exc).__name__,
                             error_reason=error_reason(exc))
            except KeyboardInterrupt:
                interrupted = True
                failures.append({'base': base['id'], 'error_type': 'KeyboardInterrupt'})
                status.event(base['id'], 'Остановлено пользователем', status='failed', error_type='KeyboardInterrupt')
                break
            finally:
                elapsed = round(time.monotonic() - base_started, 1)
                timings.append({'base': base['id'], 'status': base_status, 'elapsed_seconds': elapsed})
                logging.info('base=%s date=%s status=%s elapsed_seconds=%.1f',
                             base['id'], day, base_status, elapsed)
        # Daily output is handled first. Recheck the ORIGINAL days still in
        # the ledger, including publication interrupted after reconciliation.
        # One source generation per invocation is shared by transport retries.
        for (base_id, original_day), observation in ([] if interrupted or end_day else sorted(pending.items())):
            if original_day >= (recheck_before or day).isoformat() or base_id not in {b['id'] for b in selected}:
                continue
            base = next(b for b in selected if b['id'] == base_id)
            old_day = date.fromisoformat(original_day)
            status.data.setdefault('rechecks', []).append({'base': base_id, 'date': original_day, 'status': 'running'})
            record = status.data['rechecks'][-1]
            status.save()
            print(f'Повторная сверка {base["name"]} за {original_day}', flush=True)
            try:
                refresh_id = refresh_for_observation(observation, status.stem)
                files = retry_operation(lambda: run_base_attempt(runtime, base, cfg, old_day, credentials,
                    args.local_only, refresh_id=refresh_id), lambda e, a, r: logging.warning(
                        'recheck base=%s date=%s attempt=%s retry=%s error=%s', base_id, old_day, a, r, type(e).__name__))
                record['status'] = 'preliminary' if any(f.get('quality') == 'preliminary' for f in files) else 'complete'
            except Exception as exc:
                record['status'] = 'failed'
                record['error_type'] = type(exc).__name__
                failures.append({'base': base_id + ' за ' + original_day, 'error_type': type(exc).__name__, 'error_reason': error_reason(exc)})
                logging.exception('Historical reconciliation failed base=%s date=%s', base_id, original_day)
            except KeyboardInterrupt:
                record['status'] = 'failed'
                record['error_type'] = 'KeyboardInterrupt'
                failures.append({'base': base_id + ' за ' + original_day, 'error_type': 'KeyboardInterrupt'})
                interrupted = True
            status.save()
            if interrupted:
                break
        warnings = any(r['status'] == 'preliminary' for r in status.data['results'] + status.data.get('rechecks', []))
        status.finish('failed' if failures else ('preliminary' if warnings else 'complete'), failures, timings)
    return (130 if interrupted else (1 if failures else 0)), status


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except (ValueError, FileNotFoundError) as exc:
        print(str(exc), file=sys.stderr)
        raise SystemExit(2)
