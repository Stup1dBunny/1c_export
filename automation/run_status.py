"""Local, human-readable run evidence. No credentials or customer rows."""
from __future__ import annotations

from datetime import datetime
from html import escape
from pathlib import Path

from report_tools import MSK, write_json

STATUS = {'waiting': 'Ожидает', 'running': 'В работе', 'complete': 'Готово',
          'failed': 'Ошибка', 'local': 'Готово локально', 'preliminary': 'Есть предварительные отчёты'}
FILE_STATUS = {'verified': 'Загружен и проверен', 'unchanged': 'Уже был на Drive; проверен',
               'updated': 'Обновлён и проверен'}
CHECKS = [
    'Сопоставимые наборы чеков сверены со сводным отчётом. «Пробитый» и «Архивный» включены в ежедневные суммы.',
    'Начальный остаток + приход − расход = конечный остаток; проверены строки и итоги.',
    'Карты сопоставлены с полным справочником; неоднозначные имена не подставляются наугад.',
    'Три итоговых XLSX прочитаны и проверены; записаны контрольные суммы файлов.',
]


REASONS = {
    'unknown_check_status': 'В журнале есть документ с неизвестным статусом. Требуется определить, состоялась ли продажа; автоматически включать или исключать его нельзя.',
    'missing_fiscal_fields': 'В журнале есть чек без обязательных реквизитов: номера, смены или подтверждающей ссылки. Обработка остановлена при проверке журнала 1С, до загрузки файлов на Drive.',
    'unconfirmed_archived_refund': 'Архивный возврат без номера чека и ссылки на отчёт о розничных продажах не удалось подтвердить сводным отчётом. Требуется сверка количества продаж и возвратов и их сумм.',
    'drive_folder_trashed': 'Папка этого дня находится в корзине Google Drive. Восстановите её; если удалили окончательно, требуется создать новую папку и обновить привязку программы.',
    'drive_folder_missing': 'Ранее сохранённая папка дня больше не доступна на Google Drive. После окончательного удаления требуется создать новую папку и обновить привязку; обычный повтор этого не делает.',
    'drive_folder_changed': 'Папку дня переместили или заменили другой. Проверьте её расположение и привязку программы.',
    'shop_scope_mismatch': 'В отчёте по чекам и товарном отчёте различаются магазины. Комплект остановлен до сверки источников.',
    'journal_scope_mismatch': 'В полном журнале чеков и сводном отчёте различаются магазины. Требуется сверка источников.',
    'journal_totals_mismatch': 'Суммы или количество чеков в полном журнале не совпали со сводным отчётом.',
}


def error_reason(error):
    # Only known, fixed messages are public. Do not copy arbitrary exception
    # strings, URLs, customer data or credentials into user-facing reports.
    if str(error).startswith('Check status needs explicit validation: '):
        return 'unknown_check_status'
    return {
        'Archived check has no fiscal number or shift': 'missing_fiscal_fields',
        'Issued check has no fiscal number or shift': 'missing_fiscal_fields',
        'Archived refund without fiscal references does not reconcile with the receipt report': 'unconfirmed_archived_refund',
        'The daily report folder is in Drive trash': 'drive_folder_trashed',
        'The previously verified daily folder is missing': 'drive_folder_missing',
        'The daily folder identity or location changed': 'drive_folder_changed',
        'Receipt and sales reports contain different shops': 'shop_scope_mismatch',
        'Check list shops differ from receipt report': 'journal_scope_mismatch',
        'Check list does not reconcile with the receipt report': 'journal_totals_mismatch',
    }.get(str(error))


def error_hint(kind, reason=None):
    if reason in REASONS:
        return REASONS[reason]
    return {
        'PermissionError': 'Проверьте учётные данные 1С и действительность подключения Google.',
        'FileNotFoundError': 'Не найден обязательный файл программы или настроек.',
        'CardConflict': 'Сведения о карте в журнале и полном справочнике противоречат друг другу. Требуется проверка сопоставления; покупки не привязываются к случайной карте.',
        'CardNotFound': 'Карта из журнала не найдена в полном справочнике даже после его обновления. Проверьте источник и доступность записи карты.',
        'FileExistsError': 'Обнаружен уже существующий файл или незавершённая операция; проверьте журнал.',
        'ValueError': 'Не прошла проверка данных, файла или его прежней версии на Drive. Подробности в журнале.',
        'TimeoutError': 'Превышено время ожидания: проверьте 1С, сеть и загрузки Chrome.',
        'TimeoutException': '1С не ответила вовремя; проверьте окно браузера и сеть.',
        'SessionNotCreatedException': 'Chrome робота не смог запуститься. Закройте оставшееся окно Chrome робота после прежнего запуска и повторите команду.',
        'ConnectionError': 'Не удалось связаться с сервером. Повторите после восстановления связи.',
        'KeyboardInterrupt': 'Запуск остановлен пользователем.',
    }.get(kind, 'Смотрите журнал запуска; после устранения причины повторите ту же дату.')


class RunStatus:
    def __init__(self, state, day, bases, local_only=False, end_day=None):
        self.state = Path(state)
        self.started = datetime.now(MSK)
        self.stem = self.started.strftime('%Y%m%d-%H%M%S-%f')
        self.data = {'date': day.isoformat(), 'end_date': end_day.isoformat() if end_day else None,
                     'started_at': self.started.isoformat(),
                     'finished_at': None, 'status': 'running', 'local_only': local_only,
                     'bases': [b['id'] for b in bases], 'failures': [], 'timings': [],
                     'elapsed_seconds': 0, 'results': [
                         {'base': b['id'], 'name': b['name'], 'status': 'waiting',
                          'stage': 'Ожидание запуска', 'checks': [], 'files': []} for b in bases]}
        self.save()

    def row(self, base_id):
        return next(r for r in self.data['results'] if r['base'] == base_id)

    def event(self, base_id, stage, **fields):
        row = self.row(base_id)
        row.update(stage=stage, **fields)
        self.save()
        print(f"{row['name']}: {stage}", flush=True)

    def save(self):
        self.data['elapsed_seconds'] = round((datetime.now(MSK) - self.started).total_seconds(), 1)
        history = self.state / 'run-reports'
        history.mkdir(parents=True, exist_ok=True)
        write_json(history / (self.stem + '.json'), self.data)
        write_json(self.state / 'last-run.json', self.data)
        html = self.render()
        paths = [history / (self.stem + '.html'), self.state / 'last-run.html']
        if self.data.get('end_date'):
            write_json(self.state / 'last-period.json', self.data)
            paths.append(self.state / 'last-period.html')
        for path in paths:
            pending = path.with_suffix('.html.tmp')
            pending.write_text(html, encoding='utf-8')
            pending.replace(path)

    def finish(self, status, failures, timings):
        self.data.update(status=status, finished_at=datetime.now(MSK).isoformat(),
                         failures=failures, timings=timings)
        self.save()
        print('Отчёт запуска: ' + str(self.state / 'last-run.html'), flush=True)

    def render(self):
        d = self.data
        esc = lambda value: escape(str(value), quote=True)
        cards = []
        for row in d['results']:
            checks = ''.join('<li>' + esc(c) + '</li>' for c in row['checks'])
            files = ''.join('<li>' + esc(f['filename']) + ' — ' +
                            esc(FILE_STATUS.get(f['status'], f['status'])) +
                            ('; ПРЕДВАРИТЕЛЬНЫЙ' if f.get('quality') == 'preliminary' else '') + '</li>' for f in row['files'])
            error = ('<p class="error">' + esc(row['error_type']) + ': ' +
                     esc(error_hint(row['error_type'], row.get('error_reason'))) + '</p>') if row.get('error_type') else ''
            lists = ('<ul>' + checks + '</ul>' if checks else '') + ('<ul>' + files + '</ul>' if files else '')
            cards.append(f'<section><h2>{esc(row["name"])} <span>{esc(STATUS[row["status"]])}</span></h2>'
                         f'<p>{esc(row["stage"])}</p>{error}{lists}</section>')
        count = sum(len(r['files']) for r in d['results'])
        complete = sum(r['status'] in {'complete', 'local', 'preliminary'} for r in d['results'])
        preliminary_count = sum(f.get('quality') == 'preliminary' for r in d['results'] for f in r['files'])
        rechecks = ''.join('<li>' + esc(r['date'] + ' · ' + r['base']) + ': ' +
                          esc(STATUS[r['status']]) + '</li>' for r in d.get('rechecks', []))
        warning = '<p class="error">' + '<br>'.join(
            esc(f.get('base', 'Запуск')) + ': ' + esc(f['error_type']) + ' — ' + esc(error_hint(f['error_type'], f.get('error_reason')))
            for f in d['failures']) + '</p>' if d['failures'] else ''
        mode = 'Только локальные файлы, без публикации на Drive' if d['local_only'] else 'Выгрузка на Google Drive'
        period = d['date'] + (' — ' + d['end_date'] if d.get('end_date') else '')
        log_key = d['date'] + ('_' + d['end_date'] if d.get('end_date') else '')
        file_count_label = 'Подготовлено локальных файлов' if d['local_only'] else 'Подтверждено файлов на Drive'
        if d['local_only']:
            count = complete * 3
        return f'''<!doctype html><html lang="ru"><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Результат запуска 1С — {esc(period)}</title>
<style>body{{font:16px/1.55 Segoe UI,Arial,sans-serif;max-width:1080px;margin:40px auto;padding:0 24px;color:#172d40;background:#f3f6f9}}h1{{line-height:1.2}}section{{background:white;border:1px solid #d4dfe7;border-radius:12px;padding:16px 22px;margin:18px 0}}h2{{font-size:20px}}span{{font-size:14px;margin-left:15px;color:#526676}}li{{margin:6px 0;overflow-wrap:anywhere}}.error{{color:#a32424}}code{{overflow-wrap:anywhere}}@media print{{body{{margin:0;background:white}}section{{break-inside:avoid}}}}</style>
<h1>Результат запуска 1С</h1><p>{'Отчётный период' if d.get('end_date') else 'Отчётный день'}: <b>{esc(period)}</b> · {esc(mode)}</p>
<p>Состояние: <b>{esc(STATUS[d['status']])}</b>. Завершено баз: {complete}/{len(d['results'])}. {file_count_label}: {count}/{len(d['results'])*3}.</p>
<p>Из опубликованных файлов предварительных: {preliminary_count}. Подтверждение загрузки файла не означает устранение расхождения его данных.</p>
<p>Начало: {esc(d['started_at'])}<br>Завершение: {esc(d['finished_at'] or 'Ещё выполняется')}<br>Прошло: {d['elapsed_seconds']:.1f} сек.</p>
<p>Это проверки конкретной выгрузки. Автотесты исходного кода запускаются отдельно при изменении программы.
Проверки источников отмечаются только после успешной обработки; при повторе явно указан ранее проверенный комплект.</p>
{warning}{''.join(cards)}<p>Подробный журнал: <code>{esc(self.state / (log_key + '.log'))}</code>.</p>
{'<h2>Повторные проверки прошлых дней</h2><ul>' + rechecks + '</ul>' if rechecks else ''}
<p><a href="{esc((self.state / 'reconciliation' / 'index.html').as_uri())}">Контроль расхождений: открытые и устранённые</a></p>
<p><a href="{esc((self.state.parent / 'program-tests.html').as_uri())}">Последние автотесты программы</a> · <a href="{esc((self.state.parent / 'program-guide.html').as_uri())}">Инструкция по запуску и переносу</a></p>
<p>Страница обновляется при этапах работы; нажмите F5, чтобы увидеть новое состояние. Если процесс был принудительно закрыт, сохранённое состояние может остаться «В работе».</p></html>'''
