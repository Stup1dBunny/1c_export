"""Persist business discrepancies and recheck their original calendar day.

Only a complete, parsed source set may close an issue. Transport failures and
receipt status changes by themselves cannot close one. No customer/card PII.
"""
from __future__ import annotations

import hashlib
import json
from datetime import date, datetime
from html import escape
from pathlib import Path

from report_tools import MSK, write_json, period_key
from source_extract import amount, extract_check_list, json_numbers

METRICS = ('receipt_count', 'sale_receipt_count', 'refund_receipt_count',
           'sales_amount', 'refund_amount')

LABELS = {'summary': 'Сводный отчёт', 'journal': 'Все действительные чеки', 'archived': 'Архивные чеки для сверки сводного отчёта', 'receipt_count': 'Всего чеков',
          'sale_receipt_count': 'Продаж', 'refund_receipt_count': 'Возвратов', 'sales_amount': 'Сумма продаж, ₽',
          'refund_amount': 'Сумма возвратов, ₽', 'goods_amount': 'Товарный отчёт, ₽',
          'journal_amount': 'Журнал чеков, ₽', 'goods_minus_journal': 'Разница, ₽',
          'timestamp': 'Дата и время', 'cash': 'Касса', 'shift': 'Смена', 'number': 'Номер чека',
          'status': 'Статус 1С', 'operation': 'Операция', 'amount': 'Сумма, ₽',
          'date': 'Дата покупки', 'matching_numbers': 'Подходящих номеров карты'}


def describe_detail(detail):
    return '\n'.join(LABELS.get(k, k) + ': ' + (describe_detail(v) if isinstance(v, dict) else
                    'нет в источнике' if v is None else str(v))
                    for k, v in detail.items() if k != 'shop')


def stable_id(*parts):
    return hashlib.sha256(json.dumps(parts, ensure_ascii=False).encode()).hexdigest()[:24]


def reconcile_sources(receipts, goods, text, report_date, end_date=None):
    bonus = extract_check_list(text, report_date, receipts, preliminary=True, end_date=end_date)
    expected = {r['shop']: r for r in receipts['shops']}
    actual = bonus['shops']
    merchandise = {r['shop']: r for r in goods['shops']}
    archived, issued_shops = {}, set()
    for check in bonus['records']:
        if check['status'] == 'Пробитый':
            issued_shops.add(check['shop'])
            continue
        totals = archived.setdefault(check['shop'], {k: 0 for k in METRICS})
        refund = check['operation'] == 'Возврат'
        totals['receipt_count'] += 1
        totals['refund_receipt_count' if refund else 'sale_receipt_count'] += 1
        totals['refund_amount' if refund else 'sales_amount'] += amount(check['amount'])
    issues, sales, unverified_receipts = [], [], set()
    for shop in sorted(expected.keys() | actual.keys() | merchandise.keys()):
        r, j, g = expected.get(shop), actual.get(shop), merchandise.get(shop)
        # Missing shop is distinct from a numeric zero, including net-zero
        # sale/refund days. Compare gross amounts and both operation counts.
        full_match = r is not None and j is not None and all(amount(r[k]) == j[k] for k in METRICS)
        a = archived.get(shop)
        archived_match = ((a is None and r is None) or
                          (a is not None and r is not None and all(amount(r[k]) == a[k] for k in METRICS)))
        # Both statuses represent real purchases. The native summary may use
        # only archived receipts. Compare like-for-like subsets instead of
        # treating valid issued receipts as a discrepancy. Output always uses
        # the FULL journal; no purchase or refund is dropped.
        matched = full_match or (shop in issued_shops and archived_match)
        if not matched:
            unverified_receipts.add(shop)
            detail = {
                'summary': {k: r[k] for k in METRICS} if r else None,
                'journal': {k: j[k] for k in METRICS} if j else None,
                'archived': a,
            }
            issues.append({'id': stable_id('journal', shop), 'kind': 'journal', 'shop': shop,
                'reports': ['Продажи', 'Бонусы'], 'detail': detail,
                'message': 'Итоги источников не совпали даже с учётом разных наборов чеков. Все действительные покупки включены; суммы предварительные.'})
            for check in bonus['records']:
                if check['shop'] == shop:
                    issues.append({'id': stable_id('receipt', shop, check['cash'], check['shift'],
                                                   check['number'], check['timestamp']),
                        'kind': 'receipt', 'shop': shop, 'reports': ['Продажи', 'Бонусы'],
                        'detail': check,
                        'message': 'Чек относится к магазину с несверенными итогами. Это не доказывает ошибку самого чека.'})
        goods_matched = j is not None and g is not None and amount(g['amount']) == j['revenue']
        if not goods_matched:
            issues.append({'id': stable_id('goods', shop), 'kind': 'goods', 'shop': shop,
                'reports': ['Продажи'],
                'detail': {'goods_amount': g['amount'] if g else None,
                           'journal_amount': j['revenue'] if j else None,
                           'goods_minus_journal': amount(g['amount']) - j['revenue'] if g and j else None},
                'message': 'Количество алкоголя не подтверждено: нет сопоставимого набора чеков и товарных продаж.'})
        sales.append({'shop': shop, 'revenue': j['revenue'] if j is not None else None,
                      'receipt_count': j['receipt_count'] if j is not None else None,
                      'refund_receipt_count': j['refund_receipt_count'] if j is not None else None,
                      'refund_amount': j['refund_amount'] if j is not None else None,
                      'alcohol_quantity': g['categories'].get('Да', {}).get('quantity', 0) if goods_matched else None,
                      'verification': 'Проверено' if matched and goods_matched else 'Предварительно'})
    # User-approved policy: retain every issued sale/refund and its real amount
    # in preliminary output. Never silently drop a purchase because its closing
    # report has not caught up. Missing sources still remain blank, not zero.
    bonus['rows'] = [dict(r, verification='Предварительно' if r['shop'] in unverified_receipts else 'Проверено')
                     for r in bonus['rows']]
    issues = json_numbers(issues)
    for issue in issues:
        issue['description'] = describe_detail(issue['detail'])
    bonus['quality'] = 'preliminary' if unverified_receipts else 'verified'
    bonus['issues'] = [i for i in issues if 'Бонусы' in i['reports']]
    # A missing shop can have unknown bonus purchases even when every known
    # bonus row is reconciled. Do not display a deceptively complete total.
    bonus['incomplete_scope'] = bool(expected.keys() - actual.keys())
    return {'rows': sales, 'quality': 'preliminary' if issues else 'verified',
            'issues': issues}, bonus, issues


class DiscrepancyLedger:
    def __init__(self, state):
        self.root = Path(state) / 'reconciliation'

    def path(self, base_id, day, end_day=None):
        return self.root / (period_key(day, end_day) + '_' + base_id + '.json')

    def load(self, base_id, day, end_day=None):
        path = self.path(base_id, day, end_day)
        return json.loads(path.read_text(encoding='utf-8')) if path.exists() else None

    def observe(self, base, day, issues, generation, end_day=None, *, analysis_version=1):
        previous = self.load(base['id'], day, end_day)
        now = datetime.now(MSK).isoformat()
        saved = previous or {'base_id': base['id'], 'base_name': base['name'], 'date': day.isoformat(),
                             **({'end_date': end_day.isoformat()} if end_day else {}),
                             'issues': [], 'history': []}
        # Rebuilding after an interrupted upload is not a new source check.
        if saved.get('generation') == generation and saved.get('analysis_version', 1) == analysis_version:
            return saved
        old = {i['id']: i for i in saved['issues']}
        current = {i['id']: i for i in issues}
        unverified_shops = {i['shop'] for i in issues if i['kind'] == 'journal'}
        for identity, prior in old.items():
            if (identity not in current and prior['kind'] == 'receipt' and prior['status'] == 'open'
                    and prior['shop'] in unverified_shops):
                current[identity] = {k: prior[k] for k in ('id', 'kind', 'shop', 'reports', 'detail', 'message')}
                current[identity]['message'] = 'Ранее наблюдавшийся чек отсутствует в текущем активном наборе. Итоги магазина ещё не сверены.'
        records = []
        for identity in sorted(old.keys() | current.keys()):
            prior, issue = old.get(identity), current.get(identity)
            if issue:
                records.append(dict(issue, status='open', first_seen=prior['first_seen'] if prior else now,
                                    last_checked=now, resolved_at=None))
            else:
                records.append(dict(prior, status='resolved', last_checked=now,
                                    resolved_at=prior.get('resolved_at') or now))
        saved.update(issues=records, generation=generation, checked_at=now,
                     open_count=len(current), publication_pending=True, analysis_version=analysis_version)
        saved['history'].append({'checked_at': now, 'generation': generation, 'issues': issues,
                                'analysis_version': analysis_version})
        self.save(saved)
        return saved

    def published(self, base_id, day, end_day=None):
        saved = self.load(base_id, day, end_day)
        if saved:
            saved['publication_pending'] = False
            self.save(saved)

    def save(self, saved):
        write_json(self.path(saved['base_id'], date.fromisoformat(saved['date']),
                             date.fromisoformat(saved['end_date']) if saved.get('end_date') else None), saved)
        self.render()

    def pending(self):
        if not self.root.exists():
            return []
        return [s for p in sorted(self.root.glob('*.json'))
                if not (s := json.loads(p.read_text(encoding='utf-8'))).get('end_date')
                and (s.get('open_count') or s.get('publication_pending'))]

    def render(self):
        sections = []
        for p in sorted(self.root.glob('*.json'), reverse=True):
            s = json.loads(p.read_text(encoding='utf-8'))
            if not s['issues']:
                continue
            lines = []
            for i in sorted(s['issues'], key=lambda i: i['status'] != 'open'):
                detail = describe_detail(i['detail'])
                lines.append('<tr><td>' + ('Открыто' if i['status'] == 'open' else 'Устранено') +
                    '</td><td>' + escape(i['shop']) + '</td><td>' + escape(i['message']) +
                    '</td><td>' + escape(detail) + '</td><td>' + escape(i['first_seen']) +
                    '</td><td>' + escape(i['last_checked']) + '</td></tr>')
            sections.append('<h2>' + escape(s['date'] + (' — ' + s['end_date'] if s.get('end_date') else '') + ' · ' + s['base_name']) + '</h2>' +
                '<p>Открыто: ' + str(s['open_count']) + '. Проверок источников: ' + str(len(s['history'])) +
                ('; публикация ожидается.' if s.get('publication_pending') else '.') + '</p><table>' +
                '<tr><th>Состояние</th><th>Магазин</th><th>Причина</th><th>Исходные значения</th>' +
                '<th>Обнаружено</th><th>Проверено</th></tr>' + ''.join(lines) + '</table>')
        html = '<!doctype html><html lang="ru"><meta charset="utf-8"><title>Контроль расхождений 1С</title>' + \
            '<style>body{font:15px/1.5 Segoe UI,Arial;margin:30px;color:#203548}table{border-collapse:collapse;width:100%}td,th{padding:10px;border:1px solid #ccd6df;text-align:left;overflow-wrap:anywhere}th{background:#eaf0f5}</style>' + \
            '<h1>Контроль расхождений 1С</h1><p>Проверяется исходный день. Смена статуса чека сама по себе не устраняет расхождение. ' + \
            'Ошибка связи не закрывает запись. Предыдущие наблюдения сохраняются.</p>' + ''.join(sections) + '</html>'
        target = self.root / 'index.html'
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix('.tmp')
        tmp.write_text(html, encoding='utf-8')
        tmp.replace(target)
