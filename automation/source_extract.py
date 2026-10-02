"""Read verified 1C XLSX exports without executing formulas or writing workbooks."""
from __future__ import annotations

import argparse
import csv
import io
import json
import posixpath
import re
import zipfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from report_tools import inspect_xlsx, write_json

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


@dataclass
class SourceRow:
    index: int
    level: int
    cells: dict


def read_first_sheet(path: Path) -> tuple[list[SourceRow], dict]:
    if path.suffix.casefold() != '.xlsx':
        raise ValueError('Only completed .xlsx downloads can be extracted')
    provenance = inspect_xlsx(path)
    with zipfile.ZipFile(path) as archive:
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            root = ET.fromstring(archive.read('xl/sharedStrings.xml'))
            strings = [''.join(t.text or '' for t in si.findall('.//s:t', NS))
                       for si in root.findall('s:si', NS)]
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        sheet = workbook.find('s:sheets/s:sheet', NS)
        relationships = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        target = next(rel.attrib['Target'] for rel in relationships if rel.attrib['Id'] == sheet.attrib[RID])
        target = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        if not target.startswith('xl/worksheets/'):
            raise ValueError('Unexpected worksheet relationship')
        data = ET.fromstring(archive.read(target))
        rows = []
        for row in data.findall('s:sheetData/s:row', NS):
            cells = {}
            for cell in row.findall('s:c', NS):
                if cell.find('s:f', NS) is not None:
                    raise ValueError('Unexpected formula in a raw 1C export')
                column = re.match(r'[A-Z]+', cell.attrib['r']).group()
                value = cell.findtext('s:v', default=None, namespaces=NS)
                cell_type = cell.attrib.get('t', 'n')
                if cell_type == 's' and value is not None:
                    value = strings[int(value)]
                elif cell_type == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.findall('s:is//s:t', NS))
                elif cell_type == 'n' and value is not None:
                    value = Decimal(value)
                elif cell_type == 'e' and value not in (None, ''):
                    raise ValueError('Spreadsheet error in a raw 1C export')
                cells[column] = value
            rows.append(SourceRow(int(row.attrib['r']), int(row.attrib.get('outlineLevel', '0')), cells))
    return rows, provenance | {'sheet': sheet.attrib['name']}


def amount(value, *, blank_zero=False) -> Decimal:
    if value is None or value == '':
        if blank_zero:
            return Decimal(0)
        raise ValueError('Missing numeric value')
    if isinstance(value, Decimal):
        result = value
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        result = Decimal(str(value))
    elif isinstance(value, str):
        text = re.sub(r'[\s\u00a0\u202f]', '', value).replace(',', '.')
        if not re.fullmatch(r'[+-]?\d+(?:\.\d+)?', text):
            raise ValueError('Invalid numeric text in the 1C report')
        result = Decimal(text)
    else:
        raise ValueError('Unexpected numeric type')
    if not result.is_finite():
        raise ValueError('Non-finite report amount')
    return result


def extract_check_list(text: str, report_date: str, receipts: dict, *, preliminary=False, end_date=None) -> dict:
    """Aggregate the complete native text export, never virtualized UI rows.

    Native refunds carry positive amounts. Use their explicit operation to
    subtract them, and reconcile gross amounts and both counts per shop.
    """
    reader = csv.DictReader(io.StringIO(text.lstrip('\ufeff')), delimiter='\t')
    required = {'Дата', 'Смена', '№ чека ККМ', 'Сумма', 'Статус',
                'Дисконтная карта', 'Магазин', 'Касса (ККМ)'}
    if not required.issubset(reader.fieldnames or []):
        raise ValueError('Required columns missing from the complete check list')
    has_operations = 'Операция' in reader.fieldnames
    if not has_operations and amount(receipts['totals']['refund_receipt_count']):
        raise ValueError('Refund days require operation type in the check list')
    seen, shops, cards, records = set(), {}, {}, []
    excluded = {'annulled': 0, 'unissued': 0}
    archived_without_fiscal_fields = 0
    refunds_requiring_summary = set()
    for row in reader:
        if None in row or any(value is None for value in row.values()):
            raise ValueError('Malformed check list row')
        timestamp = datetime.strptime(row['Дата'], '%d.%m.%Y %H:%M:%S')
        purchase_date = timestamp.date().isoformat()
        if not report_date <= purchase_date <= (end_date or report_date):
            raise ValueError('Check outside the requested calendar day or period')
        if row['Статус'] == 'Аннулированный':
            excluded['annulled'] += 1
            continue
        if not row['Статус'].strip():
            # User confirmed on 02 Oct: documents with an empty status are
            # pending deletion, not sales, even if a shift is already assigned.
            # Any goods-report difference remains visible in reconciliation.
            excluded['unissued'] += 1
            continue
        if (row['Статус'] == 'Отложенный' and row['№ чека ККМ'] == ''
                and row['Смена'] == '' and row.get('Отчет о розничных продажах') == ''):
            # Observed in MKret2: a deferred basket with no issued fiscal
            # check, shift or closing report. Active checks reconcile below.
            excluded['deferred_unissued'] = excluded.get('deferred_unissued', 0) + 1
            continue
        if row['Статус'] != 'Архивный' and not (preliminary and row['Статус'] == 'Пробитый'):
            raise ValueError('Check status needs explicit validation: ' + row['Статус'])
        if row['Статус'] == 'Пробитый' and not has_operations:
            raise ValueError('Issued checks require explicit operation type')
        if not row['№ чека ККМ'] or not row['Смена']:
            if row['Статус'] != 'Архивный':
                raise ValueError('Issued check has no fiscal number or shift')
            # VDA 30 Sep: an archived 472.00 purchase has a closing retail
            # report but both fiscal fields are empty. It is included in the
            # independent receipt totals. Preserve it and require the full
            # per-shop count, gross-sale and refund reconciliation below.
            retail = row.get('Отчет о розничных продажах', '')
            if (not row['№ чека ККМ'] and not retail and row['Смена'].strip()
                    and row['Касса (ККМ)'].strip() and row.get('Операция') == 'Возврат'):
                # SVU August: an archived refund has a shift but neither a
                # fiscal number nor a closing report. Keep the source fields
                # empty and require independent count/gross-total agreement,
                # even in preliminary mode, before accepting this fallback.
                refunds_requiring_summary.add(row['Магазин'])
                key = ('archived-refund', row['Касса (ККМ)'], row['Смена'], row['Дата'])
            elif not row['№ чека ККМ'] and retail.startswith('Отчет о розничных продажах '):
                key = ('retail', row['Касса (ККМ)'], retail, row['Дата'])
            else:
                raise ValueError('Archived check has no fiscal number or shift')
            # SVU September also has a reconciled archived document with an
            # assigned shift but no fiscal number. The closing report, cash
            # register and timestamp still identify it without inventing an ID.
            archived_without_fiscal_fields += 1
        else:
            key = (row['Касса (ККМ)'], row['Смена'], row['№ чека ККМ'], row['Дата'])
        if key in seen:
            raise ValueError('Duplicate check in source')
        seen.add(key)
        shop, card = row['Магазин'], row['Дисконтная карта']
        value = amount(row['Сумма'])
        if not shop or value < 0:
            raise ValueError('Invalid check shop or amount')
        operation = row['Операция'] if has_operations else 'Продажа'
        if operation not in ('Продажа', 'Возврат'):
            raise ValueError('Check operation needs explicit validation: ' + operation)
        refund = operation == 'Возврат'
        signed_value = -value if refund else value
        records.append({'shop': shop, 'timestamp': row['Дата'], 'cash': row['Касса (ККМ)'],
                        'shift': row['Смена'], 'number': row['№ чека ККМ'],
                        'status': row['Статус'], 'operation': operation, 'amount': value})
        aggregate = shops.setdefault(shop, {'receipt_count': 0, 'revenue': Decimal(0),
            'sale_receipt_count': 0, 'refund_receipt_count': 0,
            'sales_amount': Decimal(0), 'refund_amount': Decimal(0)})
        aggregate['receipt_count'] += 1
        aggregate['revenue'] += signed_value
        aggregate['refund_receipt_count' if refund else 'sale_receipt_count'] += 1
        aggregate['refund_amount' if refund else 'sales_amount'] += value
        if card:
            aggregate = cards.setdefault((purchase_date, shop, card), {
                'date': purchase_date, 'card': card, 'shop': shop,
                'amount': Decimal(0), 'receipt_count': 0})
            aggregate['amount'] += signed_value
            aggregate['receipt_count'] += 1
    expected = {r['shop']: r for r in receipts['shops']}
    for shop in refunds_requiring_summary:
        metrics = ('receipt_count', 'sale_receipt_count', 'refund_receipt_count',
                   'sales_amount', 'refund_amount')
        archived = dict.fromkeys(metrics, Decimal(0))
        for record in records:
            if record['shop'] != shop or record['status'] != 'Архивный':
                continue
            is_refund = record['operation'] == 'Возврат'
            archived['receipt_count'] += 1
            archived['refund_receipt_count' if is_refund else 'sale_receipt_count'] += 1
            archived['refund_amount' if is_refund else 'sales_amount'] += record['amount']
        reference = expected.get(shop)
        if reference is None or not any(
                all(actual[k] == amount(reference[k]) for k in metrics)
                for actual in (shops[shop], archived)):
            raise ValueError('Archived refund without fiscal references does not reconcile with the receipt report')
    if not preliminary and set(shops) != set(expected):
        raise ValueError('Check list shops differ from receipt report')
    for shop, actual in ([] if preliminary else shops.items()):
        reference = expected[shop]
        if (actual['receipt_count'] != amount(reference['receipt_count'])
                or actual['revenue'] != amount(reference['sales_amount']) - amount(reference['refund_amount'])):
            raise ValueError('Check list does not reconcile with the receipt report')
        for key in ('sale_receipt_count', 'refund_receipt_count', 'sales_amount', 'refund_amount'):
            if actual[key] != amount(reference[key]):
                raise ValueError('Check list operations do not reconcile with the receipt report: ' + key)
    rows = [cards[k] for k in sorted(cards)]
    return {'rows': rows, 'totals': {
        'receipt_count': sum(r['receipt_count'] for r in rows),
        'amount': sum((r['amount'] for r in rows), Decimal(0))},
        'all_receipt_count': len(seen), 'date': report_date,
        'excluded_records': excluded,
        'archived_without_fiscal_fields': archived_without_fiscal_fields,
        'shops': shops, 'records': records,
        'phone_source_verified': False}


def missing_card_reference(presentation: str) -> bool:
    """Recognize the native 1C dangling reference, not an empty search result."""
    return isinstance(presentation, str) and re.fullmatch(
        r'<Объект не найден> \(\d+:[0-9a-fA-F]{32}\)', presentation) is not None


def card_identity(presentation: str, *, allow_name: bool = False) -> tuple[str | None, str | None]:
    """Recognize the two observed layouts; owner may be absent from the check."""
    match = re.fullmatch(r'(\S+) \((.+)\)', presentation)
    if match:
        return match.groups()
    if re.fullmatch(r'\d+', presentation):
        return presentation, None
    if allow_name and isinstance(presentation, str) and presentation.strip():
        # Some checks show only the card's name. This is a search key, not an
        # owner or card number; both must be read from the exact catalog item.
        return None, None
    raise ValueError('Unrecognized card number and owner presentation')


def attach_owner_contacts(bonuses: dict, contacts: dict) -> dict:
    """Keep absent source phones blank, but reject unverified contact lookups."""
    result = dict(bonuses)
    rows = []
    for row in bonuses['rows']:
        number, expected_owner = card_identity(row['card'], allow_name=True)
        contact = contacts.get(number or row['card'])
        if not contact or 'phone' not in contact or not contact.get('source'):
            raise ValueError('Owner phone was not checked for card ' + (number or row['card']))
        if missing_card_reference(row['card']):
            if (contact.get('card_reference_present') is not False
                    or contact.get('card_presentation') != row['card']
                    or contact.get('owner') != '' or contact.get('phone') != ''):
                raise ValueError('A missing card reference cannot be assigned an inferred owner or phone')
            rows.append(dict(row, phone=''))
            continue
        if number is None:
            number = contact.get('card_number')
            if (contact.get('card_presentation') != row['card'] or not isinstance(number, str)
                    or not number.strip()):
                raise ValueError('Named card was not resolved to a verified card number')
        owner = contact.get('owner')
        confirmed_missing = (expected_owner is None and owner == '' and contact.get('phone') == ''
                             and contact.get('owner_reference_present') is False
                             and contact.get('card_number') == number)
        if not isinstance(owner, str) or (not owner.strip() and not confirmed_missing):
            raise ValueError('Owner reference was not verified for card ' + number)
        if expected_owner is not None and owner.strip() != expected_owner.strip():
            raise ValueError('Owner reference differs from receipt card presentation')
        if not isinstance(contact['phone'], str):
            raise ValueError('Phone must be a string, including when empty in 1C')
        # Some bases export only the card barcode in checks. Resolve the owner
        # through that exact card instead of omitting their name from the report.
        display = number + ' (' + owner.strip() + ')' if owner.strip() else number
        rows.append(dict(row, card=display, phone=contact['phone']))
    result.update(rows=rows, phone_source_verified=True,
                  missing_card_reference_rows=sum(missing_card_reference(r['card']) for r in rows))
    return result


CHECK_COLUMNS = {
    'receipt_count': 'Кол-во чеков',
    'sale_receipt_count': 'Кол-во чеков продажи',
    'refund_receipt_count': 'Кол-во чеков возврата',
    'sales_amount': 'Сумма продажи (руб.)',
    'refund_amount': 'Сумма возврата (руб.)',
}


def extract_receipts(rows: list[SourceRow]) -> dict:
    by_number = {r.index: r for r in rows}
    if any(by_number.get(i) is None or by_number[i].cells.get('A') != name
           for i, name in ((1, 'Магазин'), (2, 'Касса ККМ'), (3, 'Кассир'))):
        raise ValueError('Receipt report grouping has changed; shop-level extraction is unsafe')
    header = by_number[2].cells
    columns = {}
    for key, name in CHECK_COLUMNS.items():
        matches = [col for col, value in header.items() if value == name]
        if len(matches) != 1:
            raise ValueError(f'Required receipt column is missing or duplicated: {name}')
        columns[key] = matches[0]

    shops, totals = [], []
    seen = set()
    for row in rows:
        if row.index <= 3:
            continue
        if row.level not in (0, 1, 2):
            raise ValueError('Unexpected receipt grouping depth')
        if row.level != 0 or not any(v is not None for v in row.cells.values()):
            continue
        shop = row.cells.get('A')
        if not isinstance(shop, str) or not shop.strip():
            raise ValueError('A shop-level row has no shop name')
        numbers = {key: amount(row.cells.get(col), blank_zero=True) for key, col in columns.items()}
        for key in ('receipt_count', 'sale_receipt_count', 'refund_receipt_count'):
            if numbers[key] != numbers[key].to_integral_value():
                raise ValueError('Receipt count is not an integer')
        if shop == 'Итого':
            totals.append(numbers)
        else:
            if shop in seen:
                raise ValueError('Duplicate shop summary')
            seen.add(shop)
            shops.append({'shop': shop, 'source_row': row.index, **numbers})
    if len(totals) != 1:
        raise ValueError('Expected exactly one report total')
    for key in columns:
        if sum((shop[key] for shop in shops), Decimal(0)) != totals[0][key]:
            raise ValueError(f'Shop summaries do not reconcile to the report total: {key}')
    return {'shops': shops, 'totals': totals[0]}


def json_numbers(value):
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, dict):
        return {k: json_numbers(v) for k, v in value.items()}
    if isinstance(value, list):
        return [json_numbers(v) for v in value]
    return value


def extract_stock(rows: list[SourceRow], *, supplier_missing: bool = False) -> dict:
    """Flatten the verified shop/product layout; reconcile details and totals."""
    by_number = {r.index: r for r in rows}
    header = by_number.get(4)
    detail_header = by_number.get(5)
    if (header is None or detail_header is None or header.cells.get('A') != 'Магазин'
            or detail_header.cells.get('A') != 'Номенклатура'):
        raise ValueError('Stock grouping has changed')
    supplier_columns = [col for col, value in detail_header.cells.items()
                        if value == 'Поставщик/Дистребьютер (Вино)']
    if len(supplier_columns) > 1:
        raise ValueError('Stock supplier column is duplicated')
    supplier_column = supplier_columns[0] if supplier_columns else None
    if supplier_column is None:
        # Only a base explicitly verified and configured as lacking this field
        # may use the observed layout without it. Native empty reports have
        # narrower merged labels, so quantity can move from D to C.
        if (not supplier_missing or 'Количество' not in header.cells.values()
                or any(v not in (None, '') for k, v in detail_header.cells.items() if k != 'A')):
            raise ValueError('Stock supplier column has changed')
    columns = {'closing': 'Количество', 'opening': 'Начальный остаток',
               'incoming': 'Приход', 'outgoing': 'Расход'}
    for key, name in columns.copy().items():
        matches = [col for col, value in header.cells.items() if value == name]
        if len(matches) != 1:
            raise ValueError(f'Missing stock column: {name}')
        columns[key] = matches[0]
    shops, details, totals = [], [], []
    current_shop = None
    for row in rows:
        if row.index <= 5 or not any(v is not None for v in row.cells.values()):
            continue
        label = row.cells.get('A')
        if row.level not in (0, 1) or not isinstance(label, str) or not label.strip():
            raise ValueError('Unexpected stock row/grouping')
        nums = {key: amount(row.cells.get(col), blank_zero=True) for key, col in columns.items()}
        if nums['opening'] + nums['incoming'] - nums['outgoing'] != nums['closing']:
            raise ValueError(f'Stock balance failed at source row {row.index}')
        if row.level == 0:
            if label == 'Итого':
                totals.append(nums)
                current_shop = None
            else:
                if any(s['shop'] == label for s in shops) or totals:
                    raise ValueError('Duplicate/out-of-order stock shop summary')
                current_shop = label
                shops.append({'shop': label, 'source_row': row.index, **nums})
        else:
            if current_shop is None:
                raise ValueError('Stock detail has no parent shop')
            details.append({'shop': current_shop, 'product': label,
                            'supplier': (row.cells.get(supplier_column) or '') if supplier_column else '',
                            'source_row': row.index, **nums})
    if len(totals) != 1:
        raise ValueError('Expected one stock report total')
    for key in columns:
        if sum((s[key] for s in shops), Decimal(0)) != totals[0][key]:
            raise ValueError(f'Stock shops do not reconcile: {key}')
        for shop in shops:
            if sum((d[key] for d in details if d['shop'] == shop['shop']), Decimal(0)) != shop[key]:
                raise ValueError(f'Stock details do not reconcile: {key}')
    return {'rows': details, 'shops': shops, 'totals': totals[0],
            'supplier_source': 'report_column' if supplier_column else 'absent_in_base'}


def extract_alcohol(rows: list[SourceRow]) -> dict:
    by_number = {r.index: r for r in rows}
    if (by_number.get(4) is None or by_number[4].cells.get('A') != 'Магазин'
            or by_number.get(5) is None
            or by_number[5].cells.get('A') != 'Номенклатура.Алкогольная продукция'):
        raise ValueError('Alcohol report grouping has changed')
    header = by_number[4].cells
    columns = {}
    for key, name in {'quantity': 'Количество товаров', 'amount': 'Сумма продаж со скидкой'}.items():
        matches = [col for col, value in header.items() if value == name]
        if len(matches) != 1:
            raise ValueError('Alcohol report quantity/amount column is missing or duplicated')
        columns[key] = matches[0]
    shops, totals = [], []
    current = None
    for row in rows:
        if row.index <= 5 or not any(v is not None for v in row.cells.values()):
            continue
        label = row.cells.get('A')
        nums = {key: amount(row.cells.get(col), blank_zero=True) for key, col in columns.items()}
        if row.level == 0:
            if label == 'Итого':
                totals.append(nums)
                current = None
            else:
                if not isinstance(label, str) or not label.strip() or totals or any(s['shop'] == label for s in shops):
                    raise ValueError('Unexpected/duplicate alcohol shop summary')
                current = {'shop': label, 'source_row': row.index, **nums, 'categories': {}}
                shops.append(current)
        elif row.level == 1:
            if current is None or label not in ('Да', 'Нет') or label in current['categories']:
                raise ValueError('Unexpected/duplicate alcohol classification')
            current['categories'][label] = nums
        else:
            raise ValueError('Unexpected alcohol grouping depth')
    if len(totals) != 1:
        raise ValueError('Expected one alcohol report total')
    for key in columns:
        if sum((s[key] for s in shops), Decimal(0)) != totals[0][key]:
            raise ValueError('Alcohol shop summaries do not reconcile')
        for shop in shops:
            if sum((c[key] for c in shop['categories'].values()), Decimal(0)) != shop[key]:
                raise ValueError('Alcohol classification does not reconcile to its shop')
    return {'shops': shops, 'totals': totals[0]}


def prepare_sales(receipts: dict, alcohol: dict, *, revenue_source: str = 'strict_match') -> dict:
    """Use the approved revenue source while retaining cross-report differences.

    The active-check policy relaxes only the comparison with goods revenue.
    Source totals, store scope and the complete check journal remain validated.
    """
    if revenue_source not in {'strict_match', 'active_checks'}:
        raise ValueError('Unknown sales revenue source')
    checks = {r['shop']: r for r in receipts['shops']}
    sales = {r['shop']: r for r in alcohol['shops']}
    if checks.keys() != sales.keys():
        raise ValueError('Receipt and sales reports contain different shops')
    result, differences = [], []
    for name, r in checks.items():
        s = sales[name]
        net_revenue = amount(r['sales_amount']) - amount(r['refund_amount'])
        goods_revenue = amount(s['amount'])
        if net_revenue != goods_revenue:
            if revenue_source == 'strict_match':
                raise ValueError(f'Receipt revenue differs from the sales report: {name}; '
                                 f'checks={net_revenue}; sales={goods_revenue}; '
                                 f'difference={goods_revenue - net_revenue}')
            differences.append({'shop': name, 'checks_revenue': net_revenue,
                                'goods_revenue': goods_revenue,
                                'goods_minus_checks': goods_revenue - net_revenue})
        result.append({'shop': name, 'revenue': net_revenue, 'receipt_count': r['receipt_count'],
                       'refund_receipt_count': r['refund_receipt_count'], 'refund_amount': r['refund_amount'],
                       'alcohol_quantity': s['categories'].get('Да', {}).get('quantity', Decimal(0))})
    return {'rows': result, 'revenue_source': revenue_source,
            'reconciliation': {'differences': differences,
                               'goods_minus_checks': sum((d['goods_minus_checks'] for d in differences), Decimal(0))}}


def combine_sales(receipts: dict, alcohol: dict) -> list[dict]:
    """Legacy strict comparison; production uses prepare_sales with config policy."""
    return prepare_sales(receipts, alcohol)['rows']


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('file', type=Path)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    rows, source = read_first_sheet(args.file)
    result = json_numbers({'source': source, **extract_receipts(rows)})
    write_json(args.out, result)
    print(json.dumps({'shops': len(result['shops']), 'totals': result['totals']}, ensure_ascii=False))


if __name__ == '__main__':
    main()
