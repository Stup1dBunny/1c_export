"""Native bulk card lists and a per-base weekly cache. No owner forms are opened."""
from __future__ import annotations

import csv
import hashlib
import io
import json
import re
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from uuid import uuid4

from report_tools import write_json, inspect_xlsx, MSK
from source_extract import card_identity, missing_card_reference, amount, read_first_sheet


class CardNotFound(ValueError):
    pass


class CardConflict(ValueError):
    pass


def catalog_table_to_text(rows):
    """Preserve all string fields from the native tabular-document export."""
    fields = ('Код', 'Наименование', 'Код карты', 'Владелец карты')
    headers = []
    for row in rows:
        labels = [value for value in row.cells.values() if isinstance(value, str)]
        if set(fields).issubset(labels):
            if any(labels.count(field) != 1 for field in fields):
                raise ValueError('Duplicate columns in native card workbook')
            headers.append((row.index, {value: column for column, value in row.cells.items() if value in fields}))
    if len(headers) != 1:
        raise ValueError('Expected one complete header in the native card workbook')
    header, columns = headers[0]
    output = io.StringIO()
    writer = csv.writer(output, delimiter='\t')
    writer.writerow(fields)
    for row in rows:
        if row.index <= header:
            continue
        if not any(value not in (None, '') for value in row.cells.values()):
            continue
        if any(value not in (None, '') for column, value in row.cells.items() if column not in columns.values()):
            raise ValueError('Unexpected data outside the native card columns')
        values = [row.cells.get(columns[field]) for field in fields]
        if any(value is not None and not isinstance(value, str) for value in values):
            raise ValueError('Native card identifiers must be exported as text to preserve all digits')
        writer.writerow([value or '' for value in values])
    text = output.getvalue()
    parse_catalog(text)
    return text


def parse_catalog(text):
    reader = csv.DictReader(io.StringIO(text.lstrip('\ufeff')), delimiter='\t')
    fields = reader.fieldnames or []
    required = {'Код', 'Наименование', 'Код карты', 'Владелец карты'}
    if not required.issubset(fields) or len(fields) != len(set(fields)):
        raise ValueError('The bulk card list has missing or duplicate columns')
    by_number, by_name, codes = defaultdict(set), defaultdict(set), defaultdict(set)
    unnumbered = set()
    number_from_name = set()
    source_rows, groups = 0, 0
    for row in reader:
        if None in row or any(value is None for value in row.values()):
            raise ValueError('Incomplete bulk card row')
        source_rows += 1
        code, name, number, owner = (row[k].strip() for k in ('Код', 'Наименование', 'Код карты', 'Владелец карты'))
        if not number:
            # Observed legacy records carry the number in their presentation,
            # with an empty "Card code". An exact numeric presentation supplies
            # the number. Conflicting names in the two source fields must stay
            # ambiguous, using the same policy as duplicate numbered cards.
            named = re.fullmatch(r'(\d+)(?: \((.+)\))?', name)
            if named and (owner or named[2] is None):
                number = named[1]
                number_from_name.add(number)
                if named[2] is not None and named[2] != owner:
                    by_number[number].add(named[2])
        if not number:
            if owner:
                if not code:
                    raise ValueError('An unnumbered card with an owner has no internal code')
                unnumbered.add((code, name, owner))
            else:
                groups += 1
            continue
        if not code:
            raise ValueError('A numbered card has no internal catalog code')
        identity = (name, number, owner)
        # Some bases reuse visible catalog codes. This field is not a GUID
        # and must never be used to overwrite or join the card records.
        codes[code].add(identity)
        by_number[number].add(owner)
        if name:
            by_name[name].add(number)
    if not by_number:
        raise ValueError('The bulk card catalog is empty; its scope needs checking')
    return {'by_number': dict(by_number), 'by_name': dict(by_name), 'unnumbered': sorted(unnumbered),
            'number_from_name': number_from_name,
            'statistics': {'source_rows': source_rows, 'unique_numbers': len(by_number),
                           'group_rows': groups,
                           'unnumbered_records': len(unnumbered),
                           'numbers_from_presentation': len(number_from_name),
                           'reused_internal_codes': sum(len(v) > 1 for v in codes.values()),
                           'conflicting_numbers': sum(len(v - {''}) > 1 for v in by_number.values()),
                           'blank_and_named_numbers': sum('' in v and len(v) > 1 for v in by_number.values())}}


def attach_bulk_owners(bonuses, catalog):
    rows = []
    ambiguous = set()
    unnumbered_rows = 0
    resolution_issues = []
    updated_owners = set()
    for row in bonuses['rows']:
        presentation = row['card']
        if missing_card_reference(presentation):
            rows.append(dict(row, phone=''))
            continue
        number, receipt_owner = card_identity(presentation, allow_name=True)
        if number is None:
            numbers = catalog['by_name'].get(presentation.strip(), set())
            if not numbers:
                # The full catalog can verify a named card whose number is
                # genuinely empty. Retain the purchase and label the missing
                # identifier rather than inventing one or dropping the check.
                if any(name == presentation.strip() for code, name, owner in catalog.get('unnumbered', [])):
                    rows.append(dict(row, card=presentation + ' [номер не указан в 1С]', phone=''))
                    unnumbered_rows += 1
                    continue
                raise CardNotFound('A receipt card name is absent from the bulk catalog')
            if len(numbers) != 1:
                # The native journal supplied a name, not a unique identifier.
                # Retain its purchases once, without assigning either number.
                rows.append(dict(row, card=presentation + ' [номер неоднозначен]',
                                 phone='', verification='Предварительно'))
                identity = json.dumps([row.get('date'), row.get('shop'), presentation], ensure_ascii=False)
                detail = {'date': row.get('date'), 'matching_numbers': len(numbers),
                          'receipt_count': row['receipt_count'],
                          'amount': str(row['amount']) if row['amount'] is not None else None}
                resolution_issues.append({
                    'id': hashlib.sha256(('card-number:' + identity).encode('utf-8')).hexdigest()[:24],
                    'kind': 'card_number', 'shop': row.get('shop', ''), 'reports': ['Бонусы'],
                    'detail': detail,
                    'message': 'Наименованию карты в журнале соответствуют несколько номеров в справочнике. Покупки сохранены с исходным наименованием; номер не выбран.',
                    'description': f'Дата покупки: {row.get("date")}\nПодходящих номеров: {len(numbers)}\nКоличество чеков: {row["receipt_count"]}\nСумма, ₽: {row["amount"]}'})
                continue
            number = next(iter(numbers))
        owners = catalog['by_number'].get(number)
        if not owners:
            raise CardNotFound('A receipt card number is absent from the bulk catalog')
        named_owners = owners - {''}
        if len(named_owners) > 1:
            # User-approved policy: retain the purchase, never guess a person.
            owner = ''
            ambiguous.add(number)
        else:
            owner = next(iter(named_owners), '')
            if receipt_owner is not None and receipt_owner.strip() != owner:
                # User policy: the number is the key; the current bulk
                # catalog takes precedence over a historical receipt name.
                updated_owners.add(number)
        display = number + (' (' + owner + ')' if owner else '')
        rows.append(dict(row, card=display, phone=''))
    combined = {}
    for row in rows:
        key = (row.get('date'), row.get('shop'), row['card'])
        if key in combined:
            previous = combined[key]
            previous['amount'] = (None if previous['amount'] is None or row['amount'] is None
                                  else amount(previous['amount']) + amount(row['amount']))
            previous['receipt_count'] = (None if previous['receipt_count'] is None or row['receipt_count'] is None
                                         else previous['receipt_count'] + row['receipt_count'])
        else:
            combined[key] = dict(row)
    rows = list(combined.values())
    if resolution_issues:
        rows = [dict(r, verification=r.get('verification', 'Проверено')) for r in rows]
    result = dict(bonuses, rows=rows, card_source_verified=True,
                card_source='native_bulk_catalog', phone_source_verified=False,
                phone_policy='not_requested',
                ambiguous_owner_card_count=len(ambiguous),
                updated_owner_card_count=len(updated_owners),
                ambiguous_card_rows=len(resolution_issues),
                card_resolution_issues=resolution_issues,
                unnumbered_card_rows=unnumbered_rows,
                missing_card_reference_rows=sum(missing_card_reference(r['card']) for r in rows))
    if resolution_issues:
        result['quality'] = 'preliminary'
        result['issues'] = list(bonuses.get('issues', [])) + resolution_issues
    return result


class CardCatalogCache:
    def __init__(self, folder, base_id, link, now=None):
        self.folder = Path(folder) / base_id
        self.marker = self.folder / 'current.json'
        self.context = {'schema': 1, 'base_id': base_id, 'cards_link': link}
        self.now = now or (lambda: datetime.now(timezone.utc))

    def load(self):
        if not self.marker.is_file():
            return None
        data = json.loads(self.marker.read_text(encoding='utf-8'))
        if data.get('context') != self.context:
            raise ValueError('Card catalog cache belongs to a different base or source')
        filename = data['filename']
        if Path(filename).name != filename or not filename.endswith('.tsv'):
            raise ValueError('Unsafe catalog snapshot filename')
        snapshot = self.folder / filename
        payload = snapshot.read_bytes()
        if hashlib.sha256(payload).hexdigest() != data['sha256']:
            raise ValueError('Saved bulk card catalog has changed')
        fetched = datetime.fromisoformat(data['fetched_at'])
        if fetched.tzinfo is None or fetched > self.now() + timedelta(minutes=5):
            raise ValueError('Invalid catalog snapshot timestamp')
        if self.now() - fetched >= timedelta(days=7):
            return None
        return parse_catalog(payload.decode('utf-8-sig'))

    def refresh(self, collect):
        self.folder.mkdir(parents=True, exist_ok=True)
        snapshot = self.folder / (uuid4().hex + '.tsv')
        collect(snapshot)
        payload = snapshot.read_bytes()
        catalog = parse_catalog(payload.decode('utf-8-sig'))
        # Failed or partial refreshes never replace the last validated pointer.
        write_json(self.marker, {'context': self.context, 'filename': snapshot.name,
            'fetched_at': self.now().isoformat(), 'sha256': hashlib.sha256(payload).hexdigest(),
            'statistics': catalog['statistics']})
        return catalog

    def enrich(self, bonuses, collect):
        catalog = self.load()
        fresh = catalog is None
        if fresh:
            catalog = self.refresh(collect)
        try:
            result = attach_bulk_owners(bonuses, catalog)
        except (CardNotFound, CardConflict):
            if fresh:
                raise
            catalog = self.refresh(collect)
            result = attach_bulk_owners(bonuses, catalog)
            fresh = True
        if not fresh and result.get('updated_owner_card_count'):
            # A valid weekly snapshot taken after the historical purchases is
            # already the newer source. Refresh only if it predates the checks.
            snapshot = json.loads(self.marker.read_text(encoding='utf-8'))
            fetched_day = datetime.fromisoformat(snapshot['fetched_at']).astimezone(MSK).date().isoformat()
            if any(not r.get('date') or r['date'] >= fetched_day for r in bonuses['rows']):
                catalog = self.refresh(collect)
                result = attach_bulk_owners(bonuses, catalog)
        result['card_catalog_statistics'] = catalog['statistics']
        return result

    def prepare_workbook(self, runtime, base):
        """Build once per validated snapshot; keep relative paths portable."""
        catalog = self.load()
        if catalog is None:
            raise ValueError('A fresh card catalog is required before publication')
        snapshot = json.loads(self.marker.read_text(encoding='utf-8'))
        # Keep prior validated exports for audit when the presentation policy changes.
        format_version = 7
        folder = self.folder / ('export-' + Path(snapshot['filename']).stem + '-v' + str(format_version))
        marker = folder / 'validated.json'
        filename = base['name'] + '_Дисконтные карты.xlsx'
        target = folder / filename
        if marker.exists():
            saved = json.loads(marker.read_text(encoding='utf-8'))
            checked = inspect_xlsx(target)
            if saved['source_sha256'] != snapshot['sha256'] or saved['sha256'] != checked['sha256']:
                raise ValueError('The validated catalog workbook was changed')
            return target
        rows = [{'number': number, 'owner': owner,
                 'note': '; '.join(note for note in (
                     'В 1С различаются имена владельца одного номера' if len(owners - {''}) > 1 else '',
                     'Номер из наименования: в одной из записей не заполнен «Код карты»'
                     if number in catalog['number_from_name'] else '') if note)}
                for number, owners in sorted(catalog['by_number'].items())
                for owner in sorted(owners - {''} or {''})]
        rows.extend({'number': '', 'owner': owner,
                     'note': f'В 1С не указан номер карты. Наименование: {name}. Код записи: {code}'}
                    for code, name, owner in catalog['unnumbered'])
        folder.mkdir(parents=True, exist_ok=True)
        fetched = datetime.fromisoformat(snapshot['fetched_at']).astimezone(MSK)
        pending = folder / 'pending.xlsx'
        write_catalog_workbook(pending, rows, base['name'], fetched)
        extracted, _ = read_first_sheet(pending)
        actual = [r for r in extracted if r.index >= 4 and r.index < 4 + len(rows)]
        if len(actual) != len(rows):
            raise ValueError('Card catalog workbook row count differs from the source')
        for source, output in zip(rows, actual):
            if any(str(output.cells.get(column, '')) != source[key]
                   for column, key in (('A', 'number'), ('B', 'owner'), ('C', 'note'))):
                raise ValueError('Card catalog workbook differs from the source')
        inspect_xlsx(pending)
        pending.replace(target)
        checked = inspect_xlsx(target)
        write_json(marker, {'format_version': format_version, 'source_sha256': snapshot['sha256'], 'sha256': checked['sha256'],
                            'rows': len(rows), 'fetched_at': snapshot['fetched_at']})
        return target


def write_catalog_workbook(target, rows, base_name, fetched):
    """Streaming XLSX fallback: the artifact Windows renderer crashes on large catalogs."""
    import xlsxwriter
    with xlsxwriter.Workbook(str(target), {'constant_memory': True, 'strings_to_numbers': False,
                                         'strings_to_formulas': False, 'strings_to_urls': False}) as workbook:
        sheet = workbook.add_worksheet('Справочник карт')
        sheet.hide_gridlines(2)
        common = {'font_name': 'Arial', 'font_size': 11, 'font_color': '#202B38', 'valign': 'vcenter'}
        title = workbook.add_format(dict(common, font_size=16, bold=True, font_color='#233A53'))
        subtitle = workbook.add_format(common)
        header = workbook.add_format(dict(common, bold=True, font_color='white', bg_color='#233A53',
                                         align='center', text_wrap=True))
        normal = workbook.add_format(dict(common, num_format='@', text_wrap=True))
        alternate = workbook.add_format(dict(common, num_format='@', text_wrap=True, bg_color='#F2F5F8'))
        sheet.set_column('A:A', 23)
        sheet.set_column('B:B', 55)
        sheet.set_column('C:C', 58)
        sheet.set_row(0, 29)
        sheet.set_row(1, 25)
        sheet.set_row(2, 40)
        sheet.write_string(0, 0, 'Справочник карт', title)
        sheet.write_string(1, 0, f'{base_name}, обновлено {fetched:%d.%m.%Y}', subtitle)
        for column, label in enumerate(('Номер карты', 'Владелец карты', 'Примечание')):
            sheet.write_string(2, column, label, header)
        for index, row in enumerate(rows, 3):
            lines = max(1, (len(row['owner']) + 44) // 45, len(row['owner'].splitlines()),
                        (len(row['note']) + 49) // 50, len(row['note'].splitlines()))
            sheet.set_row(index, max(23, lines * 15))
            style = alternate if index % 2 else normal
            for column, key in enumerate(('number', 'owner', 'note')):
                sheet.write_string(index, column, row[key], style)
        sheet.autofilter(2, 0, len(rows) + 2, 2)
        sheet.freeze_panes(3, 0)


def publish_catalog(publisher, source):
    """Weekly replacement retains the ID and rejects edits made on Drive."""
    journal = publisher.journal_path(source.name)
    saved = json.loads(journal.read_text(encoding='utf-8')) if journal.exists() else {}
    previous = saved.get('sha256') if saved.get('status') == 'verified' else None
    return publisher.publish(source, replace_previous_sha=previous)
