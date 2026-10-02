"""Observed 1C Retail 2.3 web controls for the configured report exports.

This module only exports reports and reads forms. It never posts documents,
changes contacts, or accepts browser security warnings.
"""
from __future__ import annotations

import time
import logging
from datetime import date
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from report_tools import inspect_xlsx, copy_checked
from source_extract import missing_card_reference


def merge_duplicate_contacts(records, catalog_codes):
    """Only reuse contact information shared by every matching catalog record."""
    if len(records) < 2 or len(records) != len(catalog_codes) or len(set(catalog_codes)) != len(records):
        raise ValueError('Incomplete duplicate-card verification')
    fields = ('card_number', 'owner', 'phone', 'owner_reference_present')
    expected = tuple(records[0].get(key) for key in fields)
    if any(tuple(record.get(key) for key in fields) != expected for record in records[1:]):
        raise ValueError('Duplicate card identifiers have conflicting contacts')
    return {**records[0], 'duplicate_catalog_codes': catalog_codes,
            'duplicate_contacts_verified': True}


class OneCUI:
    def __init__(self, driver, downloads: Path, card_catalog_format='text'):
        self.driver = driver
        self.downloads = downloads
        if card_catalog_format not in {'text', 'xlsx'}:
            raise ValueError('Unsupported card catalog export format')
        self.card_catalog_format = card_catalog_format
        self._usage_notice_dismissed = False

    def visible(self, selector):
        # 1C replaces controls while a dialog closes. A stale result must be
        # queried again, not mistaken for an absent control or a failed export.
        for attempt in range(3):
            try:
                return [e for e in self.driver.find_elements(By.CSS_SELECTOR, selector)
                        if e.is_displayed()]
            except StaleElementReferenceException:
                if attempt == 2:
                    raise

    def in_front(self, elements):
        # A 1C modal may leave the underlying form geometrically visible.
        # Only the control hit by an ordinary click belongs to the front form.
        return self.driver.execute_script('''
            return arguments[0].filter(e=>{
                const r=e.getBoundingClientRect();
                if(r.width <= 0 || r.height <= 0) return false;
                const hit=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);
                return hit && (hit===e || e.contains(hit));
            });
        ''', elements)

    def one(self, selector, timeout=30):
        def find(_):
            try:
                items = self.visible(selector)
                items = self.in_front(items)
                if len(items) == 1:
                    return items[0]
                self.dismiss_usage_notice()
                return False
            except StaleElementReferenceException:
                return False
        return WebDriverWait(self.driver, timeout).until(find)

    def dismiss_usage_notice(self):
        """Use the ordinary postponement button explicitly approved by the user.

        Never alter licensing state. A recurring notice stops this base.
        """
        titles = self.driver.find_elements(By.XPATH,
            "//*[normalize-space(.)='Проверка правомерности использования конфигурации']")
        if not any(e.is_displayed() for e in titles):
            return False
        buttons = self.driver.find_elements(By.XPATH,
            "//*[self::a or self::button or @role='button'][normalize-space(.)='Повторить позже']")
        buttons = self.in_front([e for e in buttons if e.is_displayed()])
        if len(buttons) != 1:
            return False
        if self._usage_notice_dismissed:
            raise RuntimeError('1C usage notice reappeared after its ordinary postponement')
        self.pointer_click(buttons[0])
        self._usage_notice_dismissed = True
        logging.warning('1C usage notice postponed using its ordinary Repeat later button')
        return True

    def click(self, selector):
        self.pointer_click(self.one(selector))

    def pointer_click(self, element):
        # Element.click() scrolls 1C's custom menus and can move a confirmed
        # visible item under the caption bar. Pointer actions preserve the
        # menu's current, already checked visible position.
        if not self.in_front([element]):
            raise ValueError('The intended click target is no longer in front')
        ActionChains(self.driver).move_to_element(element).click().perform()

    def open_menu(self, button, menu):
        """Retry only opening a menu if a freshly rebuilt page ignored the click."""
        for attempt in range(3):
            if button == '#captionbarMore':
                self.dismiss_cashier_startup()
            if self.in_front(self.visible(menu)):
                return self.one(menu)
            try:
                self.click(button)
                return self.one(menu, 8)
            except TimeoutException:
                if attempt == 2:
                    raise

    def dismiss_cashier_startup(self):
        """Close the observed optional cashier workspace, never a cash shift."""
        titles = [e for e in self.visible('[id$="headerTopLine_title"]')
                  if (e.get_attribute('title') or '').startswith('РМК:')
                  and (e.get_attribute('title') or '').endswith('РМК (управляемый режим)')]
        if not titles:
            return False
        if len(titles) != 1:
            raise ValueError('More than one cashier startup workspace is open')
        title_id = titles[0].get_attribute('id')
        panel = title_id.removesuffix('headerTopLine_title')
        info = self.in_front(self.visible('input[id$="_ДекорацияЗаголовокИнформации_i0"]'))
        if info:
            if len(info) != 1 or info[0].get_attribute('value') != 'Ошибка определения доступных касс ККМ':
                raise ValueError('Unrecognized information dialog; it was not dismissed')
            form = info[0].get_attribute('id').split('_')[0]
            body = self.one('#' + form + '_ДекорацияТекстИнформации_i0').get_attribute('value')
            if body != 'Не смогли определить доступные кассы ККМ.\nОбратитесь к администратору системы.':
                raise ValueError('Unrecognized cashier information; it was not dismissed')
            self.click('#' + form + '_КнопкаЗакрыть')
            WebDriverWait(self.driver, 30).until(lambda _: not self.visible('#' + form + '_КнопкаЗакрыть'))
        # This exact decoration closes the workspace. Closing/opening a shift
        # and terminating the application have different controls and are unused.
        close = self.one('#' + panel + 'formContent [id$="_ДекорацияЗакрыть"]')
        self.pointer_click(close)
        WebDriverWait(self.driver, 30).until(lambda _: not self.visible('#' + title_id))
        return True

    def enter(self, selector, value):
        field = self.one(selector)
        field.click()
        field.send_keys(Keys.CONTROL, 'a')
        field.send_keys(value if value else Keys.BACKSPACE)
        field.send_keys(Keys.TAB)

    def form(self, suffix='_СформироватьОтчет'):
        return self.one('a[id$="' + suffix + '"]').get_attribute('id').split('_')[0]

    def text(self, value, scope=None, timeout=30):
        # Text is a program constant or a value taken from the requested source.
        literal = ('"' + value + '"') if '"' not in value else ("'" + value + "'")
        if '"' in value and "'" in value:
            literal = 'concat(' + ',"\'",'.join("'" + part + "'" for part in value.split("'")) + ')'
        xpath = './/*[normalize-space(translate(text(),"\u00a0"," "))=' + literal + ']'
        def find(_):
            try:
                parent = self.driver.find_element(By.CSS_SELECTOR, scope) if scope else self.driver
                candidates = [e for e in parent.find_elements(By.XPATH, xpath) if e.is_displayed()]
                candidates = self.in_front(candidates)
                return candidates[0] if len(candidates) == 1 else False
            except StaleElementReferenceException:
                return False
        return WebDriverWait(self.driver, timeout).until(find)

    def open_link(self, link, suffix='_СформироватьОтчет'):
        if not link.startswith(('e1cib/data/', 'e1cib/app/', 'e1cib/list/')):
            raise ValueError('Expected an internal 1C navigation link')
        # A keyboard shortcut can be ignored during web-client startup.
        # The observed service-menu command also waits for an interactive UI.
        if not self.in_front(self.visible('input[id$="_CurrentUrlContent_i0"]')):
            self.open_menu('#captionbarMore', '#GotoURLButton')
            self.click('#GotoURLButton')
        field = self.one('input[id$="_CurrentUrlContent_i0"]')
        dialog = field.get_attribute('id').split('_')[0]
        field.clear()
        field.send_keys(link)
        field.send_keys(Keys.TAB)
        WebDriverWait(self.driver, 15).until(lambda _: self.one(
            '#' + dialog + '_CurrentUrlContent_i0').get_attribute('value') == link)
        self.click('#' + dialog + '_GotoUrl')
        WebDriverWait(self.driver, 30).until(lambda _: not self.visible('#' + dialog + '_CurrentUrlContent_i0'))
        return self.form(suffix) if suffix is not None else None

    def card_catalog(self):
        selectors = ('a[id$="_ФормаКоманднаяПанельСписокSearchControl"],'
                     'input[id$="_ФормаКоманднаяПанельСписокSearchString_i0"]')
        def ready(_):
            forms = {e.get_attribute('id').split('_')[0]
                     for e in self.in_front(self.visible(selectors))}
            return next(iter(forms)) if len(forms) == 1 else False
        return WebDriverWait(self.driver, 30).until(ready)

    def open_card_catalog(self, link):
        self.open_link(link, suffix=None)
        return self.card_catalog()

    def menu_action(self, form, suffix, toolbar='КоманднаяПанельФормы', target_id=None):
        """Choose an observed journal menu item, including clipped items."""
        menu_id = form + '_allActions' + toolbar
        scope = '#' + menu_id + '_submenu'
        menu = self.open_menu('#' + menu_id, scope)
        target = '#' + (target_id or form + '_popup_' + suffix)
        for _ in range(80):
            # A clipped item is present but Selenium calls it not displayed.
            # Locate it to choose the scroll direction, then require a real
            # unobstructed hit target before performing the pointer click.
            candidates = self.driver.find_elements(By.CSS_SELECTOR, target)
            if len(candidates) == 1 and self.in_front(candidates):
                self.pointer_click(candidates[0])
                return
            if len(candidates) != 1:
                raise ValueError('The requested journal command is missing or ambiguous')
            above = self.driver.execute_script('return arguments[0].getBoundingClientRect().top < arguments[1].getBoundingClientRect().top;', candidates[0], menu)
            arrows = self.in_front(self.visible(scope + (' > .jumpBack > i' if above else ' > .jumpFwrd > i')))
            if len(arrows) == 1:
                self.pointer_click(arrows[0])
            time.sleep(.15)
        # The command exists, but startup notifications or rendering can
        # temporarily cover it. Let the bounded transport/UI retry reopen 1C.
        raise TimeoutError('The requested journal menu item is not reachable')

    def report_dates_and_filters(self, form, day: date, end_day: date | None = None):
        # A saved variant can inherit the user's current shop. The first
        # setting enables the period itself, not a shop/product filter.
        selector = '[id^="' + form + '_Компоновщик"][id$="Использование"]'
        period_id = form + '_КомпоновщикНастроекПользовательскиеНастройкиЭлемент0Использование'
        period = self.one('#' + period_id)
        start = self.one('#' + form + '_Период1ДатаНачала_i0')
        if not (0 < start.rect['x'] - period.rect['x'] < 120
                and abs(start.rect['y'] - period.rect['y']) < 15):
            raise ValueError('The period checkbox is not beside the date field')
        filter_ids = [e.get_attribute('id') for e in self.visible(selector)
                      if e.get_attribute('id') != period_id]
        for control_id in filter_ids:
            control = self.one('#' + control_id)
            if 'select' in control.get_attribute('class').split():
                control.click()
        for part, requested in (('Начала', day), ('Окончания', end_day or day)):
            selector_date = '#' + form + '_Период1Дата' + part + '_i0'
            self.enter(selector_date, requested.strftime('%d%m%Y'))
            if self.one(selector_date).get_attribute('value') != requested.strftime('%d.%m.%Y'):
                raise ValueError('1C did not accept the requested date')
        if 'select' not in self.one('#' + period_id).get_attribute('class').split():
            self.click('#' + period_id)
        if any('select' in self.one('#' + control_id).get_attribute('class').split()
               for control_id in filter_ids):
            raise ValueError('A report quick filter is still active')

    def generate(self, form, timeout=180):
        document_id = form + '_ОтчетТабличныйДокумент'
        # Require a changed/new report so an old render cannot be accepted.
        def pending(_):
            return self.driver.execute_script(
                "return document.getElementById(arguments[0])?.innerText.includes('Нажмите')", document_id)
        WebDriverWait(self.driver, 30).until(pending)
        self.click('#' + form + '_СформироватьОтчет')
        def ready(_):
            return self.driver.execute_script("""
                const e = document.getElementById(arguments[0]);
                // Native iframe controls also exist while the old report is
                // covered by "Отчет формируется...". Wait for both the status
                // text and its observed blocking surfaces to disappear.
                return e && e.innerText.trim() === '' && e.querySelector('iframe') !== null
                    && [...e.querySelectorAll('.stateWindowSupport,.stateWindowSupportSurface')]
                        .every(s => getComputedStyle(s).display === 'none'
                            || getComputedStyle(s).visibility === 'hidden');
            """, document_id)
        WebDriverWait(self.driver, timeout).until(ready)

    def export_xlsx(self, target: Path, timeout=180):
        before = self.download_snapshot()
        self.click('span[id$="headerTopLine_cmd_SaveButton"]')
        def picker(_):
            body = self.driver.execute_script('return document.body.innerText')
            return ('Лист Microsoft Excel 2007 (.xlsx)' in body
                    or any(e.text.replace('\u00a0',' ') == 'Продолжить без установки'
                           for e in self.in_front(self.visible('a[id$="_ПродолжитьБезУстановки"]'))))
        WebDriverWait(self.driver, 30).until(picker)
        continuation = [e for e in self.in_front(self.visible('a[id$="_ПродолжитьБезУстановки"]'))
                        if e.text.replace('\u00a0',' ') == 'Продолжить без установки']
        if continuation:
            continuation[0].click()
        self.text('Лист Microsoft Excel 2007 (.xlsx)').click()
        self.text('ОК').click()
        return self.wait_xlsx(before, target, timeout)

    def download_snapshot(self):
        result = {}
        for path in self.downloads.iterdir():
            try:
                if path.is_file():
                    stat = path.stat()
                    result[path.name] = (stat.st_size, stat.st_mtime_ns)
            except FileNotFoundError:
                # Chrome renames its temporary file on completion.
                continue
        return result

    def export_native_spreadsheet_xlsx(self, target: Path, timeout=300):
        # A standalone native document's Save button downloads MXL directly.
        # Its observed File -> Save as dialog has a separate format selector.
        self.click('span[id$="headerTopLine_cmd_MoreButton"]')
        self.click('#FileBand')
        self.pointer_click(self.text('Сохранить как...'))
        kind = self.one('input[id$="_FileType_i0"]')
        dialog = kind.get_attribute('id').split('_')[0]
        if not kind.get_attribute('value').endswith('(*.xlsx)'):
            kind.send_keys(Keys.ARROW_DOWN)
            self.pointer_click(self.text('Лист Excel2007-... (*.xlsx)'))
        WebDriverWait(self.driver, 20).until(lambda _: self.driver.find_element(
            By.ID, dialog + '_FileType_i0').get_attribute('value').endswith('(*.xlsx)'))
        before = self.download_snapshot()
        self.click('#' + dialog + '_OK')
        return self.wait_xlsx(before, target, timeout)

    def wait_xlsx(self, before, target: Path, timeout=180):
        deadline, previous, stable = time.monotonic() + timeout, None, None
        while time.monotonic() < deadline:
            changed = []
            for path in self.downloads.iterdir():
                if not path.is_file():
                    continue
                try:
                    stat = path.stat()
                except FileNotFoundError:
                    continue
                state = (stat.st_size, stat.st_mtime_ns)
                if before.get(path.name) != state:
                    changed.append((path, state))
            complete = [(p, s) for p, s in changed if p.suffix.lower() == '.xlsx']
            partial = [p for p, _ in changed if p.suffix.lower() in {'.crdownload', '.tmp', '.part'}]
            if len(complete) > 1:
                raise ValueError('More than one new XLSX appeared during export')
            if len(complete) == 1 and not partial:
                path, state = complete[0]
                key = (path.name, state)
                if key == previous and time.monotonic() - stable >= 2:
                    inspect_xlsx(path)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    result = copy_checked(path, target)
                    return result
                if key != previous:
                    previous, stable = key, time.monotonic()
            else:
                previous = stable = None
            time.sleep(.5)
        raise TimeoutError('Chrome did not finish the report download')

    def collect_report(self, link, day: date, target: Path, end_day: date | None = None):
        form = self.open_link(link)
        self.report_dates_and_filters(form, day, end_day)
        self.generate(form, timeout=600 if end_day else 180)
        result = self.export_xlsx(target, timeout=600 if end_day else 180)
        result.update(report_date=day.isoformat(), navigation_link=link)
        if end_day:
            result['report_end_date'] = end_day.isoformat()
        self.click('span[id$="headerTopLine_cmd_CloseButton"]')
        return result

    def ensure_check_operations(self, form):
        """Enable the observed personal journal column without editing checks."""
        try:
            self.text('Операция', '#' + form + '_Список', timeout=1)
            return
        except TimeoutException:
            pass
        self.menu_action(form, 'ФормаИзменитьФорму')
        settings = self.one('div[id^="form"][id$="_ItemsTable"]').get_attribute('id').split('_')[0]
        grid = '#' + settings + '_ItemsTable'
        self.pointer_click(self.text('Список', grid))
        ActionChains(self.driver).send_keys(Keys.ARROW_RIGHT).perform()
        for page in range(8):
            try:
                item = self.text('Операция', grid, timeout=2)
                row = self.driver.execute_script('return arguments[0].closest(".gridLine")', item)
                checkbox = row.find_element(By.CSS_SELECTOR, '.checkbox')
                if 'select' not in checkbox.get_attribute('class').split():
                    self.pointer_click(checkbox)
                    def selected(_):
                        item = self.text('Операция', grid, timeout=2)
                        row = self.driver.execute_script('return arguments[0].closest(".gridLine")', item)
                        return 'select' in row.find_element(By.CSS_SELECTOR, '.checkbox').get_attribute('class').split()
                    WebDriverWait(self.driver, 15, ignored_exceptions=(StaleElementReferenceException,)).until(selected)
                break
            except TimeoutException:
                if page == 7:
                    raise ValueError('The check operation column is absent from personal form settings')
                ActionChains(self.driver).send_keys(Keys.PAGE_DOWN).perform()
        # Some native dialogs place their footer under a resize handle.
        # Normal keyboard activation of the observed default OK button works.
        self.driver.find_element(By.CSS_SELECTOR, '#' + settings + '_OK a').send_keys(Keys.ENTER)
        WebDriverWait(self.driver, 30).until(lambda _: not self.visible(grid))

    def collect_check_list(self, link, day: date, target: Path, end_day: date | None = None):
        form = self.open_link(link, '_allActionsКоманднаяПанельФормы')
        self.menu_action(form, 'ФормаСтандартныеНастройкиДинамическогоСписка')
        self.ensure_check_operations(form)
        for suffix in ('_ОтборМагазин_i0', '_ОтборОрганизация_i0',
                       '_ОтборКассаККМ_i0', '_КоманднаяПанельФормыСписокSearchString_i0'):
            # Sever-Yug's native journal exposes no organization selector or
            # organization column. All exported checks must still reconcile
            # with the independent receipt report for every shop.
            if suffix == '_ОтборОрганизация_i0' and not self.visible('#' + form + suffix):
                continue
            control = self.one('#' + form + suffix)
            if control.get_attribute('value'):
                self.enter('#' + form + suffix, '')
        self.menu_action(form, 'ФормаУстановитьИнтервал')
        period = self.one('input[id$="_DateBegin_i0"]').get_attribute('id').split('_')[0]
        for suffix, requested in (('_DateBegin_i0', day), ('_DateEnd_i0', end_day or day)):
            self.enter('#' + period + suffix, requested.strftime('%d%m%Y'))
            if self.one('#' + period + suffix).get_attribute('value') != requested.strftime('%d.%m.%Y'):
                raise ValueError('The check journal did not accept the requested date')
        self.click('#' + period + '_Select')
        WebDriverWait(self.driver, 30).until(lambda _: not self.visible('#' + period + '_Select'))
        self.menu_action(form, 'ФормаВывестиСписок')
        kind = self.one('input[id$="_DocumentTypeElement_i0"]')
        dialog = kind.get_attribute('id').removeprefix('input_').split('_')[0]
        if kind.get_attribute('value') != 'Текстовый документ':
            kind.send_keys(Keys.ARROW_DOWN)
            self.text('Текстовый документ').click()
        self.click('#' + dialog + '_CheckAll')
        only_selected = self.one('#' + dialog + '_OnlySelectedElement')
        if 'select' in only_selected.get_attribute('class').split():
            only_selected.click()
        self.click('#' + dialog + '_Ok')
        document = self.one('textarea[id$="_TextDocument"]', 600 if end_day else 180)
        text = document.get_attribute('value')
        if not text.startswith('Дата\t'):
            raise ValueError('Unexpected native check list text layout')
        if 'Операция' not in text.splitlines()[0].split('\t'):
            raise ValueError('Native check export omitted the operation column')
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open('x', encoding='utf-8', newline='') as stream:
            stream.write(text)
        self.click('span[id$="headerTopLine_cmd_CloseButton"]')
        return text

    def collect_card_list(self, link, target: Path):
        """Export all catalog columns and child groups through native Output list."""
        previous_timeout = self.driver.timeouts.page_load
        previous_script_timeout = self.driver.timeouts.script
        transport = self.driver.command_executor.client_config
        previous_transport_timeout = transport.timeout
        transport.timeout = 660
        self.driver.set_page_load_timeout(600)
        self.driver.set_script_timeout(600)
        try:
            return self._collect_card_list(link, target)
        finally:
            # Restore the local transport even if the browser has disconnected.
            transport.timeout = previous_transport_timeout
            try:
                transport.timeout = 3
                self.driver.set_page_load_timeout(previous_timeout)
                self.driver.set_script_timeout(previous_script_timeout)
            except Exception:
                logging.warning('card_catalog could not restore browser page timeout')
            finally:
                transport.timeout = previous_transport_timeout

    def _collect_card_list(self, link, target: Path):
        logging.info('card_catalog stage=open')
        form = self.open_card_catalog(link)
        toolbar = 'ФормаКоманднаяПанель'
        self.menu_action(form, 'ФормаСтандартныеНастройкиДинамическогоСписка', toolbar)
        self.enter('#' + form + '_ФормаКоманднаяПанельСписокSearchString_i0', '')
        # Flat list includes every group, regardless of the previous folder.
        self.menu_action(form, '', toolbar, target_id=form + '_ФормаРежимПросмотра_all_sub')
        self.click('#' + form + '_ФормаСписок')
        logging.info('card_catalog stage=flat_list')
        self.menu_action(form, 'ФормаВывестиСписок', toolbar)
        kind = self.one('input[id$="_DocumentTypeElement_i0"]')
        dialog = kind.get_attribute('id').split('_')[0]
        wanted_kind = 'Табличный документ' if self.card_catalog_format == 'xlsx' else 'Текстовый документ'
        if kind.get_attribute('value') != wanted_kind:
            kind.send_keys(Keys.ARROW_DOWN)
            self.text(wanted_kind).click()
        self.click('#' + dialog + '_CheckAll')
        for suffix, wanted in (('_OnlySelectedElement', False), ('_WithChildrenElement', True)):
            if suffix == '_WithChildrenElement' and not self.visible('#' + dialog + suffix):
                # The flat-list dialog has no child-group selector.
                continue
            control = self.one('#' + dialog + suffix)
            if ('select' in control.get_attribute('class').split()) != wanted:
                self.pointer_click(control)
        # Rendering a catalog with tens of thousands of rows can block 1C's
        # browser renderer longer than the normal page-navigation timeout.
        logging.info('card_catalog stage=generate_%s', self.card_catalog_format)
        self.click('#' + dialog + '_Ok')
        # Do not run a whole-document text XPath while a large native list
        # renders. Query only the expected control until the document is ready.
        def text_document_ready(_):
            documents = self.in_front(self.visible('textarea[id$="_TextDocument"]'))
            return documents[0] if len(documents) == 1 else False
        if self.card_catalog_format == 'xlsx':
            self.one('span[id$="headerTopLine_cmd_SaveButton"]', 600)
            native = target.with_suffix('.native.xlsx')
            self.export_native_spreadsheet_xlsx(native)
            from card_catalog import catalog_table_to_text
            from source_extract import read_first_sheet
            from report_tools import write_json
            rows, provenance = read_first_sheet(native)
            text = catalog_table_to_text(rows)
            write_json(target.with_suffix('.source.json'), provenance)
        else:
            document = WebDriverWait(self.driver, 600).until(text_document_ready)
            text = document.get_attribute('value')
        logging.info('card_catalog stage=text_received bytes=%s', len(text.encode('utf-8')))
        from card_catalog import parse_catalog
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open('x', encoding='utf-8', newline='') as stream:
            stream.write(text)
        self.click('span[id$="headerTopLine_cmd_CloseButton"]')
        parse_catalog(text)
        return text

    def card_owner_field(self, timeout=45):
        """Dismiss only 1C's optional workstation offer; never install/create it."""
        def ready(_):
            try:
                fields = self.in_front(self.visible('input[id$="_ВладелецКарты_i0"]'))
                if len(fields) == 1:
                    return fields[0]
                install = self.in_front(self.visible('a[id$="_УстановитьИПродолжить"]'))
                create = self.in_front(self.visible('a[id$="_ПродолжитьБезУстановки"]'))
                if (len(install) == len(create) == 1
                        and install[0].text == 'Установить и продолжить'
                        and create[0].text == 'Создать рабочее место без установки компоненты'
                        and install[0].get_attribute('id').split('_')[0]
                        == create[0].get_attribute('id').split('_')[0]):
                    close = self.in_front(self.visible('span[id$="headerTopLine_cmd_CloseButton"]'))
                    if len(close) == 1:
                        close[0].click()
                return False
            except StaleElementReferenceException:
                return False
        return WebDriverWait(self.driver, timeout).until(ready)

    def lookup_owner_phone(self, card_number, card_presentation, _catalog_code=None):
        if missing_card_reference(card_presentation):
            return {'owner': '', 'phone': '', 'card_presentation': card_presentation,
                    'card_reference_present': False, 'owner_reference_present': False,
                    'source': 'Журнал чеков 1С: объект дисконтной карты не найден'}
        catalog = self.card_catalog()
        number_only = card_number is not None and card_presentation == card_number
        search_value = _catalog_code or card_number or card_presentation
        # 1C accumulates searches on different columns. Clear the previous
        # search before switching between card name and card code.
        menu = '#' + catalog + '_allActionsФормаКоманднаяПанель'
        self.click(menu)
        self.one(menu + '_submenu')
        cancel = self.in_front(self.visible('#' + catalog + '_popup_ФормаОтменитьПоиск'))
        if len(cancel) == 1 and 'disable' not in cancel[0].get_attribute('class').lower():
            cancel[0].click()
        else:
            ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
        WebDriverWait(self.driver, 15).until(lambda _: not self.visible(menu + '_submenu'))
        search = self.in_front(self.visible('#' + catalog + '_ФормаКоманднаяПанельСписокSearchControl'))
        if search:
            search_button = '#' + catalog + '_ФормаКоманднаяПанельСписокSearchControl'
            self.open_menu(search_button, search_button + '_submenu')
            self.pointer_click(self.text('Расширенный поиск', scope=search_button + '_submenu'))
        else:
            menu = '#' + catalog + '_allActionsФормаКоманднаяПанель'
            self.click(menu)
            self.text('Расширенный поиск', scope=menu + '_submenu').click()
        field = self.one('input[id$="_Pattern_i0"]')
        dialog = field.get_attribute('id').split('_')[0]
        desired_field = 'Код' if _catalog_code else ('Код карты' if number_only else 'Наименование')
        selector = self.one('#' + dialog + '_FieldSelector_i0')
        if selector.get_attribute('value') != desired_field:
            selector.send_keys(Keys.ARROW_DOWN)
            self.text(desired_field).click()
        if self.one('#' + dialog + '_FieldSelector_i0').get_attribute('value') != desired_field:
            raise ValueError('The card search field was not accepted')
        if _catalog_code or number_only or card_number is None:
            self.text('По точному совпадению').click()
            current_group = self.one('#' + dialog + '_OnlyCurrentGroup')
            if 'select' in current_group.get_attribute('class').split():
                current_group.click()
        self.enter('#' + dialog + '_Pattern_i0', search_value)
        self.click('#' + dialog + '_Find')
        WebDriverWait(self.driver, 30).until(lambda _: not self.visible('#' + dialog + '_Find'))
        if _catalog_code or number_only or card_number is None:
            def read_column(_):
                header = self.text(desired_field, '#' + catalog + '_Список')
                return self.driver.execute_script("return arguments[0].closest('[colindex]')?.getAttribute('colindex')", header)
            column = WebDriverWait(self.driver, 30, ignored_exceptions=(StaleElementReferenceException,)).until(read_column)
            if column is None or not column.isdigit():
                raise ValueError('Cannot identify the card-code column')
            last_match, stable_since = None, None
            def exact_card(_):
                nonlocal last_match, stable_since
                try:
                    selector = '#' + catalog + '_Список .gridBox[colindex="' + column + '"] .gridBoxText'
                    matches = [e for e in self.in_front(self.visible(selector)) if e.text == search_value]
                    if len(matches) > 1:
                        if _catalog_code:
                            raise ValueError('Internal catalog code is not unique')
                        return matches
                    if not matches:
                        last_match, stable_since = None, None
                        return False
                    if matches[0].id != last_match:
                        last_match, stable_since = matches[0].id, time.monotonic()
                        return False
                    return matches[0] if time.monotonic() - stable_since >= .5 else False
                except StaleElementReferenceException:
                    last_match, stable_since = None, None
                    return False
            item = WebDriverWait(self.driver, 30).until(exact_card)
            if isinstance(item, list):
                header = self.text('Код', '#' + catalog + '_Список')
                code_column = self.driver.execute_script("return arguments[0].closest('[colindex]').getAttribute('colindex')", header)
                def read_codes(_):
                    selector = '#' + catalog + '_Список .gridBox[colindex="' + column + '"] .gridBoxText'
                    return self.driver.execute_script("return [...document.querySelectorAll(arguments[0])].filter(e=>e.offsetWidth && e.innerText===arguments[1]).map(e=>e.closest('.gridLine').querySelector('.gridBox[colindex=\"'+arguments[2]+'\"] .gridBoxText').innerText.trim())", selector, search_value, code_column) or False
                codes = WebDriverWait(self.driver, 30).until(read_codes)
                if len(codes) < 2:
                    raise ValueError('Duplicate catalog changed during verification; retry required')
                if any(not code for code in codes) or len(set(codes)) != len(codes):
                    raise ValueError('Cannot identify duplicate catalog records')
                records = [self.lookup_owner_phone(card_number, card_presentation, _catalog_code=code) for code in codes]
                return merge_duplicate_contacts(records, codes)
        else:
            item = self.text(card_presentation, '#' + catalog + '_Список')
        # The filtered grid may rebuild or ignore its first double click.
        # Retry only this read-only opening with a fresh, unobstructed match.
        for attempt in range(3):
            try:
                ActionChains(self.driver).double_click(item).perform()
                owner_field = self.card_owner_field(timeout=15)
                break
            except (StaleElementReferenceException, TimeoutException):
                if attempt == 2:
                    raise
                def recovered(_):
                    # The server may open the card after the first wait ends.
                    # Accept that completed opening instead of looking only
                    # for the now-covered catalog row and timing out again.
                    fields = self.in_front(self.visible('input[id$="_ВладелецКарты_i0"]'))
                    if len(fields) == 1:
                        return fields[0], None
                    if number_only or card_number is None:
                        fresh = exact_card(_)
                    else:
                        try:
                            fresh = self.text(card_presentation, '#' + catalog + '_Список', timeout=.5)
                        except TimeoutException:
                            return False
                    return (None, fresh) if fresh else False
                owner_field, item = WebDriverWait(self.driver, 30,
                    ignored_exceptions=(StaleElementReferenceException,)).until(recovered)
                if owner_field is not None:
                    break
        card_form = owner_field.get_attribute('id').split('_')[0]
        if number_only:
            if self.one('#' + card_form + '_КодКарты_i0').get_attribute('value') != card_number:
                raise ValueError('The opened card code does not match the receipt')
        elif self.one('#' + card_form + '_Наименование_i0').get_attribute('value') != card_presentation:
            raise ValueError('The opened card does not match the receipt')
        if card_number is None:
            card_number = self.one('#' + card_form + '_КодКарты_i0').get_attribute('value')
            if not card_number or not card_number.strip():
                raise ValueError('The named card has no card code')
        owner = owner_field.get_attribute('value')
        if not owner:
            result = {'owner': '', 'phone': '', 'card_number': card_number,
                      'card_presentation': card_presentation,
                      'owner_reference_present': False,
                      'source': 'Карточка карты: ссылка на владельца отсутствует'}
            self.click('span[id$="headerTopLine_cmd_CloseButton"]')
            self.card_catalog()
            return result
        self.click('#' + card_form + '_ВладелецКарты_OB')
        def arrived(_):
            try:
                fields = self.visible('input[id$="_Наименование_i0"]')
                if len(fields) > 1:
                    fields = self.in_front(fields)
                if len(fields) == 1 and fields[0].get_attribute('id').split('_')[0] != card_form:
                    return fields[0]
                return False
            except StaleElementReferenceException:
                return False
        owner_name = WebDriverWait(self.driver, 30).until(arrived)
        owner_form = owner_name.get_attribute('id').split('_')[0]
        if owner_name.get_attribute('value') != owner:
            raise ValueError('Owner name does not match the opened reference')
        self.click('#thpage_' + owner_form + '_СтраницаКонтактнаяИнформация')
        def read_phones(_):
            return self.driver.execute_script('''
                return [...document.querySelectorAll('label[for]')]
                    .filter(e=>e.offsetWidth && /^Телефон[ы]?:$/.test(e.innerText.trim())
                        && e.htmlFor.startsWith(arguments[0]+'_'))
                    .map(e=>({label:e.innerText,field:e.htmlFor,
                        value:document.getElementById(e.htmlFor)?.value}));
            ''', owner_form) or False
        phones = WebDriverWait(self.driver, 30).until(read_phones)
        if any(p['value'] is None for p in phones):
            raise ValueError('Phone control cannot be read')
        result = {'owner': owner, 'owner_reference_present': True, 'card_number': card_number,
                  'card_presentation': card_presentation,
                  'phone': '; '.join(dict.fromkeys(p['value'].strip() for p in phones if p['value'].strip())),
                  'source': 'Владелец карты → Контактная информация', 'phone_fields': phones}
        self.click('span[id$="headerTopLine_cmd_CloseButton"]')
        WebDriverWait(self.driver, 30).until(lambda _: bool(self.visible('#' + card_form + '_ВладелецКарты_i0')))
        self.click('span[id$="headerTopLine_cmd_CloseButton"]')
        self.card_catalog()
        return result
