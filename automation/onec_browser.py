"""Dedicated Chrome lifecycle and credential loading for the nightly exporter.

Selenium 4.49.0. Credentials are read from environment variables or the
user-provided login_test.py as literal data, never by importing that script.
"""
from __future__ import annotations

import ast
import os
import time
from dataclasses import dataclass, field
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from report_tools import inspect_xlsx


@dataclass(frozen=True)
class Credentials:
    login: str
    password: str = field(repr=False)


def load_credentials(legacy_test: Path | None = None) -> Credentials:
    login = os.environ.get('ONEC_LOGIN')
    password = os.environ.get('ONEC_PASSWORD')
    if (login is None) != (password is None):
        raise ValueError('Set both ONEC_LOGIN and ONEC_PASSWORD')
    if login is None:
        if legacy_test is None:
            raise ValueError('Credentials source is required')
        values = {}
        tree = ast.parse(legacy_test.read_text(encoding='utf-8-sig'))
        for node in tree.body:
            if isinstance(node, ast.Assign) and isinstance(node.value, ast.Constant):
                for target in node.targets:
                    if isinstance(target, ast.Name) and target.id in {'LOGIN', 'PASSWORD'}:
                        values[target.id] = node.value.value
        login, password = values.get('LOGIN'), values.get('PASSWORD')
    for value in (login, password):
        if not isinstance(value, str) or not value.strip():
            raise ValueError('Login and password must be nonempty strings')
        if value.casefold().startswith(('твой_', 'ваш_', 'your_')):
            raise ValueError('Credentials contain a placeholder')
    return Credentials(login, password)


def wait_completed_download(folder: Path, timeout: float = 180, poll: float = 1) -> Path:
    """Use a fresh folder per export. Never rename or read a partial download.

    A browser warning must be handled by the user. This function only waits
    for the browser to finish and verifies the resulting XLSX container.
    """
    deadline = time.monotonic() + timeout
    previous = None
    stable_since = None
    while time.monotonic() < deadline:
        files = [p for p in folder.iterdir() if p.is_file()]
        complete = [p for p in files if p.suffix.lower() == '.xlsx']
        temporary = [p for p in files if p.suffix.lower() in {'.crdownload', '.tmp', '.part'}]
        if len(complete) > 1:
            raise ValueError('More than one XLSX appeared in the export folder')
        if len(complete) == 1 and not temporary:
            path = complete[0]
            try:
                state = (path.name, path.stat().st_size, path.stat().st_mtime_ns)
            except FileNotFoundError:
                previous = stable_since = None
                continue
            if state == previous and stable_since is not None and time.monotonic() - stable_since >= 2:
                inspect_xlsx(path)
                return path
            if state != previous:
                previous, stable_since = state, time.monotonic()
        else:
            previous = stable_since = None
        time.sleep(poll)
    raise TimeoutError('Chrome did not finish the download; check its download panel')


class OneCBrowser:
    def __init__(self, profile: Path, downloads: Path, cache: Path):
        for folder in (profile, downloads, cache):
            folder.mkdir(parents=True, exist_ok=True)
        os.environ['SE_CACHE_PATH'] = str(cache.resolve())
        os.environ['SE_AVOID_STATS'] = 'true'
        options = webdriver.ChromeOptions()
        options.add_argument('--window-size=1400,1000')
        options.add_argument('--user-data-dir=' + str(profile.resolve()))
        options.add_experimental_option('excludeSwitches', ['disable-client-side-phishing-detection'])
        options.add_experimental_option('prefs', {
            'download.default_directory': str(downloads.resolve()),
            'download.prompt_for_download': False,
            'credentials_enable_service': False,
            'profile.password_manager_enabled': False,
        })
        self.driver = webdriver.Chrome(options=options)
        self.driver.set_page_load_timeout(90)
        self.wait = WebDriverWait(self.driver, 90, ignored_exceptions=(StaleElementReferenceException,))

    def visible(self, css: str):
        return [e for e in self.driver.find_elements(By.CSS_SELECTOR, css) if e.is_displayed()]

    def one(self, css: str):
        found = self.visible(css)
        if len(found) != 1:
            raise ValueError(f'Expected exactly one visible control: {css}')
        return found[0]

    def enter(self, css: str, value: str):
        control = self.one(css)
        control.click()
        # Native login inputs support WebDriver clear(). Ctrl+A was observed
        # to insert a literal "a" after reconnecting; verify the exact value
        # before submitting credentials, without logging either field.
        control.clear()
        if value:
            control.send_keys(value)
        control.send_keys(Keys.TAB)
        self.wait.until(lambda _: self.one(css).get_attribute('value') == value)

    def login(self, url: str, credentials: Credentials) -> str:
        self.driver.get(url)
        def page_ready(_):
            ready = self.visible('#authWindow_basic_login') or self.visible('#themesCell_theme_3')
            if ready:
                return ready
            text = self.driver.execute_script('return document.body?.innerText || ""')
            if 'HTTP ERROR 503' in text and 'Страница недоступна' in text:
                raise ConnectionError('1C server returned HTTP 503')
            return False

        self.wait.until(page_ready)
        if self.visible('#themesCell_theme_3'):
            return self.driver.title
        self.enter('#authWindow_basic_login', credentials.login)
        self.enter('#authWindow_basic_password', credentials.password)
        if (self.one('#authWindow_basic_login').get_attribute('value') != credentials.login
                or self.one('#authWindow_basic_password').get_attribute('value') != credentials.password):
            raise ValueError('The login form did not retain the exact credentials')
        self.one('#authWindow_basic_okButton').click()
        rejected = 'Неверно указан пользователь или пароль'

        def result(_):
            if rejected in self.driver.execute_script('return document.body.innerText'):
                raise PermissionError('1C rejected the supplied credentials')
            return self.visible('#themesCell_theme_3')

        self.wait.until(result)
        return self.driver.title

    def close(self):
        # An unresponsive renderer must not add several minutes of cleanup.
        self.driver.command_executor.client_config.timeout = 5
        self.driver.quit()
