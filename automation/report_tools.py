"""Планирование, проверка целостности и безопасное сохранение выгрузок.

Python 3.11+. Команды plan/verify работают без внешних библиотек.
publish требует google-api-python-client и google-auth.
Ни одна команда не меняет исходный файл и не создаёт расписание.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
MSK = timezone(timedelta(hours=3), "Europe/Moscow")
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def read_config(path: Path) -> dict:
    cfg = json.loads(path.read_text(encoding="utf-8-sig"))
    ids = [b["id"] for b in cfg["bases"]]
    names = [safe_name(b["name"]) for b in cfg["bases"]]
    if len(ids) != len(set(ids)) or len(names) != len(set(names)):
        raise ValueError("Идентификаторы и имена баз должны быть уникальны")
    return cfg


def safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value).strip(" .")
    if not value or value in {".", ".."}:
        raise ValueError("Пустое или недопустимое имя")
    if value.upper().split(".")[0] in {"CON", "PRN", "AUX", "NUL", *[f"COM{x}" for x in range(1, 10)], *[f"LPT{x}" for x in range(1, 10)]}:
        value = "_" + value
    return value


def report_day(value: str | None = None, now: datetime | None = None) -> date:
    if value:
        return date.fromisoformat(value)
    return ((now or datetime.now(MSK)).astimezone(MSK) - timedelta(days=1)).date()


def period_key(day: date, end_day: date | None = None) -> str:
    if end_day is not None and end_day < day:
        raise ValueError('Report period ends before it starts')
    return day.isoformat() + ('_' + end_day.isoformat() if end_day is not None else '')


def period_label(day: date, end_day: date | None = None) -> str:
    period_key(day, end_day)
    return day.strftime('%d.%m.%Y') + ('-' + end_day.strftime('%d.%m.%Y') if end_day is not None else '')


def report_folder(cfg, day, end_day=None):
    if end_day is not None:
        return period_label(day, end_day) + '_отчёты'
    return day.strftime(cfg.get('drive_daily_folder_pattern', '%d.%m.%Y_отчёты'))


def plan(cfg: dict, day: date, end_day: date | None = None) -> list[dict]:
    return [
        {"key": f"{period_key(day, end_day)}:{base['id']}:{kind}", "date": day.isoformat(),
         **({'end_date': end_day.isoformat()} if end_day is not None else {}),
         **({'drive_month_folder': day.strftime('%Y.%m_отчёты')} if cfg.get('drive_month_folders') else {}),
         "base_id": base["id"], "base_name": base["name"], "report_type": kind,
         "filename": f"{period_label(day, end_day)}_{safe_name(base['name'])}_{kind}.xlsx",
         "drive_subfolder": report_folder(cfg, day, end_day)
                            if cfg.get('drive_daily_folders') else None,
         "login_verified": base["login_verified"]}
        for base in cfg["bases"] if base.get("enabled", True) for kind in cfg["reports"]
    ]


def inspect_xlsx(path: Path) -> dict:
    """Проверка контейнера и всех XML. Никакой код из книги не выполняется."""
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        if not {"[Content_Types].xml", "xl/workbook.xml"} <= names:
            raise ValueError("Это не книга XLSX")
        if sum(i.file_size for i in archive.infolist()) > 512 * 1024 * 1024:
            raise ValueError("Распакованная книга превышает лимит 512 МБ")
        if any(n.endswith("vbaProject.bin") for n in names):
            raise ValueError("Макросы не допускаются в отчётах XLSX")
        bad = archive.testzip()
        if bad:
            raise ValueError(f"Не совпала контрольная сумма части книги: {bad}")
        for name in names:
            if name.endswith((".xml", ".rels")):
                parsed = ET.fromstring(archive.read(name))
                if name.startswith('xl/worksheets/') and any(
                        element.tag.endswith('}c') and element.attrib.get('t') == 'e'
                        and element.findtext('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v') not in (None, '')
                        for element in parsed.iter()):
                    raise ValueError('Spreadsheet error cell in XLSX; publication is unsafe')
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        wb = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = [s.attrib["name"] for s in wb.findall("s:sheets/s:sheet", ns)]
        if not sheets or not any(n.startswith("xl/worksheets/") and n.endswith(".xml") for n in names):
            raise ValueError("В книге нет листов")
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": digest.hexdigest(), "sheets": sheets}


def copy_checked(source: Path, destination: Path) -> dict:
    if source.suffix.lower() != ".xlsx":
        raise ValueError("Браузер должен завершить скачивание и выдать файл .xlsx")
    if destination.suffix.lower() != ".xlsx":
        raise ValueError("Выходной файл должен иметь расширение .xlsx")
    before = (source.stat().st_size, source.stat().st_mtime_ns)
    original = inspect_xlsx(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        current = inspect_xlsx(destination)
        if current["sha256"] == original["sha256"]:
            return current | {"unchanged": True}
        raise FileExistsError("Файл с таким именем уже существует и отличается")
    with tempfile.NamedTemporaryFile(dir=destination.parent, suffix=".part", delete=False) as tmp:
        staged = Path(tmp.name)
    try:
        shutil.copyfile(source, staged)
        copied = inspect_xlsx(staged)
        if before != (source.stat().st_size, source.stat().st_mtime_ns) or copied["sha256"] != original["sha256"]:
            raise ValueError("Файл изменился во время копирования")
        # Publish the complete local file atomically. Creating a hard link on
        # the same filesystem fails if the destination already exists; a
        # interrupted copy cannot leave a half-written final .xlsx behind.
        os.link(staged, destination)
    finally:
        staged.unlink(missing_ok=True)
    return inspect_xlsx(destination)


def replace_checked(source: Path, destination: Path, expected_previous_sha: str, archive: Path) -> dict:
    """Replace a reviewed report and retain the exact prior local version."""
    original = inspect_xlsx(source)
    previous = inspect_xlsx(destination)
    if previous['sha256'] == original['sha256']:
        return previous | {'unchanged': True}
    if previous['sha256'] != expected_previous_sha:
        raise ValueError('The previous report changed after review; replacement was not made')
    if source.suffix.lower() != '.xlsx' or destination.suffix.lower() != '.xlsx':
        raise ValueError('Only completed XLSX reports can be replaced')
    copy_checked(destination, archive / previous['sha256'] / destination.name)
    with tempfile.NamedTemporaryFile(dir=destination.parent, suffix='.part', delete=False) as tmp:
        staged = Path(tmp.name)
    try:
        shutil.copyfile(source, staged)
        if inspect_xlsx(staged)['sha256'] != original['sha256']:
            raise ValueError('Replacement source changed during copying')
        if inspect_xlsx(destination)['sha256'] != expected_previous_sha:
            raise ValueError('The previous report changed during replacement preparation')
        os.replace(staged, destination)
    finally:
        staged.unlink(missing_ok=True)
    return inspect_xlsx(destination)


def write_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


@contextmanager
def upload_lock(folder: Path):
    """ОС освобождает блокировку и после аварийного завершения процесса."""
    folder.mkdir(parents=True, exist_ok=True)
    with (folder / "upload.lock").open("a+b") as stream:
        stream.seek(0)
        if not stream.read(1):
            stream.write(b"0")
            stream.flush()
        stream.seek(0)
        if sys.platform == "win32":
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            stream.seek(0)
            if sys.platform == "win32":
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream, fcntl.LOCK_UN)


def publish(cfg: dict, source: Path, day: date, base_id: str, kind: str,
            token_file: Path, state_dir: Path) -> dict:
    """Публикует уже подготовленный отчёт; не извлекает и не исправляет данные 1С.

    OAuth должен быть заранее выдан пользователем. Сам авторизацию не запускает.
    Прямой импорт в Google Sheets здесь намеренно отсутствует: отправляется XLSX.
    """
    if not cfg.get("production_enabled"):
        raise ValueError("Публикация отключена: сначала необходимо сверить состав трёх отчётов")
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from googleapiclient.errors import HttpError
    job = next((j for j in plan(cfg, day) if j["base_id"] == base_id and j["report_type"] == kind), None)
    if job is None:
        raise ValueError("Неизвестная база или тип отчёта")
    if source.name != job["filename"] or source.suffix.lower() != ".xlsx":
        raise ValueError(f"Ожидается файл {job['filename']}")
    checked = inspect_xlsx(source)
    creds = Credentials.from_authorized_user_file(str(token_file))
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    folder_id = cfg["drive_folder_id"]
    folder = service.files().get(fileId=folder_id, fields="id,mimeType,capabilities(canAddChildren)", supportsAllDrives=True).execute(num_retries=3)
    if folder.get("mimeType") != "application/vnd.google-apps.folder" or not folder.get("capabilities", {}).get("canAddChildren"):
        raise PermissionError("Нет права добавлять файлы в указанную папку Google Drive")
    key_hash = hashlib.sha256(job["key"].encode()).hexdigest()
    journal_path = state_dir / (key_hash + ".json")
    with upload_lock(state_dir):
        saved = json.loads(journal_path.read_text(encoding="utf-8")) if journal_path.exists() else {}
        q = f"'{folder_id}' in parents and trashed = false and appProperties has {{ key='report_key' and value='{key_hash}' }}"
        found = service.files().list(q=q, fields="files(id,appProperties)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute(num_retries=3).get("files", [])
        if len(found) > 1:
            raise ValueError("На Диске уже несколько файлов для одной даты, базы и отчёта")
        existing = found[0] if found else None
        if existing and existing.get("appProperties", {}).get("sha256") == checked["sha256"]:
            return {"status": "unchanged", "id": existing["id"], "key": job["key"]}
        file_id = existing["id"] if existing else saved.get("file_id")
        if not file_id:
            file_id = service.files().generateIds(count=1, space="drive", type="files").execute(num_retries=3)["ids"][0]
        # ID фиксируется ДО сетевой записи: потеря ответа не создаст второй файл.
        write_json(journal_path, {"file_id": file_id, "key": job["key"], "status": "pending"})
        metadata = {"name": job["filename"], "appProperties": {"report_key": key_hash, "sha256": checked["sha256"]}}
        media = MediaFileUpload(str(source), mimetype=MIME, resumable=True)
        if existing:
            request = service.files().update(fileId=file_id, body=metadata, media_body=media, fields="id,webViewLink,md5Checksum", supportsAllDrives=True)
        else:
            request = service.files().create(body=metadata | {"id": file_id, "parents": [folder_id]}, media_body=media, fields="id,webViewLink,md5Checksum", supportsAllDrives=True)
        try:
            result = request.execute(num_retries=3)
        except HttpError as exc:
            if exc.resp.status != 409 or existing:
                raise
            remote = service.files().get(fileId=file_id, fields="id,parents,appProperties,webViewLink,md5Checksum", supportsAllDrives=True).execute(num_retries=3)
            if folder_id not in remote.get("parents", []) or remote.get("appProperties", {}).get("report_key") != key_hash:
                raise ValueError("Сохранённый ID не соответствует папке и ключу отчёта") from exc
            result = service.files().update(fileId=file_id, body=metadata, media_body=MediaFileUpload(str(source), mimetype=MIME, resumable=True), fields="id,webViewLink,md5Checksum", supportsAllDrives=True).execute(num_retries=3)
        local_md5 = hashlib.md5(source.read_bytes()).hexdigest()
        if result.get("md5Checksum") != local_md5:
            raise ValueError("Проверка загруженного файла Google Drive не прошла")
        write_json(journal_path, result | {"key": job["key"], "status": "uploaded", "sha256": checked["sha256"]})
        return result | {"status": "uploaded", "key": job["key"]}


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--config", type=Path, default=ROOT / "bases.json")
    sub = p.add_subparsers(dest="command", required=True)
    command = sub.add_parser("plan", help="Создать перечень файлов активных баз")
    command.add_argument("--date", help="YYYY-MM-DD, по умолчанию вчера по Москве")
    command.add_argument("--out", type=Path)
    command = sub.add_parser("verify", help="Проверить файл XLSX, включая временный")
    command.add_argument("file", type=Path)
    command = sub.add_parser("publish", help="Загрузить проверенный готовый отчёт после подключения OAuth")
    command.add_argument("file", type=Path)
    command.add_argument("--date", required=True)
    command.add_argument("--base", required=True)
    command.add_argument("--report", required=True)
    command.add_argument("--token", type=Path, required=True)
    command.add_argument("--state-dir", type=Path, default=ROOT / "state")
    a = p.parse_args()
    try:
        if a.command == "verify":
            result = inspect_xlsx(a.file)
        elif a.command == "plan":
            result = plan(read_config(a.config), report_day(a.date))
            if a.out:
                write_json(a.out, result)
        else:
            result = publish(read_config(a.config), a.file, report_day(a.date), a.base, a.report, a.token, a.state_dir)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        # Не выводим сетевые исключения целиком: в URL/ответах могут быть секреты.
        if isinstance(exc, (ValueError, FileExistsError, PermissionError, zipfile.BadZipFile)):
            print(f"Ошибка: {exc}", file=sys.stderr)
        else:
            print(f"Ошибка {type(exc).__name__}; операция не подтверждена", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
