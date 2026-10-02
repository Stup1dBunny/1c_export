"""Bounds for a single 1C report over a full month from the chosen date."""
from calendar import monthrange
from datetime import date, datetime, timedelta
from report_tools import MSK


def month_end(start: date, today: date | None = None) -> date:
    """Include start through the day before its next-month anniversary.

    If the anniversary day does not exist, include the next month's last day.
    Never silently shorten an unfinished period.
    """
    year, month = (start.year + 1, 1) if start.month == 12 else (start.year, start.month + 1)
    last = monthrange(year, month)[1]
    end = (date(year, month, start.day) - timedelta(days=1)
           if start.day <= last else date(year, month, last))
    today = today or datetime.now(MSK).date()
    if end >= today:
        raise ValueError(f'Период {start:%d.%m.%Y}–{end:%d.%m.%Y} ещё не завершён. '
                         f'Полную выгрузку можно запустить с {end + timedelta(days=1):%d.%m.%Y}.')
    return end

