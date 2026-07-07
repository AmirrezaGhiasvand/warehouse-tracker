# -*- coding: utf-8 -*-
"""
persian_date.py
Small helper around the jdatetime library for converting between
Persian (Jalali/Shamsi) and Gregorian dates, and for populating the
date-picker comboboxes in the UI.
"""

import datetime
import jdatetime

PERSIAN_MONTHS = [
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
]


def today_jalali():
    return jdatetime.date.today()


def days_in_month(year, month):
    """Number of days in a given Jalali month/year."""
    if month <= 6:
        return 31
    if month <= 11:
        return 30
    return 30 if jdatetime.date(year, 1, 1).isleap() else 29


def jalali_to_gregorian_str(year, month, day):
    """Convert a Jalali y/m/d into a Gregorian 'YYYY-MM-DD' string,
    suitable for matching against the 'YYYY-MM-DD HH:MM:SS' timestamps
    stored in the database."""
    jd = jdatetime.date(year, month, day)
    g = jd.togregorian()
    return g.strftime("%Y-%m-%d")


def gregorian_str_to_jalali_display(gregorian_str):
    """Convert a 'YYYY-MM-DD HH:MM:SS' (or 'YYYY-MM-DD') string into a
    human-readable Jalali date string, e.g. '16 تیر 1403'."""
    date_part = gregorian_str.split(" ")[0]
    y, m, d = (int(p) for p in date_part.split("-"))
    jd = jdatetime.date.fromgregorian(date=datetime.date(y, m, d))
    return f"{jd.day} {PERSIAN_MONTHS[jd.month - 1]} {jd.year}"