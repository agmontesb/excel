import locale
import pandas as pd
import numpy as np
import re
from functools import wraps
from datetime import datetime
from datetime import date as pydate

func_dict = {}

is_leap = lambda year: year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)    
month_days = lambda month, year: (30 + (month + month // 8) % 2) if month != 2 else [28, 29][is_leap(year)]

month_map = {}
for k in range(1, 13):
    nmonth, month, fmonth = map(str.lower, (f'{k:02d}', *pydate(2024, k, 1).strftime('%b %B').split(' ')))
    month_map[month.replace('.', '')] = nmonth
    month_map[fmonth] = nmonth

def serial_to_date(ndate):
    nyear = min(ndate//365 + 1900, 9999)
    while True:
        res = ndate - datevalue(f'{nyear-1}/12/31')
        if res > 0:
            break
        nyear -= 1
    nmonth = 1
    while True:
        ndays = month_days(nmonth, nyear)
        if res > ndays:
            nmonth += 1
            res -= ndays
        else:
            nday =  res
            break  
    return pydate(nyear, nmonth, nday)

def register_function(func_cat):
    def decorator(func):
        if func_cat not in func_dict:
            func_dict[func_cat] = []
        func_dict[func_cat].append(func)
        @wraps(func)
        def fargs(*args, **kwargs):
            data = []
            for x in args:
                try:
                    data.extend(x.flatten().tolist())
                except:
                    data.append(x)
            answer = func(*data, **kwargs)
            try:
                return answer.iloc[0]
            except:
                return answer
        return fargs
    return decorator


@register_function('math')
def sum(*data):
    s = pd.Series(data)
    return s.sum()

@register_function('math')
def if_(condition, true_value, false_value):
    args = []
    for x in (condition, true_value, false_value):
        try:
            fx = x.flatten().tolist()
        except:
            fx = [x]
        args.append(pd.Series(x))
    return np.where(*args)[0]

@register_function('text')
def char(*data):
    s = pd.Series(data)
    return s.map(lambda x: chr(x))

@register_function('text')
def clean(*data):
    class TrTable:
        def __getitem__(self, key):
            if chr(key).isprintable():
                raise LookupError()
    s = pd.Series(data)
    return s.str.translate(TrTable())

@register_function('text')
def code(*data):
    s = pd.Series(data)
    return s.str[0].map(lambda x: ord(x))

@register_function('text')
def concat(*data):
    s = pd.Series(data)
    answ = s.str.cat()
    if len(answ) > 32767:
        from excel_workbook import XlErrors
        answ = XlErrors.VALUE_ERROR
    return pd.Series([answ])

@register_function('text')
def dollar(*data):
    # locale.setlocale(locale.LC_ALL, '')

    number, decimals = (data + (None,))[:2]
    decimals = decimals or 2
    bflag = number < 0
    number =  abs(number)
    number = round(number * 10.0**decimals, 0) * 10**(-decimals)
    n = max(0, decimals)
    answ = locale.format_string('$%.{}f'.format(n), number, grouping=True)
    if bflag:
        answ =f'({answ})'
    return pd.Series([answ])


@register_function('text')
def find(find_text, within_text, start_num=1):
    if find_text == '':
        return start_num
    if start_num < 1 or start_num > len(within_text) +1 or find_text not in within_text:
        from excel_workbook import XlErrors
        return XlErrors.VALUE_ERROR
    answ = within_text.find(find_text, start_num - 1) + 1
    return answ

@register_function('text')
def fixed(number, decimals=2, no_commas=False):
    # locale.setlocale(locale.LC_ALL, '')

    bflag = number < 0
    number =  abs(number)
    number = round(number * 10.0**decimals, 0) * 10**(-decimals)
    n = max(0, decimals)
    answ = locale.format_string('%.{}f'.format(n), number, grouping= not no_commas)
    if bflag:
        answ =f'-{answ}'
    return pd.Series([answ])
    

@register_function('text')
def left(text, num_chars=1):
    num_chars = max(1, num_chars)
    return text[:num_chars]

@register_function('text')
def len(*data):
    s = pd.Series(data)
    return s.str.len()

@register_function('text')
def lower(*data):
    s = pd.Series(data)
    return s.str.lower()

@register_function('text')
def mid(text, start_num, num_chars):
    if start_num < 1:
        from excel_workbook import XlErrors
        return XlErrors.VALUE_ERROR
    return text[start_num - 1: start_num - 1 + num_chars]

@register_function('text')
def numbervalue(text, decimal_separator=None, group_separator=None):
    if not text:
        return 0
    # locale.setlocale(locale.LC_ALL, '')
    decimal_separator = decimal_separator or locale.localeconv()['decimal_point']
    group_separator = group_separator or locale.localeconv()['thousands_sep']
    decimal_separator, group_separator =  decimal_separator[0], group_separator[0]
    text = text.replace(' ', '')
    if text.count(decimal_separator) > 1:
        from excel_workbook import XlErrors
        return XlErrors.VALUE_ERROR
    if text.split(decimal_separator)[1].count(group_separator):
        from excel_workbook import XlErrors
        return XlErrors.VALUE_ERROR
    text = text.replace(group_separator, '')
    text = text.replace(decimal_separator, '.')
    npercent = len(text) - len(text.rstrip('%'))
    text = text.rstrip('%')
    if not all(map(str.isnumeric, text.split('.'))):
        from excel_workbook import XlErrors
        return XlErrors.VALUE_ERROR
    answ = float(text) / 100.0**npercent
    return answ

@register_function('text')
def proper(text):
    return text.title()

@register_function('text')
def replace(old_text, start_num, num_chars, new_text):
    return old_text[:start_num - 1] + new_text + old_text[start_num - 1 + num_chars:]

@register_function('text')
def right(text, num_chars=1):
    num_chars = max(1, num_chars)
    return text[-num_chars:]

@register_function('text')
def trim(text):
    text = text.strip(' ')
    text = ' '.join(text.split())
    return text

@register_function('text')
def upper(*data):
    s = pd.Series(data)
    return s.str.upper()

@register_function('date')
def date(year, month, day):
    if 1900 <= year <= 9999:
        pass
    elif 0 <= year <= 1899:
        year += 1900
    else:
        from excel_workbook import XlErrors
        return XlErrors.NUM_ERROR
    if 1 <= month <= 12:
        pass
    elif month > 12:
        month, dyear = month % 12, month // 12
        year += dyear
    elif month < 1:
        month, year = 12 + month, year - 1

    ndays = month_days(month, year)
    if 1 <= day <= ndays:
        pass
    elif day > ndays:
        while day > ndays:
            day -= ndays
            month += 1
            dyear, month = month // 13, month % 13 + month // 13
            year += dyear
            ndays = month_days(month, year)
    elif day < 1:
        while day < 1:
            day += ndays
            month -= 1
            if month == 0:
                month, year = 12, year - 1
            ndays = month_days(month, year)
    return datevalue(f'{day}/{month}/{year}')

@register_function('date')
def datevalue(date_text):
    try:
        prefix, month, *suffix = re.split(r'[/-]', date_text.lower())
        month = int(month_map.get(month, month))
        if not suffix:
            year = datetime.now().year
            day = int(prefix)
        else:
            suffix, = suffix    # Si len(suffix) > 1 eleva ValueError 
            year, day = (int(prefix), int(suffix)) if len(prefix) == 4 else (int(suffix), int(prefix))
        year = (year + 1900) if year < 100 else year
        ndays = month_days(month, year)
        bflags = [1900 <= year <= 9999, 1 <= month <= 12, 1 <= day <= ndays]
        year, month, day = (x for bflag, x in zip(bflags, (year, month, day)) if bflag)
        reference_date = pydate(1900 - 1, 12, 31)
        delta = pydate(year, month, day) - reference_date
        excel_bug = int(year > 1900 or month > 2) # Excel considera 1900 como año bisiesto lo cual es un error
        answ = delta.days + excel_bug
    except:
        from excel_workbook import XlErrors
        answ = XlErrors.VALUE_ERROR
    return answ

@register_function('date')
def day(serial_date):
    return serial_to_date(serial_date).day

@register_function('date')
def days360(start_date, end_date, method=False):
    pstart = serial_to_date(start_date)
    pend = serial_to_date(end_date)
    if method:
        tpl_start = (pstart.year, pstart.month, min(30, pstart.day))
        tpl_end = (pend.year, pend.month, min(30, pend.day))
    else:
        tpl_start = (
            pstart.year, 
            pstart.month, 
            30 if pstart.day == month_days(pstart.month, pstart.year) else pstart.day
        )
        pend_is_lastday = pend.day == 31 and pend.day == month_days(pend.month, pend.year)
        tpl_end = (
            pend.year, 
            pend.month + (1 if pend_is_lastday and pstart.day < 30 else 0),
            (1 if pstart.day < 30 else 30) if pend_is_lastday else pend.day
        )
    weights = [360, 30, 1]
    answer = 0
    for weight, y, x in zip(weights, tpl_start, tpl_end):
        answer += weight * (x - y)
    return answer

@register_function('date')
def month(serial_date):
    return serial_to_date(serial_date).month

@register_function('date')
def year(serial_date):
    return serial_to_date(serial_date).year

