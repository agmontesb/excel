import re
from datetime import datetime, date as pydate

from core import register_function

__all__ = ['date', 'datevalue', 'day', 'days', 'days360', 'month', 'today', 'year']

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
def days(start_date, end_date):
    from excel_workbook import XlErrors
    args = [start_date, end_date]
    for k, x in enumerate(args):
        if isinstance(x, str):
            if isinstance(dvalue := datevalue(x), XlErrors):
                return dvalue
        elif not (1 <= x <= datevalue('31/12/9999')):
            return XlErrors.VALUE_ERROR
        else:
            dvalue = x
        args[k] = dvalue
    return args[0] - args[1]


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
def today():
    sdate = pydate.today().strftime('%d/%m/%Y')
    return datevalue(sdate)

@register_function('date')
def year(serial_date):
    return serial_to_date(serial_date).year

