import locale
import pandas as pd
import numpy as np
from functools import wraps


func_dict = {}

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
    locale.setlocale(locale.LC_ALL, '')

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
    locale.setlocale(locale.LC_ALL, '')

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
    locale.setlocale(locale.LC_ALL, '')
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
