import inspect
import locale
import pandas as pd
import numpy as np
import re
from functools import wraps
from datetime import datetime
from datetime import date as pydate


__all__ = ['sum', 'if_', 'sumif', 'sumifs']

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

@register_function('math')
def sumif(range_, criteria, sum_range=None):
    sum_range = sum_range or range_
    answer = sumifs(sum_range, range_, criteria)
    return answer

@register_function('math')
def sumifs(sum_range, criteria_range1, criteria1, *criteria_pairs):
    from excel_workbook import  XlErrors
    criteria_pairs = [criteria_range1, criteria1, *criteria_pairs]
    if len(criteria_pairs) % 2 != 0:
        return XlErrors.GETTING_DATA_ERROR
    if not all(len(sum_range) == len(x) for x in criteria_pairs[::2]):
        return XlErrors.VALUE_ERROR
    tr_table = str.maketrans({'.': '', '-': '', '+':''})
    while criteria_pairs:
        criteria_range, criteria, *criteria_pairs = criteria_pairs
        if isinstance(criteria, (str, int, float)):
            criteria = str(criteria)
            if criteria.count('*') > criteria.count('~*') or criteria.count('?') > criteria.count('~?'):
                criteria = criteria.lstrip('=').replace('*', '.*').replace('?', '.?')
                fnc = lambda x: bool(re.search(criteria, x))
            else:
                criteria = criteria.replace('~*', '*').replace('~?', '?').replace('<>', '!=')
                if criteria.startswith(('!=', '>=', '<=')):
                    prefix, criteria = criteria[:2], criteria[2:]
                elif criteria.startswith(('>', '<', '=')):
                    prefix, criteria = criteria[:1], criteria[1:]
                    prefix = prefix.replace('=', '==')
                else:
                    prefix = '=='
                if not criteria.translate(tr_table).isdigit():
                    criteria = criteria.strip('"')
                    criteria = f'"{criteria}"'
                criteria = f'{prefix}{criteria}'
                fnc = lambda x: bool(eval(f'{x}{criteria}' if not isinstance(x, str) else f'"{x}"{criteria}'))
        else:
            fnc = criteria
        mask = pd.Series(criteria_range).apply(fnc)
        sum_range = pd.Series(sum_range).where(mask)
    answer = sum_range.sum()
    return answer