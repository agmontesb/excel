import itertools
import re
from typing import Callable

import pandas as pd

token_specification = [
    ('NUMBER', r'\d+(\.\d*)?'),  # Integer or decimal number
    ('STRING', r'".*?"'),  # string
    ('ASSIGN', r'\='),  # Assignment operator
    ('SOP', r'^|&|<>'),  # Special operators
    ('OP', r'[+\-*/]'),  # Arithmetic operators
    ('COMMA', r','),  # Line endings
    ('ANCHOR', r'\:'),  # Line endings
    ('OPENP', r'\('),  # Line endings
    ('CLOSEP', r'\)'),  # Line endings
    ('BOOL', r'TRUE|FALSE'),  # Line endings
    ('SHEET', r"'[^']+'!"),  # Sheet names
    ('ERROR', r'#[A-Z/0]+[!?]'), # Error values
    ('CELL', r'\$?[A-Z]\$?[1-9][0-9]*'),  # Identifiers
    ('FUNCTION', r'[A-Z]+'),  # Skip over spaces and tabs
    ('SKIP', r'[ %]+'),  # Skip over spaces and tabs
    ('MISMATCH', r'.'),  # Any other character
]
tokenizer = re.compile('|'.join('(?P<%s>%s)' % pair for pair in token_specification))

link_pattern = re.compile(r"((?:'(?:.+?)'!)+(?:\$?[A-Z])+(?:\$?[0-9])+)")
cell_pattern = re.compile(r"(?:'(?P<sht>.+?)'!)*(?P<col>\$?[A-Z]+)(?P<row>\$?[0-9]+)")
cell_address:Callable[[str], tuple[str, ...]] = lambda cell: tuple(x for x in cell_pattern.search(cell).groups()[::-1] if x)
rgn_pattern_grp = re.compile("(?:'(?P<sht>.+?)'!)*(?P<col>\\$?[A-Z]+)(?P<row>\\$?[0-9]+)(?::\\$?[A-Z]\\$?[0-9]+)?")
rgn_pattern = re.compile(r"(?:'.+?'!)*\$?[A-Z]\$?[0-9]+(?::\$?[A-Z]\$?[0-9]+)?")
tbl_pattern = re.compile(r"(?:'(?P<sht>.+?)'!)*(?P<cell>.+)")
tbl_address = lambda tbl: tbl_pattern.match(tbl).groups()

def alpha_code(id, nbase=26):
    answ = []
    while id:
        res = id % nbase
        id = id // nbase
        if res == 0:
            res = nbase
            id -=1
        answ.append(chr(ord('A') + res - 1))
    return ''.join(answ[::-1])

def code_alpha(code: str, nbase:int = 26) -> int:
    return sum((ord(c) - ord('A') + 1) * (nbase ** i) for i, c in enumerate(reversed(code)))

def offset_rng(cells: str | list[str], col_offset: int = 0, row_offset: int = 0,
               disc_cell: str | None = None, tbl: str | None = None) -> str | dict[str, str]:
    """Transforma cells (que puede ser una celda o una lista de celdas de excel) desplazadas
     por col_offset y row_offset, y devuelve una celda o si cells es una lista  un diccionario con las celdas transformadas.
     offset_rng(['$B$5', 'A$1', '$B3', 'B3'], 1, 1) = {'$B$5': '$B$5', 'A$1': 'B$1', '$B3': '$B4', 'B3': 'C4'}
    Esta función viene del proyecto excel módulo excel_workbook y clase ExcelTable con pequeñas modificaciones.
     """
    if bflag := isinstance(cells, str):
        cells = [cells]

    # disc_sht = [None, tbl.parent.title] if tbl else [None, ]
    disc_sht = [None, tbl] if tbl else [None, ]
    predicate = lambda x: True
    # Cuando se eliminen celdas, se debe asegurar que el offset no sobrepase los límites de la tabla
    rmin, cmin = 1, ord('A')
    if disc_cell:
        sht, disc_cell = tbl_address(disc_cell)
        row, col = cell_address(disc_cell)
        rmin = int(row)
        cmin = ord(col)
        if sht not in disc_sht:
            disc_sht = [sht]
        if col_offset == 0 and row_offset:
            predicate = lambda x: int(cell_address(x)[0]) >= int(cell_address(disc_cell)[0])
        if col_offset and row_offset == 0:
            predicate = lambda x: ord(cell_address(x)[1]) >= ord(cell_address(disc_cell)[1])
        else:
            predicate = lambda x: '{0: >4s}{1}'.format(*cell_address(x)) >= '{0: >4s}{1}'.format(
                *cell_address(disc_cell))
        if col_offset == 0 and row_offset:
            predicate = lambda x: int(cell_address(x)[0]) >= int(cell_address(disc_cell)[0])
        if col_offset and row_offset == 0:
            predicate = lambda x: ord(cell_address(x)[1]) >= ord(cell_address(disc_cell)[1])
        else:
            predicate = lambda x: '{0: >4s}{1}'.format(*cell_address(x)) >= '{0: >4s}{1}'.format(
                *cell_address(disc_cell))

    try:
        filter_rng, filter_sht, filter_cells = zip(
            *[
                (x, *tbl_addr) for x in cells
                if (tbl_addr := tbl_address(x)) and tbl_addr[0] in disc_sht
            ]
        )
    except ValueError:
        answ = {}
    else:
        ndx = pd.Index(
            [x for x in itertools.chain(*[y.split(':') for y in filter_cells]) if predicate(x)]
        )
        db = ndx.str.extract(cell_pattern, expand=True).set_index(ndx)

        mask = ~db.row.str.contains('$', regex=False)
        fnc = lambda x: str(max(rmin, int(x.strip('$')) + row_offset))
        db.loc[mask, 'row'] = db.loc[mask, 'row'].apply(fnc)

        mask = ~db.col.str.contains('$', regex=False)
        fnc = lambda x: chr(max(cmin, ord(x) + col_offset))
        db.loc[mask, 'col'] = db.loc[mask, 'col'].apply(fnc)

        db['cell'] = db.col + db.row
        cells_map = db.cell.to_dict()
        values = [':'.join(map(lambda x: cells_map.get(x, x), key.split(':'))) for key in filter_cells]
        values = [(f"'{sht}'!" if sht else '') + value for sht, value in zip(filter_sht, values)]
        answ = dict(zip(filter_rng, values))
    return answ.get(cells[0], cells[0]) if bflag else answ

digit_set = ('0123456789', '0')
char_set = ('ABCDEFGHIJKLMOPQRSTUVWXYZ', chr(ord('A') - 1))

def interval_regex(pos1: str, pos2: str, ref_set: str = digit_set[0], zero_char: str = digit_set[1], has_pfx: bool = False) -> str:
    '''
    cell_range = 'A5:A30'
    rgx_col = regex_range('A', 'A', *char_set)  #  'A'
    rgx_row = regex_range('5', '30')            #  '(?:[5-9]|[1-2][0-9]|30)'
    rgx = A(?:[5-9]|[1-2][0-9]|30)
    (?:<c r="A(?:[5-9]|[1-2][0-9]|30)"=adr v.*=val>)
    '''

    def clean_item(item:str, has_pfx: bool) -> str:
        if not has_pfx:
            item = item.lstrip(zero_char)
        return item.replace('{1}', '').replace('@', 'A')

    if pos1 == pos2:
        item = pos1
        return item if not has_pfx else [item]

    lpos1, lpos2 = map(len, (pos1, pos2))
    # Aseguramos que pos2 > pos1
    n = max(lpos1, lpos2)
    _pos1 = (n * zero_char + pos1)[-n:]
    _pos2 = (n * zero_char + pos2)[-n:]
    if _pos1 > _pos2:
        _pos1, _pos2 = _pos2, _pos1

    if lpos1 == lpos2 == 1:
        item = f'[{_pos1}-{_pos2}]'.replace('@', 'A')
        return item if not has_pfx else [item]

    # Sabemos que _pos1 y _pos2 son difrentes en al menos un número:
    all_match = {(k, ch) for k, ch in enumerate(_pos1[::-1])}
    all_match = sorted(all_match.intersection((k, ch) for k, ch in enumerate(_pos2[::-1])))

    pfx = ''
    if all_match and all_match[-1][0] == n - 1:
        lpos = all_match[-1][0] + 1
        while all_match and lpos - all_match[-1][0] == 1:
            *all_match, tpl = all_match
            lpos, ch = tpl
            pfx = pfx + ch

    if pfx:
        lpfx = len(pfx)
        _pos1, _pos2 = _pos1[lpfx:n], _pos2[lpfx:n]
        to_join = [f'{pfx}{item}' for item in interval_regex(_pos1, _pos2, ref_set, zero_char, True)]
        return to_join if has_pfx else f'(?:{"|".join(to_join)})'

    to_join = []
    fchars = f'[{ref_set[0]}-{ref_set[-1]}]'
    ndx = len(_pos1) - 1
    _pos1 = _pos1[:-1] + chr(max(ord(_pos1[ndx]), ord(zero_char)) - 1)
    while ndx:
        lsup = ord(ref_set[-1])
        linf = max(min(ord(_pos1[ndx]) + 1, lsup), ord(zero_char))
        if linf < lsup:
            pfx = _pos1[:ndx]
            sfx = f'{fchars}{{{(len(_pos1) - (ndx + 1))}}}' if len(_pos1) - (ndx + 1) else ''
            item = pfx + f'[{chr(ord(_pos1[ndx]) + 1)}-{ref_set[-1]}]' + sfx
            to_join.append(clean_item(item, has_pfx))
        ndx -= 1
    sfx = f'{fchars}{{{len(_pos1) - 1}}}'
    if chr(ord(_pos1[0]) + 1) <= chr(ord(_pos2[0]) - 1):
        if chr(ord(_pos1[0]) + 1) == chr(ord(_pos2[0]) - 1):
            item = chr(ord(_pos1[0]) + 1)
        else:
            item = f'[{chr(ord(_pos1[0]) + 1)}-{chr(ord(_pos2[0]) - 1)}]'
        item = item + sfx
        to_join.append(clean_item(item, has_pfx))

    posx = clean_item(_pos2[0] + (len(_pos2) - 1) * zero_char, False)
    suffix = interval_regex(posx, _pos2, ref_set, zero_char, True)
    to_join.extend(suffix)
    if not has_pfx:
        to_join = to_join[::-1]
    return to_join if has_pfx else f'(?:{"|".join(to_join)})'

def regex_range(range_str:str) -> str:
    '''
    Crea una expresión regular que cubre el rango entre range_str1 y range_str2
    :param range_str: str. Rango en formato Excel (ej: "A5:A30" o "B:C" o "10:50")
    :return: str. Expresión regular que cubre el rango
    '''
    try:
        (col1, row1), *tail = [tpl[1:] for tpl in cell_pattern.findall(range_str)]
    except ValueError:
        try:
            pos1, pos2 = range_str.split(':')
            if pos1[0].isalpha() and pos2[0].isalpha():
                # Rango de columnas
                return interval_regex(pos1, pos2, *char_set) + r'\d+'
            elif pos1[0].isnumeric() and pos2[0].isnumeric():
                # Rango de filas
                return '[A-Z]+' + interval_regex(pos1, pos2, *digit_set)
        except ValueError:
            raise ValueError(f'Invalid range string: {range_str}')
    if not tail:
        return range_str
    col2, row2 = tail[0]
    col_regex = interval_regex(col1, col2, *char_set)
    row_regex = interval_regex(row1, row2, *digit_set)
    return f'{col_regex}{row_regex}'

def from_a1_tuple(cell: str) -> tuple[int, int]:
    """Convierte una referencia de celda en formato A1 a una tupla (fila, columna) basada en 1."""
    sht_str, cell_str = tbl_address(cell)
    match = cell_pattern.fullmatch(cell_str.replace('$', ''))
    if not match:
        raise ValueError(f'Invalid cell reference: {cell}')
    _, col_str, row_str  = match.groups()
    col_num = code_alpha(col_str)
    row_num = int(row_str)
    return sht_str, col_num, row_num

def from_r1c1_tuple(cell: str) -> tuple[int, int]:
    """Convierte una referencia de celda en formato R1C1 a una tupla (fila, columna) basada en 1."""
    sht_str, cell_str = tbl_address(cell)
    match = re.fullmatch(r'R(\$?)(\d+)C(\$?)(\d+)', cell_str)
    if not match:
        raise ValueError(f'Invalid cell reference: {cell}')
    _, row_str, _, col_str = match.groups()
    col_num = int(col_str)
    row_num = int(row_str)
    return sht_str, col_num, row_num

def from_a1_r1c1(cell: str) -> str:
    """Convierte una referencia de celda en formato A1 a formato R1C1."""
    sht_str, col_num, row_num = from_a1_tuple(cell)
    prefix = f"'{sht_str}'!" if sht_str else ''
    return f'{prefix}R{row_num}C{col_num}'

def from_r1c1_a1(cell: str) -> str:
    """Convierte una referencia de celda en formato R1C1 a formato A1."""
    sht_str, col_num, row_num = from_r1c1_tuple(cell)
    prefix = f"'{sht_str}'!" if sht_str else ''
    col_str = alpha_code(col_num)
    return f'{prefix}{col_str}{row_num}'


def formulaR1C1(formulaA1: str, refA1: str) -> list[str]:
    _, ref_col, ref_row = from_a1_tuple(refA1)
    
    def fn(m):
        frng = m.group(0)
        sht, rng = tbl_address(frng)
        cells = [cell[1:] for cell in cell_pattern.findall(rng)]
        cells_r1c1 = []
        for col, row in cells:
            if col[0] == '$':
                col = f'C{code_alpha(col[1:])}'
            else:
                ncol = code_alpha(col)
                col = f'C[{ncol - ref_col}]'

            if row[0] == '$':
                row = f'R{int(row[1:])}'
            else:
                nrow = int(row)
                row = f'R[{nrow - ref_row}]'
            cells_r1c1.append(f'{row}{col}')
        cells_r1c1 = ':'.join(cells_r1c1).replace('[0]', '')
        if sht:
            cells_r1c1 = f"'{sht}'!" + cells_r1c1
        return cells_r1c1
    
    answ = rgn_pattern.sub(fn, formulaA1)
    return answ



def main():
    pass

if __name__ == '__main__':
    main()