import zipfile
import shutil
import tempfile
import itertools
import pandas as pd
import collections
import hashlib
import re
from typing import Literal

import mywidgets.Tools.uiStyle.MarkupRe as MarkupRe


# ********** From excel_workbook.py **********
wscell_address = lambda cell: tuple(x for x in wscell_pattern.search(cell).groups()[::-1] if x)
wscell_pattern = re.compile(r"(?:'(?P<sht>.+?)'!)*(?P<col>\$?[A-Z]+)(?P<row>\$?[0-9]+)")
wstbl_address = lambda tbl: wstbl_pattern.match(tbl).groups()
wstbl_pattern = re.compile(r"(?:'(?P<sht>.+?)'!)*(?P<cell>.+)")
excel_col_to_int = lambda col, base=26:sum(((ord(ch) - ord('A') + 1) * base ** k) for k, ch in enumerate(col.upper()[::-1]))

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
        sht, disc_cell = wstbl_address(disc_cell)
        row, col = wscell_address(disc_cell)
        rmin = int(row)
        cmin = ord(col)
        if sht not in disc_sht:
            disc_sht = [sht]
        if col_offset == 0 and row_offset:
            predicate = lambda x: int(wscell_address(x)[0]) >= int(wscell_address(disc_cell)[0])
        if col_offset and row_offset == 0:
            predicate = lambda x: ord(wscell_address(x)[1]) >= ord(wscell_address(disc_cell)[1])
        else:
            predicate = lambda x: '{0: >4s}{1}'.format(*wscell_address(x)) >= '{0: >4s}{1}'.format(
                *wscell_address(disc_cell))
        if col_offset == 0 and row_offset:
            predicate = lambda x: int(wscell_address(x)[0]) >= int(wscell_address(disc_cell)[0])
        if col_offset and row_offset == 0:
            predicate = lambda x: ord(wscell_address(x)[1]) >= ord(wscell_address(disc_cell)[1])
        else:
            predicate = lambda x: '{0: >4s}{1}'.format(*wscell_address(x)) >= '{0: >4s}{1}'.format(
                *wscell_address(disc_cell))

    try:
        filter_rng, filter_sht, filter_cells = zip(
            *[
                (x, *tbl_addr) for x in cells
                if (tbl_addr := wstbl_address(x)) and tbl_addr[0] in disc_sht
            ]
        )
    except ValueError:
        answ = {}
    else:
        ndx = pd.Index(
            [x for x in itertools.chain(*[y.split(':') for y in filter_cells]) if predicate(x)]
        )
        db = ndx.str.extract(wscell_pattern, expand=True).set_index(ndx)

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


def sha1_encode(text, nhexdigit=None):
    sha1 = hashlib.sha1()
    sha1.update(text.encode('utf-8'))
    hexdigest = sha1.hexdigest()
    nhexdigit = nhexdigit or len(hexdigest)
    return hexdigest[-abs(nhexdigit):]

digit_set = ('0123456789', '0')
char_set = ('ABCDEFGHIJKLMOPQRSTUVWXYZ', chr(ord('A') - 1))

def regex_range(pos1: str, pos2: str, ref_set: str = digit_set[0], zero_char: str = digit_set[1]) -> str:
    '''
    cell_range = 'A5:A30'
    rgx_col = regex_range('A', 'A', *char_set)  #  'A'
    rgx_row = regex_range('5', '30')            #  '(?:[5-9]|[1-2][0-9]|30)'
    rgx = A(?:[5-9]|[1-2][0-9]|30)
    (?:<c r="A(?:[5-9]|[1-2][0-9]|30)"=adr v.*=val>)
    '''

    if pos1 == pos2:
        return pos1

    lpos1, lpos2 = map(len, (pos1, pos2))
    # Aseguramos que pos2 > pos1
    n = max(lpos1, lpos2)
    _pos1 = (n * zero_char + pos1)[-n:]
    _pos2 = (n * zero_char + pos2)[-n:]
    if _pos1 > _pos2:
        _pos1, _pos2 = _pos2, _pos1

    if lpos1 == lpos2 == 1:
        return f'[{_pos1}-{_pos2}]'

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
        item = regex_range(_pos1, _pos2)
        return f'{pfx}{item}'

    to_join = []
    fchars = f'[{ref_set[0]}-{ref_set[-1]}]'
    ndx = len(_pos1) - 1
    _pos1 = _pos1[:-1] + chr(ord(_pos1[ndx]) - 1)
    while ndx:
        pfx = _pos1[:ndx]
        sfx = f'{fchars}{{{(len(_pos1) - (ndx + 1))}}}' if len(_pos1) - (ndx + 1) else ''
        item = pfx + f'[{chr(ord(_pos1[ndx]) + 1)}-{ref_set[-1]}]' + sfx
        to_join.append(item.lstrip(zero_char).replace('{1}', ''))
        ndx -= 1
    sfx = f'{fchars}{{{len(_pos1) - 1}}}'
    if chr(ord(_pos1[0]) + 1) != chr(ord(_pos2[0]) - 1):
        item = f'[{chr(ord(_pos1[0]) + 1)}-{chr(ord(_pos2[0]) - 1)}]'
    else:
        item = chr(ord(_pos1[0]) + 1)
    item = item + sfx
    to_join.append(item.lstrip(zero_char).replace('{1}', ''))

    ndx += 1
    while ndx < len(_pos2) - 1:
        pfx = _pos2[:ndx]
        sfx = f'{fchars}{{{len(_pos2) - (ndx + 1)}}}' if len(_pos2) - (ndx + 1) else ''
        item = pfx + f'[{ref_set[0]}-{chr(ord(_pos2[ndx]) - 1)}]' + sfx
        to_join.append(item.lstrip(zero_char).replace('{1}', ''))
        ndx += 1
    pfx = _pos2[:ndx]
    sfx = ''
    item = pfx + ((f'[{ref_set[0]}-{chr(ord(_pos2[ndx]))}]' + sfx)  if (ref_set[0] != _pos2[ndx]) else _pos2[ndx])
    item = item.lstrip(zero_char).replace('{1}', '')
    to_join.append(item.lstrip(zero_char).replace('{1}', ''))
    return f'(?:{"|".join(to_join)})'


class ExcelXml:

    def __init__(self, fname):
        self.tmpdir = tmpdir = tempfile.gettempdir()
        dstfile = shutil.copy(fname, tmpdir)
        shutil.copystat(fname, dstfile)
        self.zf = zf = zipfile.ZipFile(dstfile, 'a')

        self.file_map = {x.filename: x for x in zf.filelist}
        self.ws_fname = {}
        for key in ('xl/workbook.xml', 'xl/sharedStrings.xml'):
            wb_info = self.file_map[key]
            wb_fname = zf.extract(wb_info, path=tmpdir)
            self.ws_fname[key] = wb_fname

        content = self.get_content('xl/workbook.xml')
        wb_pattern = '(?#<sheet name=name r:id="rId(\\d+)"=id>)'
        cpattern = MarkupRe.compile(wb_pattern)
        self._sheet_names = {key: f'xl/worksheets/sheet{id}.xml' for key, id in cpattern.findall(content)}

    @property
    def sheet_names(self):
        return list(self._sheet_names.keys())

    def get_content(self, key):
        if key not in self.ws_fname:
            key_info = self._sheet_names[key]
            item_info = self.file_map[key_info]
            item_fname = self.zf.extract(item_info, path=self.tmpdir)
            self.ws_fname[key] = item_fname
        with open(self.ws_fname[key], 'r', encoding='utf-8') as f:
            content = f.read()
        return content

    def __getitem__(self, key):
        assert key in self.sheet_names, "Not a valid Woksheet name"
        return self.get_content(key)

    def extract_data(wb, content, regex_pattern, seek_pattern=None):
        it = MarkupRe.compile(regex_pattern)
        if seek_pattern:
            it._seeker = re.compile(seek_pattern)
        values = []
        parameters = None
        for grp_d in it.finditer(content):
            parameters = getattr(grp_d, 'parameters', [])
            items = [*parameters, *grp_d.groupdict().values()]
            values.append(items)
        params_fields = parameters._fields if parameters else []
        keys = [*params_fields, *grp_d.groupdict().keys()]
        df = pd.DataFrame(values, columns=keys)
        return df

    def sheet_data(wb, ws_name, tdata):
        regex_pattern, seek_pattern = wb.regex_pattern.get(ws_name)
        if regex_pattern is None:
            return
        try:
            ndx = ['input', 'calculated'].index(tdata) + 1
        except ValueError:
            ndx = 1
        regex_pattern = regex_pattern % ndx
        content = wb[ws_name]
        return wb.extract_data(content, regex_pattern, seek_pattern)

    @property
    def shared_strings(self):
        key = 'xl/sharedStrings.xml'
        content = self.get_content(key)
        regex_str = '(?#<t *=shared_str>)'
        items = MarkupRe.findall(regex_str, content)
        return items
    
    def data_in_range(wb, ws_name, ws_range:str):
        ws = wb[ws_name]
        try:
            (col1, row1), (col2, row2) = cell_pattern.findall(ws_range)
        except ValueError:
            raise ValueError(f'Invalid range: {ws_range}')
        
        col_regex = regex_range(col1, col2, *char_set)
        row_regex = regex_range(row1, row2)
        range_regex = f'(?:{col_regex})(?:{row_regex})'

        seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
        val_regex = f'(?#<c r="{range_regex}"=adr v.*=val>)'
        fml_regex = f'(?#<c r="{range_regex}"=adr f.ref=_adr f.*=".+?"=fml>)'

        df_fml = wb.get_formulas(ws_name, seek_str, fml_regex, allCells=True)

        pass




    def get_formulas(wb, ws_name, seek_str, regex_str, allCells=False):
        content = wb[ws_name]
        cpat = MarkupRe.compile(regex_str)
        cpat._seeker = MarkupRe.re.compile(seek_str)
        fml_raw = cpat.findall(content)
        if allCells:
            pairs = []
            for adr, fml in fml_raw:
                if ':' in adr:
                    cell1, cell2 = adr.split(':')
                    (linfy, linfx), (lsupy, lsupx) = map(
                            lambda x: (int(x[0]), excel_col_to_int(x[1])),
                            map(wscell_address, (cell1, cell2))
                    )
                    offsets = [
                        (col, row)
                        for col in range(0, lsupx - linfx + 1)
                        for row in range(0, lsupy - linfy + 1)
                    ]
                    fmls = [
                        [
                            offset_rng(cell1, col_offset, row_offset),
                            wscell_pattern.sub(lambda m: offset_rng(m.group(), col_offset, row_offset), fml)
                        ]
                        for col_offset, row_offset in offsets
                    ]
                    pairs.extend(fmls)
                else:
                    pairs.append([adr, fml])
        else:
            pairs = fml_raw
        return pd.DataFrame(pairs, columns=['address', 'formula']).set_index('address')


token_specification = [
    ('NUMBER', r'\d+(\.\d*)?'),  # Integer or decimal number
    ('ASSIGN', r'\='),  # Assignment operator
    ('OP', r'[+\-*/]'),  # Arithmetic operators
    ('COMMA', r','),  # Line endings
    ('ANCHOR', r'\:'),  # Line endings
    ('OPENP', r'\('),  # Line endings
    ('CLOSEP', r'\)'),  # Line endings
    ('ID', r'\$?[A-Z]\$?[1-9][0-9]*'),  # Identifiers
    ('FUNCTION', r'[A-Z]+'),  # Skip over spaces and tabs
    ('SKIP', r'[ ]+'),  # Skip over spaces and tabs
    ('MISMATCH', r'.'),  # Any other character
]
tokenizer = re.compile('|'.join('(?P<%s>%s)' % pair for pair in token_specification))

cell_pattern = re.compile(r'(\$?[A-Z]+)(\$?[0-9]+)')
rgn_pattern = re.compile(r'\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?')

def coords_from_range(rng):
    '''
    Encuentra el rango en que se revisaran las fórmulas.
    fml_col: str. Optional, si es None se supone que el rango de fórmulas es el que sigue al rng.
    rng: str. Rango de tabla que incluye concepto (ej: "B010:G227"
    '''
    rng_inf, rng_sup = rng.split(':')
    min_col, min_row = ord(rng_inf[0]) - ord('A') + 1, int(rng_inf[1:])
    max_col, max_row = (ord(rng_sup[0]) - ord('A') + 1) + 1, int(rng_sup[1:]) + 1
    return dict(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)


def excel_to_pandas_slice(excel_slice, axis, table_name, mask=None):
    CellAttrs = collections.namedtuple('CellAttrs', 'is_anchor cell_col cell_row')
    cell_attrs = lambda cell_ref, axis: CellAttrs(
        cell_ref[0] == '$' if axis == 0 else cell_ref[1:].count('$') == 1,
        cell_ref.strip('$0123456789'),
        cell_ref.strip('$')[len(cell_ref.strip('$0123456789')):].strip('$')
    )

    lst_id = excel_slice
    linf, lsup = f'{lst_id}:{lst_id}'.split(':')[:2]
    linf_cell = cell_attrs(linf, axis)
    lsup_cell = cell_attrs(lsup, axis)
    mask = mask if not linf_cell.is_anchor else None
    if mask is None or (linf_cell.cell_row not in mask and linf_cell.cell_col not in mask):
        row_range = ':'.join(f"'{x}'" for x in {linf_cell.cell_row, lsup_cell.cell_row})
        col_range = ':'.join(f"'{x}'" for x in {linf_cell.cell_col, lsup_cell.cell_col})
        prefix, suffix = f"{row_range}", f", {col_range}"
    elif linf_cell.cell_row in mask:
        row_range = ', '.join(f"'{x}'" for x in mask)
        col_range = ':'.join(f"'{x}'" for x in {linf_cell.cell_col, lsup_cell.cell_col})
        prefix, suffix = f"[{row_range}]", f", {col_range}"
    elif linf_cell.cell_col in mask:
        row_range = ':'.join(f"'{x}'" for x in {linf_cell.cell_row, lsup_cell.cell_row})
        col_range = ', '.join(f"'{x}'" for x in mask)
        prefix, suffix = f"{row_range}", f", [{col_range}]"
    else:
        row_range = ', '.join(f"'{x}'" for x in mask)
        col_range = ', '.join(f"'{x}'" for x in mask)
        prefix, suffix = f"[{row_range}]", f", [{col_range}]"
    py_term = f"{table_name}.loc[{prefix}{suffix}]"
    return py_term


def pythonize_fml(fml: str, table_name: str, axis: None|Literal[0,1]=0, mask=None):
    '''
    Convierte una fórmula de Excel a una fórmula de Python
    :param fml: str. Fórmula de Excel
    :param table_name: str. Nombre de la tabla que contiene la fórmula
    :param axis: int. Eje de la tabla que se está procesando
    :return: str. Fórmula de Python
    '''

    pyfml = ''
    lst_id = ''
    fnc_stack = []
    for mo in re.finditer(tokenizer, fml):
        kind = mo.lastgroup
        token_chr = mo.group()
        nxt_char = fml[mo.end(): mo.end() + 1] if mo.end() < len(fml) else ''
        # print(f'{kind=}: {token_chr=}')
        match kind:
            case 'FUNCTION':
                fnc_name = token_chr if token_chr != 'IF' else 'WHERE'
                fnc_stack.append(fnc_name)
                pyfml += f'np.{fnc_name.lower()}'
            case 'ASSIGN':
                if pyfml.count('=') > 0 and pyfml[-1] not in '<>':
                    pyfml += '='
                pyfml += '='
            case 'OPENP':
                pyfml += '('
                fnc_stack.append('(')
            case 'CLOSEP':
                fnc_stack.pop()
                if fnc_stack and fnc_stack[-1] != '(':
                    fnc_name = fnc_stack.pop()
                    pyfml += f', axis={axis})' if fnc_name == 'SUM' else ')'
                else:
                    pyfml += ')'
            case 'ANCHOR':
                lst_id += token_chr
            case 'ID':
                prefix = token_chr.rstrip('0123456789')
                suffix = token_chr[len(prefix):]
                lst_id += f'{prefix}{suffix:0>3s}'
                if ':' in lst_id or nxt_char != ':':
                    py_term = excel_to_pandas_slice(lst_id, axis, table_name, mask=mask)
                    pyfml += py_term
                    lst_id = ''
                    pass
            case _:
                pyfml += token_chr
    return pyfml


def fml_in_range(openpyxl_sheet, ws_range:str):
    ws = openpyxl_sheet
    rgn_coords = coords_from_range(ws_range)
    rgn_fmls = {
        cell.coordinate: cell.value.replace('=+', '=')
        for k, row in enumerate(ws.iter_rows(**rgn_coords), rgn_coords['min_row'])
        for cell in row
        if cell.value and isinstance(cell.value, str) and cell.value.startswith('=')
    }
    return rgn_fmls

def fml_dependents(rgn_fmls: dict[str, str]):
    pairs = set()
    for coord, fml in rgn_fmls.items():
        for term in rgn_pattern.findall(fml):
            term = term.replace('$', '')
            if ':' in term:
                rgn_coords = coords_from_range(term)
                pairs.update([(coord, f'{chr(ord("A") + col - 1)}{row}')
                 for row in range(rgn_coords['min_row'], rgn_coords['max_row'])
                 for col in range(rgn_coords['min_col'], rgn_coords['max_col'])
                              ])
            else:
                pairs.add((coord, term))
    return sorted(pairs)

def resolution_order(pairs:list[tuple[str, str]]):
    to_process = set(pairs)
    term_dep, term_ind = zip(*pairs)
    res_order = [set(term_ind) - set(term_dep)]
    counter = collections.Counter(term_dep)
    while to_process:
        print(f'For process={len(to_process)}')
        to_batch = [
            (term, dep)
            for term, dep in to_process
            if dep in res_order[-1]
        ]
        counter.subtract([
            term
            for term, dep in to_batch
            if not to_process.remove((term, dep))
        ])
        to_batch = {term for term in counter.keys() if counter[term] == 0}
        [counter.pop(term) for term in to_batch]
        res_order.append(to_batch)


    return res_order

def pd_formulas(res_order, rgn_fmls):
    def fml_equiv(m, item, frst_item):
        col, row = m.groups()
        if (frst_item.isnumeric() and row[0] == '$') or (frst_item.isalpha() and col[0] == '$'):
            return m[0]
        return m[0].replace(item, frst_item)

    cell_address = lambda cell: cell_pattern.match(cell).groups()[::-1]

    formulas = []

    for batch in res_order[1:]:
        rows = collections.defaultdict(list)
        cols = collections.defaultdict(list)
        [
            rows[row].append(col)
            for cell_coord in batch
            if (cell_tple := cell_address(cell_coord)) and (row := cell_tple[0]) and (col := cell_tple[1])
        ]
        # Rows with only one associated col are transferred to the cols map.
        keys = list(rows.keys())
        [cols[col].append(row) for row in keys if len(rows[row]) == 1 and (col:=rows.pop(row)[0])]

        for row in sorted(rows.keys(), key=lambda x: len(rows[x])):
            columns = rows.pop(row)
            while columns:
                frst_col, *columns = sorted(columns)
                test_fml = rgn_fmls[f'{frst_col}{row}']
                mask = [
                    col
                    for col in columns
                    if test_fml == cell_pattern.sub(lambda m: fml_equiv(m, item=col, frst_item=frst_col), rgn_fmls[f'{col}{row}'])
                ]
                if not mask:
                    cols[frst_col].append(row)
                else:
                    mask = [frst_col] + mask
                    formulas.append((row, mask, test_fml))
                    columns = [col for col in columns if col not in mask]

        # cols with only one associated row are transferred to the single_list
        keys = list(cols.keys())
        single_cell_fmls = [
            (row, col, rgn_fmls[f'{col}{row}'])
            for col in keys
            if len(cols[col]) == 1 and (row:=cols.pop(col)[0])
        ]

        for col in sorted(cols.keys(), key=lambda x: len(cols[x])):
            rows = sorted(cols.pop(col))
            while rows:
                frst_row, *rows = rows
                test_fml = rgn_fmls[f'{col}{frst_row}']
                mask = [
                    row
                    for row in rows
                    if test_fml == cell_pattern.sub(lambda m: fml_equiv(m, item=row, frst_item=frst_row), rgn_fmls[f'{col}{row}'])
                ]
                if not mask:
                    single_cell_fmls.append((frst_row, col, test_fml))
                else:
                    mask = [frst_row] + mask
                    formulas.append((col, mask, test_fml))
                    rows = [row for row in rows if row not in mask]
        formulas.extend(single_cell_fmls)
    return formulas


def formula_translation(frst_item:str|list[str], scnd_item:str|list[str],fml:str, table_name:str):
    '''
    Traduce una fórmula de Excel a una fórmula de Python
    :param fml: str. Fórmula de Excel
    :param table_name: str. Nombre de la tabla que contiene la fórmula
    :return: str. Fórmula de Python
    '''

    match (frst_item, scnd_item, fml):
        case (str(frst_item), str(scnd_item), fml):                              # cell fml
            to_pythonize = f'{scnd_item}{frst_item}{fml}'
            mask = None
            axis = None
        case (str(frst_item), list(scnd_item), fml) if frst_item.isnumeric():    # row fml
            to_pythonize = f'{scnd_item[0]}{frst_item}{fml}'
            mask = scnd_item
            axis = 0
        case (str(frst_item), list(scnd_item), fml) if frst_item.isalpha():      # column fml
            to_pythonize = f'{frst_item}{scnd_item[0]}{fml}'
            mask = [f'{row:0>3s}' for row in scnd_item]
            axis = 1
        case _:                     # (list(frst_item), list(scnd_item), fml)    # range fml
            to_pythonize = ''
            mask = None
            axis = None
    return pythonize_fml(to_pythonize, table_name='tbl', axis=axis, mask=mask)




def fml_in_range_old(openpyxl_sheet, ws_range:str, lead_col:str|None=None):
    '''
    Obtiene las fórmulas de un rango de celdas de una hoja de cálculo
    :param openpyxl_sheet: openyxl sheet object. Hoja de cálculo
    :param ws_range: str. Rango de celdas a revisar
    :param lead_col: str. Columna id que será utilizada para verificar fila de fórmulas
    :return: dict. Diccionario con las fórmulas de las celdas del rango
    '''

    ws = openpyxl_sheet
    lead_col = lead_col or ws_range[0]
    rgn_coords = coords_from_range(ws_range)
    linf, lsup = rgn_coords['min_col'], rgn_coords['max_col']
    lead_col = ord(lead_col) - ord('A') + 1 - linf
    rgn_fmls = {
        f'{k:0>3d}': [
            (chr(ord('A') + coord -1), fml)
            for coord, cell in zip(range(linf, lsup), row)
            if (fml:=cell.value.replace('=+', '=') if isinstance(cell.value, str) else cell.value)
        ]
            for k, row in enumerate(ws.iter_rows(**rgn_coords), rgn_coords['min_row'])
            if (lead_cell := row[lead_col].value) and isinstance(lead_cell, str) and lead_cell.startswith('=')
    }
    return rgn_fmls

def test1():
    fmls = ['=IF(H372-H373-H374>=0,H372-H373-H374,0)', '=H144+H161+H166+H191-H204', '=SUM($B$9:$B$375)', '=$B$9:$B$375', '=SUM($B$9:$B$375, $C$9:$C$375)', '=SUM($B$9:$B$375, $C$9:$C$375)']
    for fml in fmls[:1]:
        print(f' **** {fml} ****')
        pyfml = pythonize_fml(fml, 'tbl', 0)
        print(f'{pyfml=}')
        print(2*'\n')

def pd_fml(lst_id, token_chr, axis, nxt_char='', table_name='tbl'):
    cell_attrs = lambda cell_ref, axis: (
        cell_ref[0] == '$' if axis == 0 else cell_ref[1:].count('$') == 1,
        cell_ref.strip('$0123456789'),
        cell_ref.strip('$')[len(cell_ref.strip('$0123456789')):].strip('$')
    )

    prefix = token_chr.rstrip('0123456789')
    suffix = token_chr[len(prefix):]
    lst_id += f'{prefix}{suffix:0>3s}'
    if ':' in lst_id or nxt_char != ':':
        is_anchor = 0
        cell_col = 1
        cell_row = 2
        try:
            linf, lsup = lst_id.split(':')
        except:
            id_attrs = cell_attrs(lst_id, axis)
            if axis == 0:
                prefix = f"'{id_attrs[cell_row]}'"
                suffix = f", '{id_attrs[cell_col]}'" if id_attrs[is_anchor] else ''
            else:
                prefix = f"'{id_attrs[cell_row]}'" if id_attrs[is_anchor] else ':'
                suffix = f", '{id_attrs[cell_col]}'"
        else:
            linf_attrs = cell_attrs(linf, axis)
            lsup_attrs = cell_attrs(lsup, axis)
            if axis == 0:
                prefix = f"'{linf_attrs[cell_row]}'" if linf_attrs[cell_row] == lsup_attrs[
                    cell_row] else f"'{linf_attrs[cell_row]}:{lsup_attrs[cell_row]}'"
                if linf_attrs[cell_col] == lsup_attrs[cell_col]:
                    suffix = f", '{linf_attrs[cell_col]}'" if linf_attrs[is_anchor] or lsup_attrs[is_anchor] else ''
                else:
                    suffix = f", '{linf_attrs[cell_col]}:{lsup_attrs[cell_col]}'"
            else:
                suffix = f", '{linf_attrs[cell_col]}'" if linf_attrs[cell_col] == lsup_attrs[
                    cell_col] else f", '{linf_attrs[cell_col]}:{lsup_attrs[cell_col]}'"
                if linf_attrs[cell_row] == lsup_attrs[cell_row]:
                    prefix = f"'{linf_attrs[cell_row]}'" if linf_attrs[is_anchor] or lsup_attrs[is_anchor] else ':'
                else:
                    prefix = f"'{linf_attrs[cell_row]}:{lsup_attrs[cell_row]}'"

        py_term = f"{table_name}.loc[{prefix}{suffix}]"
        # pyfml += py_term
        # lst_id = ''
        return py_term


def test():
    '''
    H378: =L10-L12
    H382: =SUM(H383:H394)
    H395: =H378+H379+H380+H381-H382
    H398: =IF(H395-H396-H397>=0,H395-H396-H397,0)
    H399: =IF(H407-H408-H409>0,H407-H408-H409,0)
    H407: =SUM(H400:H406)
    H410: =SUM(H411:H419)-H412
    H420: =SUM(H421:H440)
    H441: =H422+H424+H425+H426+H427+H428+H429+H431+H436+H437+H440
    H442: =H421+H423+H430+H432+H433+H434+H435+H438+H439
    H443: =H410+H420
    H444: =ROUND(IF(H398+H399<=0,0,IF($L$443=0,0,IF($M$445>=H443,IF(H443>=H398+H399,H398+H399,H443),IF($M$445>=H398+H399,H398+H399,$M$445)))),0)
    H446: =IF(H395-H396-H397+H399-H444>0,H395-H396-H397+H399-H444,0)
    H450: =SUM(H451:H456)
    H457: =IF(H446=0,H450,H446-H448-H449+H450)


    H382: =SUM(H383:H394)
    I382: =SUM(I383:I394)
    J382: =SUM(J383:J394)
    K382: =SUM(K383:K394)
    L382: =H382+I382+J382+K382

    '''

    to_test = -1
    if to_test == -1:
        fname = r'C:\Users\agmontesb\Documents\GitHub\excel\tests\files\excel_module_test.zip'
        wb = ExcelXml(fname)
        df = wb.data_in_range('Parameters and inner links', 'E4:H9')
        pass
    elif to_test == 0:
        fname = 'C:\\Users\\agmontesb\\Documents\\DIAN\\Renta2024\\F2517_AG2024\\F2517_AG2024.zip'
        wb = ExcelXml(fname)
        df = wb.row_description('esf', 2024)
        pass
    elif to_test == 1:
        import openpyxl as px
        import excel_workbook as excel

        # filename = r"C:\Users\agmontesb\Downloads\test_book.xlsx"
        # ws_name = "Sheet1"
        # ws_range = "B3:F9"      # "H10:L10"

        filename = r"C:\Users\agmontesb\Documents\DIAN\Renta2023\Reporte_Conciliación_Fiscal_F2517V6_AG2023_v1.0.1-2024\Reporte_Conciliación_Fiscal_F2517V6_AG2023_v1.0.1-2024.xlsm"
        wb = px.load_workbook(filename)

        excel_wb = excel.ExcelWorkbook('Form2517')

        ws_name = "H2 (ESF - Patrimonio)"
        ws = wb[ws_name]
        wsheet = excel_wb.create_worksheet(ws_name)

        # Estado de Situación Financiera
        ws_range = "G9:K193"
        fmls, values = excel.data_in_range(ws, ws_range)
        esf_tbl = excel.ExcelTable(wsheet, 'esf_tbl', ws_range, fmls, values, recalc=True)
        esf = esf_tbl.minimun_table()

        # Patrimonio
        ws_range = "G196:K220"
        fmls, values = excel.data_in_range(ws, ws_range)
        pat_tbl = excel.ExcelTable(wsheet, 'pat_tbl', ws_range, fmls, values, recalc=True)
        pat = pat_tbl.minimun_table()

        ws_name = "H3 (ERI - Renta Liquida)"
        ws = wb[ws_name]
        wsheet = excel_wb.create_worksheet(ws_name)

        # Renta líquida cedular
        ws_range = "H376:N457"      # "H10:L10"
        fmls, values = excel.data_in_range(ws, ws_range)
        # values['L10'] = 33182000
        # values['L25'] = 13287242
        # values['M6'] = 42412
        rlc_tbl = excel.ExcelTable(wsheet, 'rlc_tbl', ws_range, fmls, values, recalc=True)
        wsheet.parameters(M6 = 42414)
        # dmy = rlc_tbl.get_formula('L412', 'H420')
        rlc = rlc_tbl.minimun_table()

        # Estado de Resultados
        ws_range = "H9:L375"      # "H10:L10"
        fmls, values = excel.data_in_range(ws, ws_range)
        edr_tbl = excel.ExcelTable(wsheet, 'edr_tbl', ws_range, fmls, values, recalc=True)
        edr = edr_tbl.minimun_table()


        ws_name = 'H7 (Resumen ESF-ERI)'
        ws = wb[ws_name]
        wsheet = excel_wb.create_worksheet(ws_name)

        ws_range = 'G11:I93'
        fmls, values = excel.data_in_range(ws, ws_range)
        res_tbl = excel.ExcelTable(wsheet, 'res_tbl', ws_range, fmls, values, recalc=True)
        res = res_tbl.minimun_table()

        wb.close()
        pass
    elif to_test == 2:
        print(f'{"H010:H73"} {pd_fml("H010:", "H73", 0)} axis=0')
        print(f'{"$H010:H73"} {pd_fml("$H010:", "H73", 0)} axis=0')
        print(f'{"B010:H73"} {pd_fml("B010:", "H73", 0)} axis=0')
        print(f'{"$H73"} {pd_fml("", "$H73", 0)} axis=0')
        print(f'{"$H010:$H73"} {pd_fml("$H010:", "$H73", 0)} axis=0')

        print(f'{"H73"} {pd_fml("", "H73", 0)} axis=0')
        print(f'{"H010:H73"} {pd_fml("H010:", "H73", 0)} axis=0')

        print(f'{"H$73"} {pd_fml("", "H$73", 0)} axis=0')
        print(f'{"H010:H$73"} {pd_fml("H010:", "H$73", 0)} axis=0')
    elif to_test == 3:
        import excel_workbook as excel
        import numpy as np
        tbl = pd.DataFrame(0, index=[f'{chr(ord("A") + k)}{j:0>3d}' for k in range(10) for j in range(10)], columns=['H', 'L', 'M'])
        # dmy=np.round(np.where(tbl.loc['398', 'H']+tbl.loc['399', 'H']<=0,0,np.where(tbl.loc['443', 'L']==0,0,np.where(tbl.loc['445', 'M']>=tbl.loc['443', 'H'],np.where(tbl.loc['443', 'H']>=tbl.loc['398', 'H']+tbl.loc['399', 'H'],tbl.loc['398', 'H']+tbl.loc['399', 'H'],tbl.loc['443', 'H']),np.where(tbl.loc['445', 'M']>=tbl.loc['398', 'H']+tbl.loc['399', 'H'],tbl.loc['398', 'H']+tbl.loc['399', 'H'],tbl.loc['445', 'M'])))),0)

        fmls = [
            "A10='Hoja 1 (Valor de las cosas)'!$B$10",
            "A10=SUM('Hoja 1 (Valor de las cosas)'!B10:B12)",
            'B375=SUM(B376:B390)',
            '=B375',
            '=ROUND(IF(H398+H399<=0,0,IF($L$443=0,0,IF($M$445>=H443,IF(H443>=H398+H399,H398+H399,H443),IF($M$445>=H398+H399,H398+H399,$M$445)))),0)',
            '=B9:B375', '=H372-H373', '=IF(H372-H373-H374>=0,H372-H373-H374,0)', '=H144+H161+H166+H191-H204', '=SUM($B$9:$B$375)', '=$B$9:$B$375', '=SUM($B$9:$B$375, $C$9:$C$375)', '=SUM($B$9:$B$375, $C$9:$C$375)']
        import numpy as np
        tbl = pd.DataFrame(0, index=[f'{chr(ord("A") + k)}{j:0>3d}' for k in range(10) for j in range(10)], columns=['H', 'L', 'M'])

        for fml in fmls[:2]:
            print(f' **** {fml} ****')
            pyfml = excel.pythonize_fml(fml, 'tbl', 0, mask=None)
            print(f'{pyfml=}')
            pyfml = excel.pythonize_fml(fml, 'tbl', 0, mask=['375', '380', '400'])
            print(f'{pyfml=}')
            pyfml = excel.pythonize_fml(fml, 'tbl', 0, mask=['A', 'B', 'C'])
            print(f'{pyfml=}')


            print(2*'\n')
    elif to_test == 4:
        # col masks:
        excel_slices = [
            ('H010', ['A', 'B', 'C']),
            ('H010:H73', ['A', 'B', 'C']),
            ('A010:A20', ['A', 'B', 'C']),
            ('A10:H73', ['A', 'B', 'C']),
        ]
        for excel_slice, mask in excel_slices:
            print(f'{excel_slice} {mask=} ==> {excel_to_pandas_slice(excel_slice=excel_slice, axis=0, table_name="tbl", mask=mask)}')

        # col masks:
        excel_slices = [
            ('H010', ['10', '11', '30']),
            ('H10', ['10', '11', '30']),
            ('H10:H73', ['10', '11', '30']),
            ('A10:A20', ['10', '11', '30']),
            ('A10:H73', ['10', '11', '30']),
        ]
        for excel_slice, mask in excel_slices:
            print(f'{excel_slice} {mask=} ==> {excel_to_pandas_slice(excel_slice=excel_slice, axis=0, table_name="tbl", mask=mask)}')

        excel_slices = ['H010', 'H010:H73', 'A010:H010', 'A10:H73']
        for excel_slice in excel_slices:
            print(f'{excel_slice}, mask=None ==> {excel_to_pandas_slice(excel_slice, axis=0, table_name="tbl")}')
if __name__ == '__main__':
    test()


