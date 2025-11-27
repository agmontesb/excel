from html import unescape
import os
import zipfile
import shutil
import tempfile
import itertools
import pandas as pd
import collections
import logging
import re
import numpy as np
from typing import Literal

import mywidgets.Tools.uiStyle.MarkupRe as MarkupRe
from mywidgets.Widgets.Custom import navigationbar

from xlobjects import XlErrors


logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# ********** From excel_workbook.py **********
wscell_address = lambda cell: tuple(x for x in wscell_pattern.search(cell).groups()[::-1] if x)
wscell_pattern = re.compile(r"(?:'(?P<sht>.+?)'!)*(?P<col>\$?[A-Z]+)(?P<row>\$?[0-9]+)")
wstbl_address = lambda tbl: wstbl_pattern.match(tbl).groups()
wstbl_pattern = re.compile(r"(?:'(?P<sht>.+?)'!)*(?P<cell>.+)")
excel_col_to_int = lambda col, base=26:sum(((ord(ch) - ord('A') + 1) * base ** k) for k, ch in enumerate(col.upper()[::-1]))

cell_pattern = re.compile(r'(\$?[A-Z]+)(\$?[0-9]+)')

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
        (col1, row1), *tail = cell_pattern.findall(range_str)
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

DEFAULT_RANGE = 'A1:CV1000'   # equivalent to 'R1C1:R100C1000

class WorkSheetXml:
    Cell = collections.namedtuple('Cell', ['address', 'formula', 'value'])

    def __init__(self, ws_name: str, fml_map: dict, val_map: dict):
        self.title = ws_name
        self.cells = {
            adr: WorkSheetXml.Cell(adr, fml_map.get(adr, value), value)
            for adr in set(fml_map.keys()).union(val_map.keys())
            if (value :=val_map.get(adr)) is not None
        }
        pass

    def __getitem__(self, range_str: str): # -> dict[str, str] | 'WorkSheetXml.Cell' | None:
        range_regex = regex_range(range_str)
        if range_regex == range_str:
            return self.cells.get(range_str, None)
        pattern = re.compile(f'^{range_regex}$')
        cells = {
            adr: cell
            for adr, cell in self.cells.items()
            if pattern.match(adr)
        }
        return cells
    
    def cell(self, row:int, column:int) -> 'WorkSheetXml.Cell':
        col_str = ''
        n = column
        while n:
            n, r = divmod(n - 1, 26)
            col_str = chr(ord('A') + r) + col_str
        adr = f'{col_str}{row}'
        return self.cells.get(adr, WorkSheetXml.Cell(adr, '', ''))


class WorkBookXml:

    def __init__(self, fname, default_range=DEFAULT_RANGE):
        loader = OXLLoader(fname)
        self.zf = loader
        zf = loader.zf
        self.file_map = {x.filename: x for x in zf.filelist}
        self.ws_fname = {}
        # tmpdir = os.path.dirname(loader.filename)
        # for key in ('xl/workbook.xml', 'xl/sharedStrings.xml'):
        #     wb_info = self.file_map[key]
        #     wb_fname = zf.extract(wb_info, path=tmpdir)
        #     self.ws_fname[key] = wb_fname

        self.sheet_obs = {}
        self.default_range = default_range
        sheet_id = self.parse_workbook_xlm()
        active_sheet = self.sheetnames[sheet_id]
        self.active = self.__getitem__(active_sheet)
        pass

    def parse_workbook_xlm(self):
        loader = self.zf
        content = loader.load_xmlfile('xl/workbook.xml', isAbs=False)

        # Active sheet
        wb_pattern = '(?#<workbookView activeTab=_tabId>)'
        cpattern = MarkupRe.compile(wb_pattern)
        active_sheet = cpattern.findall(content)[0] or '0'

        # Sheet names
        wb_pattern = '(?#<sheet name=name r:id=id>)'
        cpattern = MarkupRe.compile(wb_pattern)
        self._sheet_names = {key: f'xl/worksheets/sheet{id[3:]}.xml' for key, id in cpattern.findall(content)}

        # Named ranges
        wb_pattern = '(?#<definedName name=name *=rng>)'
        cpattern = MarkupRe.compile(wb_pattern)
        try:
            defined_names = {name: rng for name, rng in cpattern.findall(content)}
            self._defined_names = defined_names
        except Exception as e:
            logger.debug(f'Error loading named ranges: {str(e)}')
            self._defined_names = {}
        return int(active_sheet)

    def parse_worksheet_xlm(self, ws_name):
        if ws_name not in self.sheet_obs:
            try:
                fml_map, val_map = self.data_in_range(ws_name, self.default_range)
                self.sheet_obs[ws_name] = WorkSheetXml(ws_name, fml_map, val_map)
            except Exception as e:
                msg = f'Error loading worksheet "{ws_name}": {str(e)}'
                logger.debug(msg)
                raise Exception(msg)
        return self.sheet_obs[ws_name]

    @property
    def sheetnames(self):
        return list(self._sheet_names.keys())

    def __getitem__(self, key):
        assert key in self.sheetnames, "Not a valid Woksheet name"
        return self.parse_worksheet_xlm(key)

    def extract_data(wb, content, regex_pattern, seek_pattern=None):
        it = MarkupRe.compile(regex_pattern)
        if seek_pattern:
            it._seeker = re.compile(seek_pattern)
        values = []
        for grp_d in it.finditer(content):
            parameters = getattr(grp_d, 'parameters', [])
            items = [*parameters, *grp_d.groupdict().values()]
            values.append(items)
        keys = [*it.get_seeker().groupindex.keys(), *it.groupindex.keys()]
        df = pd.DataFrame(values, columns=keys)
        return df

    @property
    def shared_strings(self):
        loader = self.zf
        key = 'xl/sharedStrings.xml'
        content = loader.load_xmlfile(key, isAbs=False)
        regex_str = '(?#<t *=shared_str>)'
        items = MarkupRe.findall(regex_str, content)
        return items
    
    def data_in_range(wb, ws_name, ws_range:str=None, allCells=True) -> tuple[dict, dict]:
        loader = wb.zf
        rpath = wb._sheet_names[ws_name]
        content = loader.load_xmlfile(rpath, isAbs=False)
        # dimension element: 18.3.1.35 dimension (Worksheet Dimensions)
        rgx_str = '(?#<dimension ref=dim>)'
        dim = MarkupRe.findall(rgx_str, content)[0]

        # sheetView element: 18.3.1.86 sheetView (Chart Sheet View)
        rgx_str = '(?#<sheetView (topLeftCell) showGridLines=_showGridLines showRowColHeaders=_showRowColHeaders>)'
        topLeftCell, showGridLines, showRowColHeaders = MarkupRe.findall(rgx_str, content)[0]
        showGridLines, showRowColHeaders = (showGridLines != '0'), (showRowColHeaders != '0')

        # 18.3.1.66 pane (View Pane)
        # This element only exists if the worksheet contains panes (split or frozen).
        rgx_str = '(?#<pane xSplit=_xSplit ySplit=_ysplit (topLeftCell) (activePane) (state)>)'
        panes = MarkupRe.findall(rgx_str, content)
        if panes:
            # Only consider when state is "frozen" or "frozenSplit", because in 
            # this version the split panes is not implemented
            pass

        # 18.3.1.78 selection (Selection)
        # Exist one (no freeze panes), two (topRow or leftColumn freeze) or three (freeze panes) 
        # of this element. the pane attib is optional.
        rgx_str = '(?#<selection (activeCell) (sqref) pane=_pane>)'

        # 18.3.1.13 col (Column Width & Formatting)
        




        ws_range = ws_range or wb.default_range
        df_val = wb.get_values(content, ws_range, allCells=True)
        df_fml = wb.get_formulas(content, ws_range, allCells=True)
        if not allCells and (to_pop := df_fml.keys() & df_val.keys()):
            for key in to_pop:
                df_val.pop(key)
        return df_fml, df_val

    def get_values(wb, content, ws_range: str, allCells=True):
        range_regex = regex_range(ws_range)
        seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
        val_regex = f'(?#<c r="{range_regex}"=adr t=_t s=_s v.*=val>)'
        if not allCells:
            val_regex = f'(?#<c r="{range_regex}"=adr __NCHILDREN__="2" t=_t s=_s v.*=val>)'
        shared = wb.shared_strings
        try:
            df = (
                wb.extract_data(content, val_regex, seek_str)
                .set_index('adr')
            )
            
            # values = df.set_index('adr').val.to_dict()
            # cell_type = df.set_index('adr').t.to_dict()
            # cell_style = df.set_index('adr').s.to_dict()


            # ECMA-376 18.18.11
            # t = s (Shared String): Cell containing a shared string.
            # [values.__setitem__(key, shared[int(value)]) for key, value in cell_type.items() if value.isnumeric()]
            mask = df.t == 's'
            shrd = df.loc[mask, 'val'].astype(int).map(lambda x: shared[x]).to_dict()

            # t = str (String): Cell containing a formula string.
            # t = inlineStr (Inline String): Cell containing an inline string.
            # mask = (df.t == 'str') | (df.t == 'inlineStr') 
            # df.loc[mask, 'val'] = df.loc[mask, 'val'].map(lambda x: x) # No se hace nada, ya que el valor ya está como cadena
            
            # t = e (Error): Cell containing an error.
            mask = df.t == 'e'
            errs = df.loc[mask, 'val'].map(lambda x: XlErrors(x)).to_dict()

            # t = b (Boolean): Cell containing a boolean.
            mask = df.t == 'b'
            bools = df.loc[mask, 'val'].astype(int).map(lambda x: x != 0).to_dict()

            # t = n (Number)
            mask = (df.t == 'n') | (df.t.isna())
            nums = {key: eval(val) for key, val in df.loc[mask, 'val'].to_dict().items()}

            values = {**nums, **bools, **errs, **shrd}


            # values = (
            #     df
            #     .drop(columns=['t', 's'])
            #     .rename(columns={'val': 'value', 'adr': 'address'})
            #     .set_index('address')
            #     .sort_index(key=lambda ndx: ndx.map(lambda x: '{1: >4s}-{0: >4s}'.format(*cell_pattern.match(x).groups())))
            #     .value
            #     .to_dict()
            #     # .items()
            # )
            # cell_errors = []
            # values = {}
            # for key, value in pairs.items():
            #     try:
            #         value = eval(value)
            #     except Exception as e:
            #         cell_errors.append(key)
            #         value = '#ERROR!'
            #     values[key] = value
            # if cell_errors:
            #     dmy = ', '.join(cell_errors)
            #     msg = f'Error calculating values in cells: {dmy}'
            #     logger.debug(msg)
        except Exception as e:
            msg = f'Error loading values: {str(e)} {val_regex}'
            logger.debug(msg)
            raise Exception(msg)
        return values

    def get_formulas(wb, content, ws_range, allCells=True):
        range_regex = regex_range(ws_range)
        seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
        fml_regex = f'(?#<c r="{range_regex}"=adr f.ref=_adr f.*=".+?"=fml>)'
        try:
            df = wb.extract_data(content, fml_regex, seek_str)
            fml_raw = (
                df
                .assign(fml=lambda db: db.fml.map(unescape))
                .set_index('adr')
                .fml.to_dict()
                .items()
            )
        except Exception as e:
            msg = f'Error loading values: {fml_regex}'
            logger.debug(msg)
            raise Exception(msg)
        

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
                        (
                            offset_rng(cell1, col_offset, row_offset),
                            wscell_pattern.sub(lambda m: offset_rng(m.group(), col_offset, row_offset), fml)
                        )
                        for col_offset, row_offset in offsets
                    ]
                    pairs.extend(fmls)
                else:
                    pairs.append((adr, fml))
        else:
            pairs = fml_raw
        fmls = {key: '=' + value.lstrip('+') for key, value in pairs}
        return fmls


class OXLLoader:

    def __init__(self, filename: str):
        self.filename = filename
        tmpdir = tempfile.gettempdir()
        dstfile = shutil.copy(filename, tmpdir)
        shutil.copystat(filename, dstfile)
        zf = zipfile.ZipFile(dstfile)
        self.zf = zf
        root = f'/{os.path.basename(filename)}'
        namelist = [root + '/' + x for x in zf.namelist()]
        navigationbar.StrListObj.SEP = '/'
        self.path_obj = navigationbar.StrListObj(namelist, root)
        pass

    def load_xmlfile(self, path, isAbs=True):
        if isAbs:
            rpath = self.path_obj.relpath(path, self.path_obj.root)
        else:
            rpath = path
        content = self.zf.read(rpath).decode('utf-8')
        return content
    

    

    



def load_workbook(filename):
    xml_book = WorkBookXml(filename)
    return xml_book


def main():
    fname = r'C:\Users\agmontesb\Documents\GitHub\excel\tests\files\excel_module_test.xlsx'
    test = 'interval_regex'  # 'WorkBookXml' | 'WorkSheetXml' | 'load_workbook'

    match test:
        case 'WorkBookXml':
            wb = WorkBookXml(fname)
            sheet_names = wb.sheetnames
            df = wb.data_in_range(sheet_names[0], 'E2:H16')
        case 'WorkSheetXml':
            wb = WorkBookXml(fname)
            sheet_names = wb.sheetnames
            ws = wb[sheet_names[0]]
            cells = ws['I2:I10']
            logger.debug(cells)
        case 'load_workbook':
            wb = load_workbook(fname)
        case 'interval_regex':
            answ = interval_regex('95', '1000')
        case _:
            pass
    pass

if __name__ == '__main__':
    main()


