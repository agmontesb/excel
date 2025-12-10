from html import unescape
import os
from typing import Any
import zipfile
import shutil
import tempfile
import itertools
import pandas as pd
import collections
import logging
import re

from worksheetui import SheetState, SheetContextData, COL_CELLS_WIDTH, ROW_CELLS_HEIGHT, CELL_WIDTH, CELL_HEIGHT
import mywidgets.Tools.uiStyle.MarkupRe as MarkupRe
from mywidgets.Widgets.Custom import navigationbar
from xlobjects import XlErrors
from xlpatterns import from_a1_tuple, offset_rng, interval_regex, regex_range
from xlpatterns import (cell_address as wscell_address,
                        cell_pattern as wscell_pattern,
                        code_alpha as excel_col_to_int,
                        tbl_address,
                        from_r1c1_a1, 
                        formulaR1C1)


logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# cell_pattern = re.compile(r'(\$?[A-Z]+)(\$?[0-9]+)')

DEFAULT_RANGE = 'A1:CV1000'   # equivalent to 'R1C1:R100C1000

class WorkSheetXml:
    Cell = collections.namedtuple('Cell', ['address', 'formula', 'formulaR1C1', 'value', 'style'], defaults=(None, None, None, None))

    def __init__(self, ws_name: str):   #, fml_map: dict, val_map: dict):
        self.title = ws_name
        self.coords_vportq3: tuple[int, int] = (COL_CELLS_WIDTH, ROW_CELLS_HEIGHT)
        self.coords_vportq1: tuple[int, int] = self.coords_vportq3        
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
    
    def cell(self, row:int, column:int, *args) -> 'WorkSheetXml.Cell':
        adr1 = from_r1c1_a1(f'R{row}C{column}')
        if not args:
            return self.cells.get(adr1, WorkSheetXml.Cell(adr1))
        adr2 = from_r1c1_a1(f'R{args[0]}C{args[1]}')
        cell_rng = f'{adr1}:{adr2}'
        rng_rgx = regex_range(cell_rng)
        pattern = re.compile(rng_rgx)
        answ = {key: value for key, value in self.cells.items() if pattern.fullmatch(key)}
        return answ


class WorkBookXml:

    def __init__(self, fname:str|None=None, default_range=DEFAULT_RANGE):
        self.zf = None
        self.ws_fname = {}
        self._defined_names = {}
        self._sheet_names = {f'Sheet{n}': '' for n in (1, 2, 3,)}
        self.sheet_obs = {}
        self.default_range = default_range
        self.shared_strings = []
        self.calcMode = 'auto'
        self.refMode = 'R1C1'
        ndx = 0
        if fname:
            ndx = self.load_wb_data(fname)
        active_sheet = self.sheetnames[ndx]
        self.select_sheet(active_sheet)

    def load_wb_data(wb, fname:str) -> str:
        wb.zf = loader = OXLLoader(fname)
        wb.shared_strings = loader.shared_strings
        wb_data = loader.parse_workbook_xlm()
        active_tab = wb_data.pop('active_tab', 0)
        for key, value in wb_data.items():
            setattr(wb, key, value) 
        return active_tab

    def select_sheet(wb, sheet_name: str) -> WorkSheetXml:
        wb.active = wb[sheet_name]

    @property
    def sheetnames(self):
        return list(self._sheet_names.keys())

    def __getitem__(wb, ws_name):
        assert ws_name in wb.sheetnames, "Not a valid Woksheet name"
        if ws_name not in wb.sheet_obs:
            ws = WorkSheetXml(ws_name)
            try:
                loader = wb.zf
                rpath = wb._sheet_names[ws_name]
                ws_data = loader.parse_worksheet_xlm(rpath, isAbs=False)
            except Exception as e:
                ws_data = vars(SheetContextData())
            for key, value in ws_data.items():
                setattr(ws, key, value)
            wb.sheet_obs[ws_name] = ws 
        return wb.sheet_obs[ws_name]


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

        content = self.load_xmlfile('xl/sharedStrings.xml', isAbs=False)
        self.shared_strings = MarkupRe.findall('(?#<t *=shared_str>)', content)
        pass

    def load_xmlfile(self, path, isAbs=True):
        if isAbs:
            rpath = self.path_obj.relpath(path, self.path_obj.root)
        else:
            rpath = path
        content = self.zf.read(rpath).decode('utf-8')
        return content
    
    @staticmethod
    def extract_data(content, regex_pattern, seek_pattern=None):
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
    
    def parse_workbook_xlm(loader) -> int:
        content = loader.load_xmlfile('xl/workbook.xml', isAbs=False)

        answ = {}

        # Active sheet
        wb_pattern = '(?#<workbookView activeTab=_tabId>)'
        cpattern = MarkupRe.compile(wb_pattern)
        active_tab = cpattern.findall(content)[0] or '0'
        answ['active_tab'] = int(active_tab)

        # Sheet names
        wb_pattern = '(?#<sheet name=name r:id=id>)'
        cpattern = MarkupRe.compile(wb_pattern)
        answ['_sheet_names'] = {key: f'xl/worksheets/sheet{id[3:]}.xml' for key, id in cpattern.findall(content)}

        # Named ranges
        wb_pattern = '(?#<definedName name=name *=rng>)'
        cpattern = MarkupRe.compile(wb_pattern)
        fn = lambda coords: sum([from_a1_tuple(cell)[1:] for cell in coords.split(':')], tuple())
        try:
            defined_names = {
                name: (tpl[0], fn(tpl[1])) 
                for name, rng in cpattern.findall(content)
                if (tpl := tbl_address(rng.replace('$', '')))
            }
            answ['_defined_names'] = defined_names
        except Exception as e:
            logger.debug(f'Error loading named ranges: {str(e)}')
            answ['_defined_names'] = {}
        
        # 18.2.2 calcPr (Calculation Properties)
        # calcMode (Calculation Mode)
        # refMode (Reference Mode)
        wb_pattern = '(?#<calcPr calcMode=_calcMode refMode=_refMode>)'
        cpattern = MarkupRe.compile(wb_pattern)
        calcMode, refMode = cpattern.findall(content)[0]
        answ['calcMode'] = calcMode or 'auto'
        answ['refMode'] = calcMode or 'A1'
        return answ

    def parse_worksheet_xlm(loader, rpath: str, isAbs: bool = False) -> dict[str, Any]:
        content = loader.load_xmlfile(rpath, isAbs)
        answ = {}

        # dimension element: 18.3.1.35 dimension (Worksheet Dimensions)
        rgx_str = '(?#<dimension ref=dim>)'
        ws_range = MarkupRe.findall(rgx_str, content)[0]
        answ['ws_range'] = ws_range

        # sheetView element: 18.3.1.86 sheetView (Sheet View)
        # When more than one sheet view is defined in the file, it means that when opening
        # the workbook, each sheet view corresponds to a separate window within the spreadsheet application, where
        # each window is showing the particular sheet containing the same workbookViewId value, the last sheetView
        # definition is loaded, and the others are discarded.
        rgx_str = '(?#<sheetViews <sheetView __LCHILD__="1">>)'
        m = MarkupRe.search(rgx_str, content)
        sheetview_str = m.group(0)
        rgx_str = '(?#<sheetView topLeftCell=_topLeftCell showGridLines=_showGridLines showRowColHeaders=_showRowColHeaders>)'
        topLeftCell, showGridLines, showRowColHeaders = MarkupRe.findall(rgx_str, sheetview_str)[0]
        topLeftCell = topLeftCell or 'A1'
        showGridLines = ['NONE', 'GRIDLINES'][showGridLines != '0']
        showRowColHeaders = ['NONE', 'HEADINGS'][showRowColHeaders != '0']
        flags = SheetState[showGridLines] | SheetState[showRowColHeaders]

        # 18.3.1.66 pane (View Pane)
        # This element only exists if the worksheet contains panes (split or frozen).
        pane_rgx = '(?#<pane xSplit=_xSplit ySplit=_ysplit (topLeftCell) (activePane) (state)>)'
        m = MarkupRe.search(pane_rgx, sheetview_str)
        if m:
            xSplit, ySplit, panetopLeftCell, activePane, state= m.groups()
            # Only consider when state is "frozen" or "frozenSplit", because in 
            # this version the split panes is not implemented
            viewport_q3 = (1, 1, 1, 1)
            if state != 'split':
                tl_corner = topLeftCell
                rb_corner = offset_rng(topLeftCell, *map(int, (xSplit, ySplit)))
                viewport_q3 = sum(map(lambda x: from_a1_tuple(x)[1:], (tl_corner, rb_corner)), tuple())
                flags |= SheetState.FREEZE
            viewport_q1 = 2 * from_a1_tuple(panetopLeftCell)[1:]
            rgx_str = f'(?#<selection pane="{activePane}" (activeCell) (sqref)>)'
        else:
            viewport_q3 = (1, 1, 1, 1)
            viewport_q1 = 2 * from_a1_tuple(topLeftCell)[1:]
            rgx_str = '(?#<selection (activeCell) (sqref)>)'

        answ['flags'] = flags
        answ['viewport_q1'] = viewport_q1
        answ['viewport_q3'] = viewport_q3

        # 18.3.1.78 selection (Selection)
        # Exist one (no freeze panes), two (topRow or leftColumn freeze) or three (freeze panes) 
        # of this element. the pane attib is optional.
        active_cell, sel = MarkupRe.findall(rgx_str, sheetview_str)[0]
        answ['active_cell'] = from_a1_tuple(active_cell)[1:]
        answ['selected_cells'] = sum(map(lambda x: from_a1_tuple(x)[1:], f'{sel}:{sel}'.split(':')[:2]), tuple())

         # Column and Row dimensions

        headings_dim = {}
        headings_hidden = {}

        # 18.3.1.13 col (Column Width & Formatting)
        rgx_str = '(?#<col (min) (max) (width) (customWidth) style=_style hidden=_hidden>)'
        for min, max, width, customWidth, style, hidden in MarkupRe.findall(rgx_str, content):
            min, max = int(min), int(max)
            # Suponiendo que customWidth es '1' siempre.
            headings_map = headings_hidden if hidden is not None else headings_dim
            # headings_map.update({f'C{n}': width for n in range(min, max + 1)})
            # Esto es mientras se implementa el width a pixeles
            headings_map.update({f'C{n}': CELL_WIDTH for n in range(min, max + 1)})
            # Por implementar: style, autoFit

        # 18.3.1.73 row (Row)
        rgx_str = '(?#<row r=row ht=_ht customHeight=_customHeight hidden=_hidden>)'
        for row, ht, customHeight, hidden in MarkupRe.findall(rgx_str, content):
            row = int(row)
            # Suponiendo que customHeight es '1' siempre.
            headings_map = headings_hidden if hidden is not None else headings_dim
            # Solo se almacena ht en headings_dim si customHeight es '1'
            if ht is not None:
                # headings_map[f'R{row}'] = ht
                headings_map[f'R{row}'] = CELL_HEIGHT
            elif hidden is not None:
                # headings_hidden[f'R{row}'] = '0' # Acá se debe almacenar el valor de altura de fila por defecto
                headings_hidden[f'R{row}'] = CELL_HEIGHT # Acá se debe almacenar el valor de altura de fila por defecto

        headings_dim.update({key: 0 for key in headings_hidden})
        answ['headings_dim'] = headings_dim
        answ['headings_hided'] = headings_hidden

        val_map, style_map = loader.get_values_styles(content, ws_range, allCells=True)
        fml_map, fmlr1c1_map = loader.get_formulas(content, ws_range, allCells=True)

        cells = {
            adr: WorkSheetXml.Cell(
                adr, 
                fml_map.get(adr, value),
                fmlr1c1_map.get(adr, value),
                value, 
                style_map.get(adr, None)
            )
            for adr in set(fml_map.keys()).union(val_map.keys())
            if (value :=val_map.get(adr)) is not None
        }
        answ['cells'] = cells
        return answ
    
    def get_values_styles(loader, content, ws_range: str, allCells=True):
        range_regex = regex_range(ws_range)
        seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
        val_regex = f'(?#<c r="{range_regex}"=adr t=_t s=_s v.*=val>)'
        if not allCells:
            val_regex = f'(?#<c r="{range_regex}"=adr __NCHILDREN__="2" t=_t s=_s v.*=val>)'
        shared = loader.shared_strings
        try:
            df = (
                loader.extract_data(content, val_regex, seek_str)
                .set_index('adr')
            )

            mask = df.s.notna()
            styles = df[mask].s.to_dict()
            
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

        except Exception as e:
            msg = f'Error loading values: {str(e)} {val_regex}'
            logger.debug(msg)
            raise Exception(msg)
        return values, styles

    def get_formulas(loader, content, ws_range, allCells=True):
        range_regex = regex_range(ws_range)
        seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
        fml_regex = f'(?#<c r="{range_regex}"=adr f.ref=_adr f.*=".+?"=fml>)'
        try:
            df = loader.extract_data(content, fml_regex, seek_str)
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
        fmlsr1c1 = {key: formulaR1C1(value, key) for key, value in fmls.items()}
        return fmls, fmlsr1c1


def load_workbook(filename: str|None=None) -> WorkBookXml:
    xml_book = WorkBookXml(filename)
    return xml_book


def main():
    fname = r'C:\Users\agmontesb\Documents\GitHub\excel\tests\files\excel_module_test.xlsx'
    test = 'interval_regex'  # 'WorkBookXml' | 'WorkSheetXml' | 'load_workbook'

    test = 'EmptyWorkBook'

    match test:
        case 'EmptyWorkBook':
            wb = WorkBookXml()
            sht1 = wb.select_sheet('Sheet1')
            pass
        case 'OXLLoader':
            loader = OXLLoader(fname)
            wb_data = loader.parse_workbook_xlm()
            tabId = wb_data['active_tab'] + 1
            ws_data = loader.parse_worksheet_xlm(f'xl/worksheets/sheet{tabId}.xml', isAbs=False)
            pass

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
            pass
        case 'interval_regex':
            answ = interval_regex('95', '1000')
        case _:
            pass
    pass

if __name__ == '__main__':
    main()


