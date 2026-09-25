from html import unescape
import os
from typing import Any, Generator
import zipfile
import shutil
import tempfile
import pandas as pd
from collections import namedtuple
import logging
import re
from concurrent.futures import ThreadPoolExecutor

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
BLK_SIZE = 200_000 # Tamaño de bloque para la búsqueda de patrones en el XML, en caracteres. Se puede ajustar según el tamaño del archivo y la memoria disponible.


class Cell(namedtuple('Cell', ['address', 'formula_', 'value', 'style', 'span'], defaults=(None, None, None, None))):
    __slots__ = ()

    @property
    def formula(self):
        fml = self.formula_ if self.formula_ is not None else str(self.value)
        return fml

    @property
    def formulaR1C1(self):
        # return formulaR1C1(self.formula_, self.address) if self.formula_ is not None else self.value
        fmlr1c1 = formulaR1C1(self.formula, self.address)
        return fmlr1c1
    
    def params(self) -> dict[str, Any]:
        return {
            'anchor': None,
            'fill': 'black', 
            'font': None,
            'justify': 'left',
            'stipple': '',
        }

    def __str__(self):
        val = self.value

        match val:
            case str():
                return val
            case XlErrors():
                return str(val)
            case bool():
                return str(val).upper()
            case None:
                return 'NONE'
            case _:
                return '{:,.2f}'.format(val)


class WorkSheetXml:
    # Cell = collections.namedtuple('Cell', ['address', 'formula', 'formulaR1C1', 'value', 'style'], defaults=(None, None, None, None))

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
        fnc = lambda x: '{0: >4s}{1: >4s}'.format(*wscell_address(x))
        return {key: cells[key] for key in sorted(cells.keys(), key=lambda x: fnc(x))}
    
    def cell(self, row:int, column:int, *args) -> 'WorkSheetXml.Cell':
        adr1 = from_r1c1_a1(f'R{row}C{column}')
        if not args:
            return self.cells.get(adr1, Cell(adr1))
        adr2 = from_r1c1_a1(f'R{args[0]}C{args[1]}')
        cell_rng = f'{adr1}:{adr2}'
        rng_rgx = regex_range(cell_rng)
        pattern = re.compile(rng_rgx)
        answ = {key: value for key, value in self.cells.items() if pattern.fullmatch(key)}
        return answ

    def iter_rows(self, min_row:int=None, min_col:int=None, max_row:int=None, 
                  max_col:int=None, values_only:bool=False) -> Generator:
        
        pass


class WorkBookXml:

    def __init__(self, fname:str|None=None, default_range=DEFAULT_RANGE):
        self.fname = fname
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

        self.shared_strings = self.parse_shared_strings()
        pass

    def load_xmlfile(self, path, isAbs=True):
        if isAbs:
            rpath = self.path_obj.relpath(path, self.path_obj.root)
        else:
            rpath = path
        content = self.zf.read(rpath).decode('utf-8')
        return content
    
    @staticmethod
    def extract_data_old(content, regex_pattern, seek_pattern=None):
        it = MarkupRe.compile(regex_pattern)
        if seek_pattern:
            it._seeker = re.compile(seek_pattern)
        values = []
        for grp_d in it.finditer(content):
            parameters = getattr(grp_d, 'parameters', [])
            items = [*grp_d.span(), *parameters, *grp_d.groupdict().values()]
            values.append(items)
        keys = ['beg', 'end', *it.get_seeker().groupindex.keys(), *it.groupindex.keys()]
        df = pd.DataFrame(values, columns=keys)
        return df

    @staticmethod
    def extract_data(content, regex_pattern, seek_pattern=None):

        def fn(comp_pattern: MarkupRe.ExtRegexObject, data:str, beg:int, end:int) -> pd.DataFrame:
            print(f'starting {beg}:{end}')
            values = []
            for grp_d in comp_pattern.finditer(data, beg, end):
                parameters = getattr(grp_d, 'parameters', [])
                items = [*grp_d.span(), *parameters, *grp_d.groupdict().values()]
                values.append(items)
            keys = ['beg', 'end', *comp_pattern.get_seeker().groupindex.keys(), *comp_pattern.groupindex.keys()]
            df = pd.DataFrame(values, columns=keys)
            print(f'finished {beg}:{end}')
            return df


        comp_pat = MarkupRe.compile(regex_pattern)
        if seek_pattern:
            comp_pat._seeker = re.compile(seek_pattern)

        # Se establecen las secciones de búsqueda para evitar que la búsqueda se realice en todo el contenido de una sola vez, 
        # lo que podría ser ineficiente para archivos grandes. Se divide el contenido en bloques y se buscan coincidencias dentro de esos bloques.

        tag_pattern = comp_pat.tag_pattern
        slice_str = f'<{tag_pattern}\\s[^>]*[^>]*[/]*>'
        slice_pat = MarkupRe.compile(slice_str)

        blks = [0]
        while blks[-1] < len(content):
            pbeg = min(blks[-1] + BLK_SIZE, len(content))
            m = slice_pat.search(content, pbeg)
            if m:
                blks.append(m.start())
            else:
                blks.append(len(content))
        blks = list(zip(blks[:-1], blks[1:]))

        with ThreadPoolExecutor() as executor:
            # submit() schedules the function and returns a Future object immediately
            futures = [executor.submit(fn, comp_pat, content, beg, end) for beg, end in blks]

            # .result() blocks the main program until that specific thread finishes
            results = [future.result() for future in futures]

        df = pd.concat(results, ignore_index=True)
        return df


    def parse_shared_strings(loader) -> list[str]:
        content = loader.load_xmlfile('xl/sharedStrings.xml', isAbs=False)
        answ = MarkupRe.findall('(?#<t *=shared_str>)', content)
        return answ
    
    def parse_styles(loader) -> list[str]:
        content = loader.load_xmlfile('xl/styles.xml', isAbs=False)

        numFmts = {
            int(key): unescape(value).split(';') 
            for key, value in MarkupRe.findall('(?#<numFmts <numFmt (numFmtId) (formatCode)>>)', content)
        }

        fnc = lambda m: {
            tpl[0]: tpl[1].strip('"') for x in (
            m.group()[6:-7]
            .replace(' val=', '=')
            .replace(' theme=', '=')
        )[1:-2].split('/><') if (tpl := x.split('='))
        }


        fonts = [
            fnc(m)
            for m in MarkupRe.finditer('(?#<fonts <font>>)', content)
        ]

        fnc = lambda m: { 
            tpl[0]: int(tpl[1]) 
            for x in m.group().strip('<xf />').split(' ')
            if (tpl := x.replace('"', '').split('='))
        }

        cellStyleXfs = [fnc(m) for m in MarkupRe.finditer('(?#<cellStyleXfs <__TAG__>>)', content)]

        cellXfs = [fnc(m) for m in MarkupRe.finditer('(?#<cellXfs <__TAG__>>)', content)]

        # apply mapping
        apply_map = {
            'applyAlignment': 'aligment',
            'applyBorder': 'borderId',
            'applyFill': 'fillId',
            'applyFont': 'fontId',
            'applyNumberFormat': 'numFmtId',
            'applyProtection': 'protection',
        }

        k = 9
        xf_rec = {}
        if (xfid := cellXfs[k].get('xfId', None)) is not None:
            xf_rec.update(
                {
                    key: cellStyleXfs[xfid].get(key, 1) 
                    for appkey, key in apply_map.items() 
                    if cellStyleXfs[xfid].get(appkey, 1)
                }
            )
        xf_rec.update(
            {
                key: cellXfs[k].get(key, 1)
                for appkey, key in apply_map.items() 
                if cellXfs[k].get(appkey, 1)
            }
        )
        return cellXfs

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

    def parse_worksheet_xlm(loader, rpath: str, isAbs:bool=False, ws_range:str=None) -> dict[str, Any]:
        content = loader.load_xmlfile(rpath, isAbs)
        answ = {}

        # dimension element: 18.3.1.35 dimension (Worksheet Dimensions)
        rgx_str = '(?#<dimension ref=dim>)'
        ws_range = ws_range or MarkupRe.findall(rgx_str, content)[0]
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

        cells_df = loader.get_cell_data(ws_range, content)
        style_map = loader.get_cells_style(cells_df, ws_range, content)
        val_map = loader.get_cells_value(cells_df, ws_range, content)
        fml_map = loader.get_cells_formula(cells_df, ws_range, content, allCells=True)

        cells = {
            adr: Cell(
                adr, 
                fml_map.get(adr, None),
                value, 
                style_map.get(adr, None),
                cells_df.loc[adr, ['beg', 'end']].to_list()
            )
            for adr in set(fml_map.keys()).union(val_map.keys())
            if (value :=val_map.get(adr)) is not None
        }
        answ['cells'] = cells
        return answ

    def get_cell_data(loader, ws_range: str, content: str, allCells=True) -> pd.DataFrame:
        # Cell data types:
        # value only:  '<c r="G12" s="242"><v>8706824</v></c>'
        # formula only: '<c r="K12" s="105"><f>G12+H12-I12+J12</f><v>8706824</v></c>'
        # No value: '<c r="AD11" s="235" t="str"><f>IF(OR(AA11="Error",AB11="Error",AC11="Error"),A11,"")</f><v/></c>'
        # Value with attrs: '<c r="AE93" s="58" t="str"><f>AA93 &amp; " | " &amp; AB93</f><v xml:space="preserve">OK | ok | ok | </v></c>'
        # shared_head: '<c r="H11" s="662"><f t="shared" ref="H11:K11" si="0">SUM(H12:H16)</f><v>0</v></c>'
        # shared: '<c r="I11" s="662"><f t="shared" si="0"/><v>0</v></c>

        range_regex = regex_range(ws_range)
        val_regex = fr'<c r="(?P<adr>{range_regex})"(?: s="(?P<style>\w+)")*( t="(?P<type>\w+)")*>(?:(?:<f t="shared" si="\d+"/>)|(?:<f[^>]*>(?P<fml>.+?)</f>))*(?:(?:<v/>)|(?:<v[^>]*>(?P<val>.+?)</v>))</c>'
        if not allCells:
            val_regex = f'(?#<c r="{range_regex}"=adr __NCHILDREN__="2" t=_t s=_s v.*=val>)'
        cpat = re.compile(val_regex)
        data = [(*m.span(), *m.groupdict().values()) for m in cpat.finditer(content)]
        try:
            df = (
                pd.DataFrame(data, columns=['beg', 'end', 'adr', 'style', 'vtype', 'fml', 'val'])
                .set_index('adr')
            )
        except Exception as e:
            msg = f'Error loading values: {str(e)} {val_regex}'
            logger.debug(msg)
            raise Exception(msg)
        return df
    

    def get_cell_data_old(loader, ws_range: str, content: str, allCells=True) -> pd.DataFrame:
        range_regex = regex_range(ws_range)
        seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
        val_regex = f'(?#<c r="{range_regex}"=adr t=_t s=_s v.*=val>)'
        if not allCells:
            val_regex = f'(?#<c r="{range_regex}"=adr __NCHILDREN__="2" t=_t s=_s v.*=val>)'
        try:
            df = (
                loader.extract_data(content, val_regex, seek_str)
                .set_index('adr')
            )
        except Exception as e:
            msg = f'Error loading values: {str(e)} {val_regex}'
            logger.debug(msg)
            raise Exception(msg)
        return df

    def get_cells_style(loader, cell_df: pd.DataFrame, ws_range:str, content:str=None, allCells=True) -> dict[str, str]:
        range_regex = regex_range(ws_range)
        style_rgx = fr'<c r="(?P<adr>{range_regex})" s="(?P<style>\w+)"/>'
        cpat = re.compile(style_rgx)
        try:
            pairs = cpat.findall(content)
        except Exception as e:
            msg = f'Error loading values: {str(e)} {style_rgx}'
            logger.debug(msg)
            raise Exception(msg)

        style_map = cell_df['style'].to_dict()
        style_map.update({key: value for key, value in pairs})
        return style_map

    def get_cells_value(loader, cell_df: pd.DataFrame, ws_range:str, content:str=None, allCells=True) -> dict[str, str]:
        shared = loader.shared_strings

        # values = df.set_index('adr').val.to_dict()
        # cell_type = df.set_index('adr').t.to_dict()
        # cell_style = df.set_index('adr').s.to_dict()


        # ECMA-376 18.18.11
        # t = s (Shared String): Cell containing a shared string.
        # [values.__setitem__(key, shared[int(value)]) for key, value in cell_type.items() if value.isnumeric()]
        mask = cell_df.vtype == 's'
        shrd = cell_df.loc[mask, 'val'].astype(int).map(lambda x: shared[x]).to_dict()

        # t = str (String): Cell containing a formula string.
        # t = inlineStr (Inline String): Cell containing an inline string.
        # mask = (df.t == 'str') | (df.t == 'inlineStr') 
        # df.loc[mask, 'val'] = df.loc[mask, 'val'].map(lambda x: x) # No se hace nada, ya que el valor ya está como cadena
        
        # t = e (Error): Cell containing an error.
        mask = cell_df.vtype == 'e'
        errs = cell_df.loc[mask, 'val'].map(lambda x: XlErrors(x)).to_dict()

        # t = b (Boolean): Cell containing a boolean.
        mask = cell_df.vtype == 'b'
        bools = cell_df.loc[mask, 'val'].astype(int).map(lambda x: x != 0).to_dict()

        # t = n (Number)
        mask = (cell_df.vtype == 'n') | (cell_df.vtype.isna())
        nums = {key: eval(val) for key, val in cell_df.loc[mask, 'val'].to_dict().items()}

        values = {**nums, **bools, **errs, **shrd}
        return values

    def get_cells_formula(loader, cell_df: pd.DataFrame, ws_range:str, content:str=None, allCells=True) -> dict[str, str]:
        cell1, cell2 = ws_range.split(':')
        (min_y, min_x), (max_y, max_x) = map(
                lambda x: (int(x[0]), excel_col_to_int(x[1])),
                map(wscell_address, (cell1, cell2))
        )

        fml_rgx = fr'<f t="shared" ref="(?P<adr>.+?)" si="\d+">(?P<fml>.+?)</f>'
        cpat = re.compile(fml_rgx)
        try:
            range_fmls = cpat.findall(content)
        except Exception as e:
            msg = f'Error loading values: {str(e)} {fml_rgx}'
            logger.debug(msg)
            raise Exception(msg)

        mask = cell_df.fml.notna()
        fml_map = (
            cell_df[mask]
            .fml
            .map(unescape)
            .to_dict()
        )
        pairs = []
        for (adr, fml) in range_fmls:
            fml = unescape(fml)
            cell1, cell2 = f'{adr}:{adr}'.split(':', 2)[:2]
            (linfy, linfx), (lsupy, lsupx) = map(
                    lambda x: (int(x[0]), excel_col_to_int(x[1])),
                    map(wscell_address, (cell1, cell2))
            )

            bflag = lsupy < min_y or max_y < linfy or lsupx < min_x or max_x < linfx
            if bflag:
                continue

            origen_x, origen_y = linfx, linfy
            linfy, linfx = max(linfy, min_y), max(linfx, min_x)
            lsupy, lsupx = min(lsupy, max_y), min(lsupx, max_x)
            offsets = [
                (col, row)
                for col in range(linfx - origen_x, lsupx - origen_x + 1)
                for row in range(linfy - origen_y, lsupy - origen_y + 1)
            ]
            fmls = [
                (
                    offset_rng(cell1, col_offset, row_offset),
                    wscell_pattern.sub(lambda m: offset_rng(m.group(), col_offset, row_offset), fml)
                )
                for col_offset, row_offset in offsets
            ]
            pairs.extend(fmls)
        fml_map.update(pairs) 
        return fml_map
    

def load_workbook(filename: str|None=None) -> WorkBookXml:
    xml_book = WorkBookXml(filename)
    return xml_book


def main():
    fname = r'C:\Users\agmontesb\Documents\GitHub\excel\tests\files\excel_module_test.xlsx'
    test = 'interval_regex'  # 'WorkBookXml' | 'WorkSheetXml' | 'load_workbook'

    test = 'OXLLoader'

    match test:
        case 'cell_display':
            import tkinter as tk
            app = tk.Tk()
            canvas = tk.Canvas(app, width=800, height=600, bg='white')
            canvas.pack()

            cell1 = Cell('A1', None, 1234.5678, None)
            cell2 = Cell('B2', None, 'Hello, World!', None)
            cell3 = Cell('C3', None, XlErrors.DIV_ZERO_ERROR, None)
            cell4 = Cell('D4', None, True, None)
            cell5 = Cell('E5', None, None, None)
            cells = [cell1, cell2, cell3, cell4, cell5]

            for k, cell in enumerate(cells):
                kwargs = cell.params()
                kwargs['text'] = str(cell)
                kwargs['width'] = 200
                kwargs['anchor'] = 'sw' if isinstance(cell.value, (int, float)) else 'ne'
                kwargs['font'] = ('Arial', 16)
                kwargs['fill'] = 'blue' if isinstance(cell.value, str) else 'red' if isinstance(cell.value, XlErrors) else 'green' if isinstance(cell.value, bool) else 'black'

                canvas.create_text(
                    100, 
                    50 * (k + 1), 
                    **kwargs,
                    # text=f'Cell {cell.address}: {str(cell)}', 
                    # anchor='w', 
                    # font=('Arial', 16)
                )
                         

            app.mainloop()
        case 'EmptyWorkBook':
            wb = WorkBookXml()
            sht1 = wb.select_sheet('Sheet1')
            pass
        case 'OXLLoader':
            loader = OXLLoader(fname)
            cellXfs = loader.parse_styles()
            wb_data = loader.parse_workbook_xlm()
            tabId = wb_data['active_tab'] + 1

            content = 1000 * loader.load_xmlfile(f'xl/worksheets/sheet{tabId}.xml', isAbs=False)

            range_regex = regex_range('A1:Z100')
            seek_str = f'<c\\s[^>]*r="{range_regex}"[^>]*[/]*>'
            val_regex = f'(?#<c r="{range_regex}"=adr t=_t s=_s v.*=val>)'

            df_old = loader.extract_data_old(content, val_regex, seek_str)
            df_new = loader.extract_data(content, val_regex, seek_str)


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


