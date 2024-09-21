import collections
import functools
import operator
import re
import warnings
from abc import ABC, abstractmethod
from idlelib.configdialog import changes

import pandas as pd
import numpy as np
from typing import Literal, Optional, Any
import itertools
from typing import Protocol, Type

from collections.abc import Callable


token_specification = [
    ('NUMBER', r'\d+(\.\d*)?'),  # Integer or decimal number
    ('ASSIGN', r'\='),  # Assignment operator
    ('OP', r'[+\-*/]'),  # Arithmetic operators
    ('COMMA', r','),  # Line endings
    ('ANCHOR', r'\:'),  # Line endings
    ('OPENP', r'\('),  # Line endings
    ('CLOSEP', r'\)'),  # Line endings
    ('SHEET', r"'.+'!"),  # Sheet names
    ('CELL', r'\$?[A-Z]\$?[1-9][0-9]*'),  # Identifiers
    ('FUNCTION', r'[A-Z]+'),  # Skip over spaces and tabs
    ('SKIP', r'[ ]+'),  # Skip over spaces and tabs
    ('MISMATCH', r'.'),  # Any other character
]
tokenizer = re.compile('|'.join('(?P<%s>%s)' % pair for pair in token_specification))

cell_pattern = re.compile(r"(?:'(.+?)'!)*(\$?[A-Z]+)(\$?[0-9]+)")
cell_address:Callable[[str], tuple[str, ...]] = lambda cell: tuple(x for x in cell_pattern.search(cell).groups()[::-1] if x)
rgn_pattern = re.compile(r"(?:'.+?'!)*\$?[A-Z]\$?[0-9]+(?::\$?[A-Z]\$?[0-9]+)?")
tbl_pattern = re.compile(r"(?:'(.+?)'!)*(.+)")
tbl_address = lambda tbl: tbl_pattern.match(tbl).groups()

ndx_sorter = lambda x: ('0000' + x.str.extract(r'(\d+)', expand=False)).str.slice(-4) + x.str.extract(r'([A-Z]+)', expand=False)

def excel_to_pandas_slice(excel_slice, axis, table_name, mask=None, withValues=False):
    sheet, lst_id = tbl_address(excel_slice)
    linf = f'{lst_id}:{lst_id}'.split(':', 1)[0]
    is_anchor = (linf[0] == '$' if axis == 0 else linf[1:].count('$') == 1)
    cell = row, col = cell_address(linf.replace('$', ''))
    mask = mask if not is_anchor else None
    if mask is None or (row not in mask and col not in mask):
        prefix = f"'{excel_slice}'" if sheet is None else f'"{excel_slice}"'
        pass
    elif row in mask or col in mask:
        infml = cell[col in mask]
        if sheet is None:
            prefix = ', '.join([f"'{excel_slice.replace(infml, key)}'" for key in mask])
        else:
            prefix = ', '.join([f'"\'{sheet}\'!{lst_id.replace(infml, key)}"' for key in mask])
    else:
        prefix = ''
    suffix = '.values' if withValues else ''
    py_term = f"{table_name}[{prefix}]{suffix}"
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
            case 'SHEET':
                lst_id = token_chr
            case 'CELL':
                prefix = token_chr.rstrip('0123456789')
                suffix = token_chr[len(prefix):]
                lst_id += f'{prefix}{suffix}'
                if ':' in lst_id or nxt_char != ':':
                    py_term = excel_to_pandas_slice(lst_id, axis, table_name, mask=mask, withValues='=' in pyfml)
                    pyfml += py_term
                    lst_id = ''
                    pass
            case 'SKIP':
                pass
            case _:
                pyfml += token_chr
    return pyfml


def pass_anchors(coord1, coord2):
    anchors = ['$' if x[0] == '$' else '' for x in cell_pattern.match(coord1).groups()[1:]]
    sheet_name, *coords = cell_pattern.findall(coord2.replace('$', ''))[0]
    sheet = '' if not sheet_name else f"'{sheet_name}'!"
    return sheet + ''.join(x for tup in zip(anchors, coords) for x in tup)


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


def data_in_range(openpyxl_sheet, ws_range:str):
    ws = openpyxl_sheet
    rgn_coords = coords_from_range(ws_range)
    rgn_fmls = {}
    rgn_values = {}
    rgn_map = [rgn_values, rgn_fmls]
    [
        operator.setitem(rgn_map[k], cell.coordinate, value.replace('=+', '=') if k == 1 else value)
        for row in ws.iter_rows(**rgn_coords)
        for cell in row
        if (value:=cell.value) and (k:=int(isinstance(value, str) and value[0] == '=')) < 2
    ]
    return rgn_fmls, rgn_values


class ExcelObject(ABC):
        _title: str
        id: Any
        parent: Optional['ExcelCollection']


        def __init__(self, object_name:str, *args, **kwargs):
            self._title = object_name
            pass

        @property
        def title(self) -> str:
            return self._title

        @title.setter
        def title(self, object_name:str):
            self._title = object_name
            pass


class ExcelCollection(ExcelObject, ABC):

    _title: str
    _objects: list[ExcelObject]
    object_class: Type[ExcelObject]
    parent: 'ExcelCollection'

    def __init__(self, parent, collection_name:str, obj_cls: Type[ExcelObject]):
        super().__init__(collection_name)
        self.parent = parent
        self._objects = []
        self.object_class = obj_cls
        pass

    @abstractmethod
    def next_id(self):
        pass

    def append(self, obj: ExcelObject):
        self._objects.append(obj)
        pass

    def index(self, excel_object: ExcelObject) -> int:
        tbl_name = excel_object.title
        return self.objectnames().index(tbl_name)

    def objectnames(self) -> list[str]:
        return [obj.title for obj in self._objects]

    def objects(self) -> list[ExcelObject]:
        return self._objects

    def __getitem__(self, object_name:str) -> ExcelObject:
        attr_name = 'id' if object_name.startswith('#') else 'title'
        object_name = object_name.strip('#')
        tbl = next(itertools.dropwhile(lambda x: getattr(x, attr_name) != object_name, self._objects))
        assert getattr(tbl, attr_name) == object_name, f'{object_name} not found'
        return tbl

    def move_object(self, object_name:str, ins_point:int = None):
        ndx = self.objectnames().index(object_name)
        excel_object = self._objects.pop(ndx)
        ins_point = ins_point or len(self._objects)
        self._objects.insert(ins_point, excel_object)
        pass

    def _create_object(self, object_name:str|None, ins_point:int|None = None) -> ExcelObject:
        obj = self.object_class(self, object_name)
        if ins_point:
            self.move_object(obj.title, ins_point)
        return obj


class ExcelWorkbook(ExcelCollection):

    def __init__(self, workbook_name: str):
        super().__init__(None, workbook_name, ExcelWorksheet)
        self.links = collections.defaultdict(set)
        self.parameters = collections.defaultdict(set)
        self.count = -1
        pass

    def next_id(self):
        self.count += 1
        return f'sheet{"" if self.count == 0 else self.count}'

    # def create_worksheet(self, sheet_name:str, ins_point:int = None) -> "ExcelWorksheet":
    #     return self._create_object(sheet_name, ins_point)

    create_worksheet = ExcelCollection._create_object

    sheetnames = property(ExcelCollection.objectnames)

    sheets = property(ExcelCollection.objects)


class ExcelWorksheet(ExcelCollection):

    def __init__(self, parent:ExcelWorkbook, sheet_name:str|None):
        id = parent.next_id()
        sheet_name = sheet_name or id
        super().__init__(parent, sheet_name, ExcelTable)
        self.parent.append(self)
        self.id = id
        self.count = -1
        self._param_map = {}
        pass

    def next_id(self):
        self.count += 1
        return chr(ord('A') + self.count)

    create_table = ExcelCollection._create_object

    @property
    def tablenames(self):
        return self.objectnames()

    @property
    def tables(self):
        return self._objects

    def remove(self, tbl:'ExcelTable'):
        del tbl

    def register_tbl(self, tbl: 'ExcelTable'):
        cells_to_link = self.parameters()
        if cells_to_link and (links_in_tbl := tbl.cells_in_data_rng(list(cells_to_link))):
            old_codes = [f"'{self.id}'!Z{cell}" for cell in links_in_tbl]
            code_links = [f"'{self.id}'!{cell}" for cell in tbl.data.loc[links_in_tbl, 'code']]
            changes = dict(zip(old_codes, code_links))
            self.broadcast_changes(changes, field='cell')
            links = self.parent.links
            for old_code, code in changes.items():
                links[code] = links.pop(old_code)
            [self._param_map.pop(key) for key in links_in_tbl]
        pass

    def register_links(self, xtable: 'ExcelTable', to_link: list[str]) -> tuple[list[str], list[str], list[Any]]:
        to_link = ['!'.join(f"'{self.title}'!{x}".split('!')[-2:]) for x in to_link]
        to_link = sorted(to_link, key=lambda x: '{2: >40s}!{0: >4s}{1}'.format(*cell_address(x)))

        links = self.parent.links
        answ = []
        for cell in to_link:
            sheet_name, cell_coord = tbl_address(cell)
            sheet: ExcelWorksheet = self.parent[sheet_name]
            tbl =sheet.associated_table(cell)
            try:
                code, value = tbl.data.loc[cell_coord, ['code', 'value']]
            except AttributeError:
                value = sheet._param_map.setdefault(cell_coord, 0)
                code = 'Z' + cell_coord
            code = f"'{sheet.id}'!{code}"
            cell = f"'{sheet.id}'!{cell_coord}" if sheet.id != self.id else cell_coord
            answ.append((cell, code, value))
            links[code].add(f"'{self.id}'!{xtable.id}")

        cells, code_links, values = map(list, zip(*answ))
        return cells, code_links, values

    def parameters(self, *args, **kwargs):
        '''
        Parameters are the links in the existing tables to this sheet cells that do not belong
        to any defined table.
        '''
        # If args are provided, return the corresponding values for the parameters
        if args:
            return [self._param_map.get(arg) for arg in args]
        # If keyword arguments are provided, update the values of the parameter and broadcast changes
        elif kwargs:
            keys = kwargs.keys() & self._param_map.keys()
            self._param_map.update((key, kwargs[key]) for key in keys)
            values = {f"'{self.id}'!Z{key}": kwargs[key] for key in keys}
            self.broadcast_changes(values, field='parameter')
        # If no arguments are provided, return the list of parameters of the sheet.
        else:
            return list(self._param_map.keys())

    def link_values(self, links: list[str]):
        value_items = []
        n = 0
        ws_name = self.title
        while links and n < len(self.tables):
            tbl: ExcelTable = self.tables[n]
            if cells:=tbl.cells_in_data_rng(links):
                value_items.extend(
                    (f"'{ws_name}'!{cell}", value)
                    for cell, value in
                    tbl.data.loc[cells, 'value'].to_dict().items()
                )
                links = list(set(links) - set(cells))
            n += 1
        if links:
            value_items.extend([(f"'{ws_name}'!{cell}", 0) for cell in links])
        value_items = sorted(value_items, key=lambda x: '{2: >40s}!{0: >4s}{1}'.format(*cell_address(x[0])))
        return list(zip(*value_items)[1])

    def broadcast_changes(self, changes: dict[str, ...], *, field: Literal['value', 'parameter', 'cell'] = 'value'):
        ws_id = self.id
        links = self.parent.links
        keys = list(changes.keys() & links.keys())
        if keys:
            tbl_df = (
                pd.DataFrame(
                    [{tbl: 1 for tbl in links[key]} for key in keys],
                    index=keys
                )
                .fillna(0)
            )
            columns = [key for key in sorted(tbl_df.columns, key=lambda x: tbl_address(x))]
            for tbl_addr in columns:
                if not any(mask:=tbl_df[tbl_addr] == 1):
                    continue
                cells = tbl_df[mask].index.tolist()
                ws_id, tbl_id = map(lambda x: f'#{x}', tbl_address(tbl_addr))
                ws = self.parent[ws_id]
                tbl = ws[tbl_id]
                match field:
                    case 'value' | 'parameter':
                        cells = tbl.encoder('decode', pd.Series(cells, index=cells)).to_dict()
                        mapper = {key: changes[cell] for cell, key in cells.items()}
                        tbl.set_values(mapper, field=field, recalc=True)
                    case _:  # field == 'cell'
                        mapper = {cell: changes[cell] for cell in cells}
                        tbl.set_field(mapper, field='code')
        pass

    def associated_table(self, cell:str) -> ExcelObject | None:
        tables: list[ExcelObject] = self._objects
        for tbl in tables:
            if cell in tbl:
                return tbl
        return None


class ExcelTable(ExcelObject):

    def __init__(self, parent:ExcelWorksheet, tbl_name:str|None, table_rng:str, fmls: dict[str,str]|None, values:dict[str, ...]|None, recalc:bool=False):
        super().__init__(tbl_name)
        self.parent: ExcelWorksheet = parent
        self.id = parent.next_id()
        self.data = None
        self.data_rng = table_rng
        self.needUpdate: bool = False
        self.changed = []
        parent.append(self)
        if fmls:
            self.set_fmls(fmls, values, recalc)
        pass

    def normalize_data(self, df):
        df = (
            df
            .assign(code=lambda db: [f'{self.id}{n}' for n in range(1, len(db.index) + 1)])
            .assign(fml= lambda db: self.encoder('encode', db['fml'], df=db))
            .assign(dependents=lambda db: db.dependents.apply(lambda x: {db.code[term] for term in x} if x is not np.nan else x))
        )
        return df

    def set_fmls(self, fmls: dict[str,str]|None, values:dict[str, ...]|None, recalc:bool=False):
        self.data = df = (
            pd.DataFrame.from_dict(fmls, orient='index', columns=['fml'])
            .pipe(self.fml_dependents)
            .pipe(self.res_order)
            .drop(columns='independents')
            .pipe(self.formula_type)
            .assign(value=0)
            .pipe(self.normalize_data)
        )
        df.index.rename('cell', inplace=True)
        self.parent.register_tbl(self)
        values = values or {}
        if links:=self.links():
            ws = self.parent
            links, code_links, link_values = ws.register_links(self, links)
            # Asignación de code para los enlaces (links)
            old_codes = df.loc[links, 'code'].tolist()
            changes = dict(zip(old_codes, code_links))
            self.set_field(changes, field='code')
            # Asignación de valores para los enlaces (links)
            df.loc[links, 'value'] = link_values
        # Se filtran values para tener en cuenta solo las claves que corresponden a entrada de datos en el rango
        # de la tabla
        # self.data = df.set_index('code', append=True)
        if values:
            self.set_values(values, recalc=recalc)

    def cells_in_data_rng(self, cells: list[str]) -> list[str]:
        (rmin, cmin),  (rmax, cmax) = map( cell_address, self.data_rng.split(':'))
        db = pd.Series(cells).str.extract(r'(?P<col>[A-Z]+)(?P<row>\d+)')
        mask = db.row.astype(int).between(int(rmin), int(rmax)) & db.col.between(cmin, cmax)
        return pd.Series(cells, index=mask).loc[True].tolist() if mask.any() else []

    def links(self):
        cells = self.data.index.tolist()
        cells_in_rgn = self.cells_in_data_rng(cells)
        cell_links = list(set(cells) - set(cells_in_rgn))
        return cell_links

    def __del__(self):
        ndx = self.parent.index(self)
        self.parent._objects.pop(ndx)

    def set_field(self, changes: dict[str, ...], *, field: Literal['code', 'cell']= 'code'):
        assert (df := self.data) is not None, 'Table not initialized'
        if field == 'code':
            old_codes = list(changes.keys())
            gmask = df.code.isin(old_codes)
            old_codes, dependents = zip(*df.loc[gmask, ['code', 'dependents']].values)
            code_links = [changes[key] for key in old_codes]
            for dependents_set, old_code, code in zip(dependents, old_codes, code_links):
                mask = df.code.isin(dependents_set)
                pattern = r"(?:'({0})'!)*(\$?{1})(\$?{2})".format(*cell_pattern.match(old_code).groups())
                df.loc[mask, 'fml'] = df[mask].fml.str.replace(
                    pattern,
                    lambda m: pass_anchors(m[0], code),
                    regex=True
                )
            df.loc[gmask, 'code'] = code_links
        else:  # field == 'cell'
            mask = df.code.isin(changes.keys())
            keys = df.loc[mask, 'code']
            df.rename(index={key: changes[cell] for key, cell in keys}, inplace=True)
        pass

    def set_values(self, values: dict[str, ...], field: Literal['value', 'parameter']='value', recalc:bool=False):
        assert (df := self.data) is not None, 'Table not initialized'
        mask = (df.res_order == 0) & df.index.isin(values.keys())
        [values.pop(key) for key in (values.keys() - df[mask].index)]
        changed = values.keys()
        if (parameters:= values.keys() & set(self.parent.parameters()) and field == 'values'):
            warnings.warn(
                f"You are modifying this parameter(s): {parameters}. "
                f"At this level it only modifies the table in which they are set."
                f"To be valid across the workbook, try at the worksheet level with the function "
                f"ws.parameters(param1=value1, param2=value2)",
                UserWarning
            )
        self.changed = df.loc[changed].code.tolist()
        if changed:
            df.loc[changed, 'value'] = list(values.values())
            self.recalculate(recalc)
        pass

    def get_cells_to_calc(self):
        to_process = set()
        changed, self.changed = self.changed, []
        for res_order, changed in self.cells_to_calc(changed):
            to_process.update(changed)
        return to_process

    def ordered_formulas(self, order, feval=False):
        assert self.data is not None, 'Table not initialized'
        gmask = (self.data.code.isin(order) & (self.data.res_order > 0)).tolist()
        df = (
            self.data
            .reset_index()
            .set_index('code')
        )
        formulas = []
        for (res_order, ftype), cells in df[gmask].groupby(['res_order', 'ftype']):
            cells.fml = self.encoder('decode', cells.fml, df=df)
            if ftype == '$' or len(cells) == 1:
                formulas.extend(
                    [
                        (*cell_address(rec.cell), rec.fml)
                        for code in cells.index.tolist()
                        if (rec := cells.loc[code]).any()
                    ]
                )
            else:
                mask = cells.cell.tolist()
                row, col = map(set, zip(*[cell_address(cell) for cell in mask]))
                ftype, cell_range = (row.pop(), sorted(col)) if len(row) == 1 else (col.pop(), sorted(row))
                formulas.append(
                    (ftype, cell_range, cells.iloc[0].fml)
                )
            pass

        pyfmls = []
        for frst_item, scnd_item, fml in formulas:
            pyfml = self.formula_translation(frst_item, scnd_item, fml)
            if feval:
                pyfml = '({0}, {1})'.format(*pyfml.replace('tbl', '', 1).split('=', 1))
            pyfmls.append(pyfml)
            # print(f'{frst_item} {scnd_item} {fml} ==> {pyfml}')

        return pyfmls

    def formula_translation(self, frst_item, scnd_item, fml, table_name='tbl'):
        '''
        Traduce una fórmula de Excel a una fórmula de Python
        :param fml: str. Fórmula de Excel
        :param table_name: str. Nombre de la tabla que contiene la fórmula
        :return: str. Fórmula de Python
        '''

        match (frst_item, scnd_item, fml):
            case (str(frst_item), str(scnd_item), fml):  # cell fml
                to_pythonize = f'{scnd_item}{frst_item}{fml}'
                mask = None
                axis = 0            # To avoid FutureWarning
            case (str(frst_item), list(scnd_item), fml) if frst_item.isnumeric():  # row fml
                to_pythonize = f'{scnd_item[0]}{frst_item}{fml}'
                mask = sorted(scnd_item)
                axis = 0
            case (str(frst_item), list(scnd_item), fml) if frst_item.isalpha():  # column fml
                to_pythonize = f'{frst_item}{scnd_item[0]}{fml}'
                mask = sorted([f'{row}' for row in scnd_item], key=lambda x: int(x))
                axis = 1
            case _:  # (list(frst_item), list(scnd_item), fml)    # range fml
                to_pythonize = ''
                mask = None
                axis = None
        return pythonize_fml(to_pythonize, table_name=table_name, axis=axis, mask=mask)

    def formula_type(self, df):
        order = (
            df
            .loc[df.res_order > 0, ['fml', 'res_order']]
            .reset_index(names='cell')
            .set_index('res_order')
            .sort_index()
        )

        def fml_equiv(m, item, frst_item):
            sheet, col, row = m.groups()
            if (frst_item.isnumeric() and row[0] == '$') or (frst_item.isalpha() and col[0] == '$'):
                return m[0]
            return m[0].replace(item, frst_item)

        rgn_fmls = lambda col, row: df.loc[f'{col}{row}' if col.isalpha() else f'{row}{col}', 'fml']

        formulas = []

        for res_order in order.index.unique():
            batch = order.loc[order.index == res_order, ['cell', 'fml']].values.tolist()
            batch = sorted(batch, key=lambda x: '{0: >4s}{1}'.format(*cell_address(x[0])))

            maps = [collections.defaultdict(list), collections.defaultdict(list), collections.defaultdict(list)]
            rmap, next = maps[:2]
            [
                rmap[row].append(col)
                for cell_coord, fml in batch
                if (cell_tple := cell_address(cell_coord)) and (row := cell_tple[0]) and (col := cell_tple[1])
            ]
            # Rows with only one associated col are transferred to the cols map.
            keys = list(rmap.keys())
            [next[col].append(row) for row in keys if len(rmap[row]) == 1 and (col := rmap.pop(row)[0])]

            for _ in range(2):
                for row in sorted(rmap.keys(), key=lambda x: len(rmap[x])):
                    columns = rmap.pop(row)
                    while columns:
                        frst_col, *columns = sorted(columns)
                        test_fml = rgn_fmls(frst_col, row)
                        mask = [
                            col
                            for col in columns
                            if test_fml == cell_pattern.sub(lambda m: fml_equiv(m, item=col, frst_item=frst_col),
                                                            rgn_fmls(col, row))
                        ]
                        if not mask:
                            next[frst_col].append(row)
                        else:
                            mask = [frst_col] + mask
                            formulas.extend([(f'{col}{row}' if col.isalpha() else f'{row}{col}', row) for col in mask])
                            columns = [col for col in columns if col not in mask]

                rmap, next = maps[1:]
                # cols with only one associated row are transferred to the single_list
                keys = list(rmap.keys())
                [
                    next[row].append(col)
                    for col in keys
                    if len(rmap[col]) == 1 and (row := rmap.pop(col)[0])
                ]

            formulas.extend([(f'{col}{row}', '$') for row, mask in next.items() for col in mask])
        ndx, ftype = zip(*formulas)
        df.loc[ndx, 'ftype'] = ftype
        return df

    def cells_to_calc(self, init_changed):
        df = (
            self.data
            .reset_index()
            .set_index('code')
        )
        to_report = df.loc[init_changed, ['res_order']].groupby(by='res_order').groups
        while True:
            try:
                k_min = min(to_report.keys())
            except ValueError:
                break
            changed = list(to_report.pop(k_min))
            changed = (yield (k_min, changed)) or changed
            if changed:
                dependents = df.loc[changed].dependents
                mask = ~dependents.isnull()
                if mask.any():
                    changed = list(functools.reduce(lambda t, e: t.union(e), dependents[mask], set()))
                    grouped_changed = df.loc[changed, ['res_order']].groupby(by='res_order').groups.items()
                    [to_report.setdefault(k, set()).update(v) for k, v in grouped_changed]

    def recalculate(tbl, recalc:bool=False):
        if recalc:
            reduce = lambda items: functools.reduce(lambda t, e: t.extend(e) or t, items, [])
            changed, tbl.changed = tbl.changed, []
            values = {}
            changed_cells, f_changed = tbl.cells_to_calc(changed), None
            while True:
                try:
                    res_order, changed = changed_cells.send(f_changed)
                except StopIteration:
                    break
                if not res_order:
                    mask = tbl.data.code.isin(changed)
                    before_values = dict(tbl.data.loc[mask, ['code', 'value']].values)
                    values.update(before_values)
                    continue
                formulas = tbl.ordered_formulas(changed, feval=True)
                items = eval(f"[{', '.join(formulas)}]", globals(), locals())
                cells, vals = map(reduce, zip(*[(cell, value.flatten()) for cell, value in items]))
                s1 = pd.Series(vals, index=cells)     # New values
                s2 = tbl.data.loc[cells].value          # old values
                mask = s1 != s2
                f_matched = s1[mask].index.tolist()
                vals = s1[f_matched]
                tbl.data.loc[f_matched, 'value'] = vals
                f_changed = tbl.data.loc[f_matched].code.tolist()
                values.update(zip(f_changed, vals))

            ws = tbl.parent
            values = dict((f"'{ws.id}'!{key}", value) for key, value in values.items())
            ws.broadcast_changes(values, field='value')
        tbl.needUpdate = not recalc

    def fml_dependents(self, df):
        rgn_fmls = df['fml'].to_dict()
        independents = {}
        for coord, fml in rgn_fmls.items():
            pairs = set()
            for term in rgn_pattern.findall(fml):
                term = term.replace('$', '')
                if ':' in term:
                    ws_name, cell_rng = f"!{term}".split('!')[-2:]
                    ws_name = f'{ws_name}!' if ws_name else ''
                    rgn_coords = coords_from_range(cell_rng)
                    pairs.update([f'{ws_name}{chr(ord("A") + col - 1)}{row}'
                                  for row in range(rgn_coords['min_row'], rgn_coords['max_row'])
                                  for col in range(rgn_coords['min_col'], rgn_coords['max_col'])
                                  ])
                else:
                    pairs.add(term)
            independents[coord] = pairs
        dependents = collections.defaultdict(set)
        for coord, terms in independents.items():
            for term in terms:
                dependents[term].add(coord)
        df = pd.concat(
            [
                df,
                pd.DataFrame(independents.items(), columns=['coord', 'independents']).set_index('coord'),
                pd.DataFrame(dependents.items(), columns=['coord', 'dependents']).set_index('coord')
            ],
            axis=1
        )
        return df

    def res_order(self, df):
        independents = dict(df['independents'].to_dict().items())
        pairs = set(
            (key, dep) for key, deps in independents.items()
            if isinstance(deps, set)
            for dep in deps
        )
        term_dep, term_ind = zip(*pairs)
        res_order = [set(term_ind) - set(term_dep)]
        counter = collections.Counter(term_dep)
        while pairs:
            # print(f'For process={len(pairs)}')
            to_batch = [
                (term, dep)
                for term, dep in pairs
                if dep in res_order[-1]
            ]
            counter.subtract([
                term
                for term, dep in to_batch
                if not pairs.remove((term, dep))
            ])
            to_batch = {term for term in counter.keys() if counter[term] == 0}
            [counter.pop(term) for term in to_batch]
            res_order.append(to_batch)
        res_order_df = pd.concat(
            pd.DataFrame(k, index=sorted(res_order[k]), columns=['res_order'])
            for k in range(len(res_order)
                           )
        )
        df = pd.concat(
            [
                df,
                res_order_df
            ],
            axis=1
        )
        return df

    def _cell_rgn(self, excel_slice: tuple[str, ...]):
        if isinstance(excel_slice, str):
            excel_slice = [excel_slice]
        cell_rgn = set()
        for slice in excel_slice:
            sheet, slice = tbl_address(slice)
            prefix = '' if sheet is None else f'\'{sheet}\'!'
            slice = slice.upper().replace('$', '')
            if ':' not in slice:
                cell_rgn.add(f'{prefix}{slice}')
                continue
            linf, lsup = slice.split(':')
            linf_cell = cell_address(linf)
            lsup_cell = cell_address(lsup)
            cell_rgn.update(
                [
                    f'{prefix}{col}{row}' for row, col in itertools.product(
                    [x for x in range(int(linf_cell[0]), int(lsup_cell[0]) + 1)],
                    [chr(x) for x in range(ord(linf_cell[1]), ord(lsup_cell[1]) + 1)])
                ]
            )
        cell_rgn = sorted(cell_rgn, key=lambda x: '{0: >4s}{1}'.format(*cell_address(x)))
        return cell_rgn

    def excel_table(self, data:pd.Series):
        return (
            data
            .set_axis(
                pd.MultiIndex.from_tuples(
                    [cell_address(cell)[:2] for cell in data.index],
                    names=['row', 'col']
                )
            )
            .unstack(level=-1)
            .sort_index(key=lambda x: x.str.extract(r'(\d+)', expand=False).astype(int))
        )

    def encoder(self, action: Literal['encode', 'decode'], in_fmls: list[str] | pd.Series, df: pd.DataFrame | None = None) -> pd.Series:
        if df is None:
            df = self.data
        if action == 'decode':
            df = df.reset_index().set_index('code')
            field = 'cell'
        else:
            field = 'code'
        if (bflag := isinstance(in_fmls, list)):
            in_fmls = pd.Series(in_fmls)
        out_fmls = in_fmls.str.replace(
            cell_pattern,
            lambda m: pass_anchors(m[0], df.loc[m[0].replace('$', ''), field]),
            regex=True
        )
        if bflag:
            out_fmls = out_fmls.tolist()
        return out_fmls

    def get_formula(self, *excel_slice: tuple[str, ...]):
        cell_rgn = self._cell_rgn(excel_slice)
        coded_fmls = self.data.loc[cell_rgn, 'fml']
        fmls = self.encoder('decode', coded_fmls)
        if len(cell_rgn) == 1:
            return fmls.iloc[0]
        return self.excel_table(fmls)

    def __getitem__(self, excel_slice: str|list[str]) -> pd.DataFrame:
        cell_rgn = self._cell_rgn(excel_slice)
        df = self.excel_table(self.data.loc[cell_rgn, 'value'])
        return df

    def __setitem__(self, excel_slice, value):
        cell_rgn = self._cell_rgn(excel_slice)
        self.data.loc[cell_rgn, 'value'] = np.array(value).flatten()

    def __contains__(self, item):
        row, col, sh_name = (cell_address(item) + (None, ))[:3]
        sh_name = sh_name or self.parent.title
        if (bflag := sh_name == self.parent.title):
            (rmin, cmin),  (rmax, cmax) = map( lambda x: (int((tpl:=cell_address(x))[0]), f'{tpl[1]: >2s}'), self.data_rng.split(':'))
            bflag = (rmin <= int(row) <= rmax) and (cmin <= f'{col.upper(): >2s}' <= cmax)
        return bflag

    def minimun_table(self):
        cell_rgn = self.cells_in_data_rng(self.data.index.tolist())
        mask = (self.data.index.isin(cell_rgn)) & (self.data.value != 0)
        excel_slice = self.data.loc[mask, :].index.tolist()
        df = self[excel_slice].fillna(0)
        return df