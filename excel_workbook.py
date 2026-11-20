from __future__ import annotations
import collections
import re
import warnings
from abc import ABC, abstractmethod
import pandas as pd
import numpy as np
from typing import Literal, Optional, Any, Generator, List, TypedDict, Sequence, Iterable
import itertools
from typing import Protocol, Type
from collections.abc import Callable
from enum import Enum, Flag
import logging

import excelxml
import xlfunctions as xlf
from xlobjects import *
from xlpatterns import (
    token_specification, tokenizer,
    link_pattern, 
    cell_pattern, cell_address,
    rgn_pattern_grp, rgn_pattern,
    tbl_pattern, tbl_address,
    code_alpha, alpha_code,
    MIN_ROWS, MIN_COLS,
    LINK_PREFIX, ERR_CODE_PREFIX
)

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


ndx_sorter = lambda x: ('0000' + x.str.extract(r'(\d+)', expand=False)).str.slice(-4) + x.str.extract(r'([A-Z]+)', expand=False)

flatten_sets = lambda items: set(itertools.chain(*items))

def excel_to_pandas_slice(excel_slice, axis, table_name, mask=None, withValues=False):
    """
    Transforms an Excel slice to a pandas slice.

    """
    """
    Transforms:
        F12 => "tbl['F12'].values"  (excel_slice='F12', axis=None, table_name='tbl', mask=None, withValues=True)
        F12 => "tbl['F12', 'F13', 'F14'].values"  (excel_slice='F12', axis=0, table_name='tbl', mask=['12', '13', '14'], withValues=True)    sheet, lst_id = tbl_address(excel_slice)
    """
    sheet, lst_id = tbl_address(excel_slice)
    linf = f'{lst_id}:{lst_id}'.split(':', 1)[0]
    row, col = cell_address(linf)
    is_anchor = (col[0] == '$' if axis == 0 else row[0] == 1)
    cell = row, col = row.replace('$', ''), col.replace('$', '')
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

def pythonize_fml(fml: str, table_name: str, axis: None|Literal[0,1]=None, mask=None):
    '''
    Convierte una fórmula de Excel a una fórmula de Python
    :param fml: str. Fórmula de Excel
    :param table_name: str. Nombre de la tabla que contiene la fórmula
    :param axis: int. Eje de la tabla que se está procesando
    :return: str. Fórmula de Python

    :Examples:
        '=DAYS360(F12, G12, TRUE)' => "=xlf.days360(tbl['F12'].values,tbl['G12'].values,True)"
        '=DAYS360(F12, G12)' => "=xlf.days360(tbl['F12'].values,tbl['G12'].values)"
        '=SUM(A1:A10) + #REF!' => "=xlf.sum(tbl['A1:A10'].values)+XlErrors.REF_ERROR"
    '''

    pyfml = ''
    lst_id = ''
    fnc_stack = []
    for mo in re.finditer(tokenizer, fml):
        kind = mo.lastgroup
        token_chr = mo.group()
        nxt_char = fml[mo.end():].lstrip(' ')[0] if mo.end() < len(fml) else ''
        # print(f'{kind=}: {token_chr=}')
        match kind:
            case 'FUNCTION':
                fnc_name = token_chr
                suffix = '_' if fnc_name in ('IF', 'AND', 'OR', 'ROUND') else ''
                fnc_stack.append(fnc_name)
                fnc_name = (fnc_name + suffix).lower()
                try:
                    getattr(xlf, fnc_name)
                except AttributeError:
                    logger.log(logging.ERROR, f'Function {fnc_name} ({token_chr}) not found in xlfunctions')
                pyfml += f'xlf.{fnc_name}'
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
                    # pyfml += f', axis={axis})' if fnc_name == 'SUM' else ')'
                    pyfml += ')'
                else:
                    pyfml += ')'
            case 'ANCHOR':
                lst_id += token_chr
            case 'BOOL':
                pyfml += token_chr.capitalize()
            case 'SHEET':
                lst_id = token_chr
            case 'ERROR':
                err_value = XlErrors(token_chr)
                err_str = f'XlErrors.{err_value.name}'
                py_term = err_str
                pyfml += py_term
                lst_id = ''
            case 'CELL':
                prefix = token_chr.rstrip('0123456789')
                suffix = token_chr[len(prefix):]
                lst_id += f'{prefix}{suffix}'
                if ':' in lst_id or nxt_char != ':':
                    py_term = excel_to_pandas_slice(lst_id, axis, table_name, mask=mask, withValues='=' in pyfml)
                    # py_term = excel_to_pandas_slice(lst_id, axis, table_name, mask=mask, withValues=True)
                    pyfml += py_term
                    lst_id = ''
                    pass
            case 'NUMBER':
                if nxt_char == '%':
                    token_chr = f'({token_chr}/100.0)'
                pyfml += token_chr
            case 'SOP':
                if token_chr == '^':
                    token_chr = '**'
                elif token_chr == '&':
                    token_chr = '+'
                elif token_chr == '<>':
                    token_chr = '!='
                pyfml += token_chr
            case 'SKIP':
                pass
            case _:
                pyfml += token_chr
    return pyfml


def pass_anchors(coord1, coord2):
    """
        ('A12', '$D2') => 'D2'
        ('A12', '$D$2') => 'D2'
        ('$A12', '$D$2') => '$D2'
        ('A$12', '$D$2') => 'D$2'
        ("'sheet1'!A$12", '$D2') => 'D$2'
        ("'sheet1'!A$12", "'sheet2'!$D2") => "'sheet2'!D$2"

    """
    anchors = ['$' if x[0] == '$' else '' for x in cell_pattern.match(coord1).groups()[1:]]
    sheet_name, *coords = cell_pattern.findall(coord2.replace('$', ''))[0]
    sheet = '' if not sheet_name else f"'{sheet_name}'!"
    return sheet + ''.join(x for tup in zip(anchors, coords) for x in tup)


def coords_from_range(rng:str) -> dict[str, int]:
    '''
    Encuentra el rango en que se revisaran las fórmulas.
    fml_col: str. Optional, si es None se supone que el rango de fórmulas es el que sigue al rng.
    rng: str. Rango de tabla que incluye concepto (ej: "B010:G227"
    
    Examples:
        ('A10:B15') => {'min_col': 1, 'min_row': 10, 'max_col': 3, 'max_row': 16}
        ('A10') => ValueError()
        ('$A10:B15') => ValueError()
        ('$A10:B$15') => ValueError()

    '''
    min_row, min_col, max_row, max_col = XlRef(rng).coords
    return dict(min_col=min_col, min_row=min_row, max_col=max_col + 1, max_row=max_row + 1)


def shrink_range(linf:str, lsup:str, changes:list[str]) -> list[str]:
    """
        Examples:
        parameters /=> known bug => answer 
        ('A5', 'F10', ['B', 'D']) => ['A5', 'D10']
        ('A5', '$F10', ['B', 'D']) => ['A5', '$F10']
        ('J9', 'O15', ['K', 'N']) => ['J9', 'M15']
        ('J9', '$O15', ['K', 'N']) /=> ['J9', '$O15']  => ['J9', '$M15'] 
        ('J9', 'O15', ['Q', 'R']) /=> ('J9', 'O15') => ['J9', 'O15']
        ('J9', 'O15', ['G', 'H']) /=> ['J9', 'M15'] => ['H9', 'M15']
        ('J9', '$O15', ['G', 'H']) /=> ('J9', '$O15') => ['H9', '$M15']
        ('$J9', '$O15', ['G', 'H']) /=> ('J9', '$O15') => ['$H9', '$M15']
        ('J9', 'O15', ['H', 'M', 'O']) /=> ['J9', 'L15'] => ['I9', 'M15']
        # Los elementos de changes no pueden incluir referencias a rangos de columns o rows
        ('J9', 'O10', ['J:L']) /=> TypeError() => ['J9', 'K9', 'L9', 'J10', 'K10', 'L10']
        ('J9', 'O15', list('HIJKLMNOPQ')) /=> ['J9', 'G15'] => '#REF!'
        # Se puede incluir la hoja del ancla:
        ("'sht1'!J9", "'sht2'!O15", ['J', 'K', 'L']) /=> ["'sht1'!J9", "'sht2'!L15"]
    """
    del_what = 'row' if changes[0].isnumeric() else 'col'
    bflag = del_what == 'col'
    k1, k2, fnc, gnc = (2, 3, code_alpha, alpha_code) if bflag else (3, 2, int, str)
    ucol = cell_pattern.match(lsup).group(k1)
    # lx_row = cell_pattern.match(changes[0]).group(k2)
    row_changes = [
        x for x in changes 
        if x <= ucol
    ]
    if not row_changes:
        return linf, lsup
    min_col = fnc(row_changes[0]) if linf in changes else (fnc(row_changes[0]) - 1)
    n = len(row_changes)
    answ = [linf]
    for x in (linf, lsup)[1 - int(linf in changes):]:
        tpl = list(cell_pattern.match(x).groups())
        tpl[2 - int(bflag)] = gnc(max(min_col, (fnc(tpl[2 - int(bflag)]) - n)))
        sht, col, row = tpl
        x = f"'{sht}'!{col}{row}" if sht else f"{col}{row}"
        answ.append(x)

    return answ[-2:]


def cells_in_range(ref_range: str, excel_slice: tuple[str, ...] | str) -> list[str]:
    """
        Lista las celdas en excel_slice (('I9:M15', 'I1:I5') => ['I1', 'I2', 'I3', 'I4', 'I5']). 
        ref_range se utiliza solamente para acotar referencias a columnas o filas (('I9:K11', 'I') => ['I9', 'I10', 'I11'])

        Examples:
        parameters /=> known bug => answer 
        ('I9:M15', 'J11') => ['J11']            # Forma básica
        # excel_slice acepta referencias de rango (incluye ':'), anchors ('$'), letras mayúsculas y minúsculas
        ('I9:M15', 'j11') => ['J11']            
        ('I9:M15', 'J$11') => ['J11']
        ('I9:M15', 'J11:K12') => ['J11', 'K11', 'J12', 'K12']
        ('I9:M15', 'J$11:K12') => ['J11', 'K11', 'J12', 'K12']
        ('I9:M15', '$J$11:$K$12') => ['J11', 'K11', 'J12', 'K12']
        # ref_range debe ser una referncia a rango (incluye ':') sin anchors, sin minúsculas
        ('I9:m15', 'J11') /=> AttributeError() => ['J11']
        ('I$9:M15', 'J11') /=> ValueError() => ['J11']
        # ref_range solo se utiliza cuando se tiene referencia a columnas('I') o filas ('10')
        ('I9:M15', 'I1:I5') => ['I1', 'I2', 'I3', 'I4', 'I5']
        ('I9:K11', 'I') => ['I9', 'I10', 'I11']
        ('I9:K11', 'I:J') => ['I9', 'J9', 'I10', 'J10', 'I11', 'J11']
        
    """
    parent_title, ref_range = tbl_address(ref_range)
    (rmin, cmin),  (rmax, cmax) = map( cell_address, ref_range.split(':'))
    rmin, rmax = int(rmin), int(rmax)

    if isinstance(excel_slice, str):
        excel_slice = (excel_slice, )
    cell_rgn = set()
    for slice in excel_slice:
        sheet, slice = tbl_address(slice)
        prefix = '' if sheet in (None, parent_title) else f'\'{sheet}\'!'
        slice = slice.upper().replace('$', '')
        linf, lsup = f'{slice}:{slice}'.split(':', 2)[:2]
        if (linf + lsup).isnumeric():
            if not prefix and (rmin <= int(linf) <= rmax) and (rmin <= int(lsup) <= rmax):
                linf = f'{cmin}{linf}'
                lsup = f'{cmax}{lsup}'
            else: 
                continue
        if (linf + lsup).isalpha():
            if not prefix and (cmin <= linf <= cmax) and (cmin <= lsup <= cmax):
                linf = f'{linf}{rmin}'
                lsup = f'{lsup}{rmax}'
            else:    
                continue
        linf_cell = cell_address(linf)
        lsup_cell = cell_address(lsup)
        fnc = code_alpha
        linf_cell, lsup_cell = map(lambda x: (int(x[0]), int(fnc(x[1]))), (linf_cell, lsup_cell))
        cell_rgn.update(
            [
                f'{prefix}{col}{row}' for row, col in itertools.product(
                [x for x in range(linf_cell[0], lsup_cell[0] + 1)],
                [alpha_code(x) for x in range(linf_cell[1], lsup_cell[1] + 1)])
            ]
        )
    cell_rgn = sorted(cell_rgn, key=lambda x: '{0: >4s}{1: >4s}'.format(*cell_address(x)))
    return cell_rgn


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

    def _create_object(self, object_name:str|None, *args, ins_point:int|None = None, **kwargs) -> ExcelObject:
        obj = self.object_class(self, object_name, *args, **kwargs)
        if ins_point:
            self.move_object(obj.title, ins_point)
        return obj
    
    @property
    def data(self):
        data = []
        for item in self._objects:
            df = item.data
            if 'tbl_id' not in df.columns:
                df = (
                    df
                    .assign(
                        tbl_id=f"'{self.id}'!{item.id}", 
                        dependents=lambda db: db.dependents.astype(str).apply(lambda x: x.replace("'", '"'))
                    )
            )

            data.append(df)
        return pd.concat(data)


class ExcelWorkbook(ExcelCollection):

    def __init__(self, workbook_name: str):
        super().__init__(None, workbook_name, ExcelWorksheet)
        self.links = collections.defaultdict(set)
        self.parameters = collections.defaultdict(set)
        self.count = 0
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
        self.count = 0
        self._param_map = {}
        self._data_rgn = None
        pass
   
    def reset_link(self, tbl: 'ExcelTable', link_code: str, reset_params=False):
        wb = self.parent
        sht, cell = tbl_address(link_code)
        ws = wb['#' + sht]

        if link_code in wb.links:
            wb.links[link_code].difference_update([f"'{self.id}'!{tbl.id}"])
            if not wb.links[link_code]:
                wb.links.pop(link_code)
                # Se trata de un parámetro:
                if reset_params and cell.startswith(LINK_PREFIX):
                    cell = cell[1:]
                    ws._param_map.pop(cell)

    def all_dependents(ws, to_process: dict['ExcelTable', set[str]], with_links=False) -> dict['ExcelTable', set[str]]:
        wb = ws.parent
        answer = {}
        while to_process:
            tbls = list(to_process.keys())
            while tbls:
                tbl = tbls.pop()
                codes = tbl.get_cells_to_calc(to_process.pop(tbl))
                answer.setdefault(tbl, set()).update(codes)
                ext_dep_df = tbl.external_dependents(codes)
                if ext_dep_df.empty:
                    continue
                for lnk, dep in list(ext_dep_df.dependent.items()):
                    sht_id, tbl_id = cell_pattern.match(dep).groups()[:2]
                    tbl = wb['#' + sht_id]['#' + tbl_id]
                    to_process.setdefault(tbl, set()).add(dep)
                    if with_links:
                        answer.setdefault(tbl, set()).add(lnk)
            common_keys = answer.keys() & to_process.keys()
            for key in common_keys:
                value = to_process.pop(key) - answer[key]
                if value:
                    to_process[key] = value
        return answer
    
    def propagate_circular_ref(ws, tbl: 'ExcelTable', circular_refs: list[str]):
        # Se marcan con circular reference todos los cell aguas arriba para las
        dependents = ws.all_dependents({tbl: set(circular_refs)}, with_links=True)
        # cells diferentes a la cell que establece la circulear_reference
        # además de los enlaces externos.
        inner_dependents = dependents.pop(tbl)
        dependents[tbl] = inner_dependents - set(circular_refs)
        for ftbl, fcodes in dependents.items():
            df = ftbl.data 
            mask = df.code.isin(fcodes)
            df.loc[mask, 'value'] = df.loc[mask].value.apply(lambda x: CIRCULAR_REF.get_instance(x))
        # cells que establecen la circular_reference
        mask = df.code.isin(circular_refs)
        tbl.data.loc[mask, 'value'] = CIRCULAR_REF

    def register_code(self, tbl, code_link, codes, default_value=None):
            df = tbl.data
            err_cell, err_code, err_value = next(zip(*self.register_links(tbl, [code_link], default_value=default_value)))
            try:
                error_rec = df.loc[err_cell].copy()
            except:
                # Si no existe código del error en la tabla, se crea un nuevo registro para el error
                error_rec = pd.Series(
                    dict(fml='', dependents=set(), res_order=0, ftype='$', value=err_value, code=err_code), 
                    dtype=object
                )
            # Se agrega el código de la celda al campo dependents del registro del error
            error_rec.dependents.update(tbl_address(x)[-1] for x in codes)
            df.loc[err_cell] = error_rec

    def unregister_code(self, tbl, code_link, codes, default_value=None):
            wb = self.parent
            df = tbl.data
            err_cell, err_code, err_value = next(zip(*self.register_links(tbl, [code_link], default_value=default_value)))
            error_rec = df.loc[err_cell].copy()
            # Se agrega el código de la celda al campo dependents del registro del error
            error_rec.dependents.difference_update(tbl_address(x)[-1] for x in codes)
            df.loc[err_cell] = error_rec
            if not error_rec.dependents:
                df.drop(err_cell)
                wb.links[err_code].difference_update([f"'{tbl.parent.id}'!{tbl.id}"])
                if not wb.links[err_code]:
                    wb.links.pop(err_code)
                    tbl.parent._param_map.pop(err_cell)

    def propagate_error(self, xl_error: XlErrors, codes: list[str] | None = None, reg_value=None):
        assert reg_value is None or not isinstance(reg_value, XlErrors), 'reg_value must be None or not a XlErrors'

        wb = self.parent
        to_process = {}
        if codes is None:
            assert xl_error is XlErrors.REF_ERROR
            # Se quiere registrar todos los códigos de las celdas que tienen un error
            # por el momento esto se hace solo en REF_ERROR ya que los otros salen del cálculo
            # de las celdas.
            for tbl in self.associated_table(xl_error.code, scope='parameter'):
                df = tbl.data
                dependents = df.loc[xl_error.code, 'dependents']
                to_process[tbl] = [f"'{tbl.parent.id}'!{x}" for x in dependents]
        else:
            # Si codes no es None, se debe registrar el codigo de la celda en los dependents de la
            # tabla correspondiente
            # Se clasifica la lista de códigos según la tabla a que pertenece.
            for code in codes:
                ws_id, tbl_id, _ = cell_pattern.match(code).groups()
                tbl = wb['#' + ws_id]['#' + tbl_id]
                # Se consigna para procesamiento en la clave de la tabla correspondiente
                to_process.setdefault(tbl, set()).add(code)
            
            fnc: Callable[[ExcelTable, XlErrors, Iterable[str]], None] = self.register_code if reg_value is None else self.unregister_code
            [fnc(tbl, xl_error.code, codes, default_value=xl_error) for tbl, codes in to_process.items()]

        propagate_error = self.all_dependents(to_process)
        reg_value = reg_value or xl_error
        rng_codes = None
        for tbl, codes in propagate_error.items():
            cells = [ExcelTable.encoder('decode', x, df=tbl.data) for x in codes]
            tbl.data.loc[cells, 'value'] = reg_value

    def insert(self, cell_slice=None, **kwargs):
        if not kwargs:
            excel_slice = f'{cell_slice}:{cell_slice}'.split(':', 2)[:2]
            slice_type = cell_slice.replace(':', '')
            is_numeric = slice_type.isnumeric()
            is_alpha = slice_type.isalpha()
            is_cell = is_numeric and is_alpha
            if is_cell:    # insert columns / rows
                return
            if is_numeric:    # insert rows
                field = 'row_offset'
                fnc = int
            else:                           # insert columns
                field = 'col_offset'
                fnc = code_alpha

            nitems = fnc(excel_slice[1]) - fnc(excel_slice[0]) + 1
            kwargs = {field: nitems}
            kwargs['disc_cell'] = f'A{excel_slice[0]}' if is_numeric else f'{excel_slice[0]}1'
            tbls = self.tables
        else:
            # Se filtran de las tablas entregadas solo aquellas que pertenecen a la hoja en que 
            # se delete filas o columnas.
            tbls = [tbl for tbl in kwargs.pop('tbls', []) if tbl in self.tables]
            # Se filtran de kwargs los parametros necesarios:
            required = set(['disc_cell', 'col_offset', 'row_offset'])
            present = kwargs.keys() & required
            assert len(present) == 2 and present != {'col_offset', 'row_offset'}, 'Required parameters in kwargs not present'
            kwargs = {key: kwargs[key] for key in present}
        ws = self
        wb = ws.parent


        for tbl in tbls:
            if (data_rng := tbl.offset_rng(tbl.data_rng, **kwargs, tbl=tbl)) == tbl.data_rng:
                continue

            df = tbl.data

            # modificación data range
            cells = tbl.cells_in_data_rng(df.index.tolist())
            tbl.data_rng = data_rng

            # modificación del índice del tbl.data frame (cells labels)
            cells_map = tbl.offset_rng(cells, **kwargs, tbl=tbl)
            tbl.data.rename(index=cells_map, inplace=True)

            # modificación enlaces externos a cells modificadas
            mask = df.index.isin(cells_map.values())
            external_links = wb.links.keys() & set(f"'{ws.id}'!{x}" for x in df.loc[mask].code)
            codes = [tbl_address(code)[-1] for code in external_links]
            codes_map = df.loc[df.code.isin(codes)].code.to_dict()

            to_broadcast = {
                f"'{ws.id}'!{code}": f"'{ws.title}'!{cell}" 
                for cell, code in codes_map.items()
            }
            ws.broadcast_changes(to_broadcast, field='cell')
        self._data_rgn = None

    def process_codes_to_delete(ws, to_process: dict['ExcelTable', set[str]], offset_kwargs):
        """

        """
        for tbl, codes in to_process.items():
            err_ref_cell, err_ref_code = next(
                zip(*ws.register_links(
                    tbl, [f"'{ws.title}'!{XlErrors.REF_ERROR.code}"], default_value=XlErrors.REF_ERROR
                    )
                )
            )[:2]
            df = tbl.data
            cells, codes = zip(*list(df.loc[df.code.isin(codes)].code.items()))

            # ** Modificación de fórmulas que están enlazadas con los códigos a eliminar **          
            direct_desc = flatten_sets(df.loc[list(cells)].dependents) - set(codes)
            mask = df.code.isin(direct_desc)
            fmls_df = df.loc[mask, ['fml', 'code']]

            # transformación fml a dominio cell
            code_ndx = dict(zip(df.code, df.index))
            fmls_df = tbl.set_field(code_ndx, field='code', data=fmls_df, with_dependents=False)

            # case_lst: lambda fnc que una lista de celdas las convierte en una lista de columnas
            # (kcase = 1) o una lista de filas (kcase = 2)
            # case_lst(['A1', 'C10', 'B5'], 1) = ['A', 'C', 'B']
            # case_lst(['A1', 'C10', 'B5'], 2) = ['1', '10', '5']

            k_case = 1 if 'col_offset' in offset_kwargs else 2
            case_lst = lambda lst, kcase: [
                tpl[kcase] for x in lst 
                if (tpl := cell_pattern.match(x).groups())
            ]

            changes = list(collections.Counter(case_lst(cells, k_case)).keys())
            err_ref = err_ref_cell
            err_sht_code, _ = tbl_address(err_ref_code)
            err_sht = None if tbl.parent.id == err_sht_code else ws.parent[f'#{err_sht_code}'].title
            pairs = []
            def replacement(m):
                if m[1] != err_sht:
                    # En el dominio cell, cuando se tiene sht in cell, esa cell no 
                    # está en la tabla a procesar
                    return m[0]
                if ':' in m[0]:
                    # Cuando se tiene un rango de celdas en la fórmula ('A20:D30') dos casos:
                    m = tbl_pattern.match(m[0])
                    prefix = f"'{m[1]}'!" if m[1] else ''
                    linf, lsup = map(lambda x: f"{prefix}{x}", m[0].replace('$', '').split(':'))
                    if all(x in changes for x in case_lst((linf, lsup), k_case)):
                        # 1 - Cuando los anclas ('A20' y 'D30') ambas se encuentran dentro de las códigos 
                        #     a eliminar entrega 'err_ref'
                        return err_ref
                    # 2 - Cuando alguna de las anclas se encuentra dentro de los códigos a eliminar.
                    #     encoge el rango ('shrink_range') y los almacena para procesamiento posterior.
                    left, right = shrink_range(linf, lsup, changes)
                    [pairs.append(tpl) for tpl in zip((linf, lsup), (left, right)) if tpl[0] != tpl[1]]

                    # Se deja para ser procesados posteriormente.
                    return m[0]
                # Cuando se tiene una referencia de celda ('D25'), si su código está dentro de los códigos
                # a eliminar, se entrega 'err_ref' en caso contrario no toma acción alguna.
                key = case_lst((m[0].replace('$', ''), ), k_case)[0]
                return err_ref if key in changes else m[0]

            fmls_df.loc[:, 'fml'] = fmls_df.fml.str.replace(
                rgn_pattern_grp,
                replacement,
                regex=True
            )
            if pairs:
                pairs = set(pairs)
                map_pairs = dict(tpl  if tpl[0] in cells else tpl[::-1] for tpl in pairs)
                fmls_df.loc[:, 'fml'] = fmls_df.fml.str.replace(
                    cell_pattern,
                    lambda m: pass_anchors(m[0], map_pairs.get(m[0].replace('$', ''), m[0])),
                    regex=True
                )

            # transformación fml a dominio code
            code_ndx = df.code.to_dict()
            code_ndx[err_ref_cell] = err_ref_code
            fmls_df = tbl.set_field(code_ndx, field='code', data=fmls_df, with_dependents=False)

            # Integramos la fórmulas modificadas al df
            df.loc[fmls_df.index, fmls_df.columns] = fmls_df

            # Modificación de códigos
            df.loc[list(cells), 'code'] = err_ref_code
            
            # Se eliminan en el campo dependents las referencias a los códigos a ser eliminados
            old_codes = set(codes)
            if (mask := df.dependents.apply(lambda x: bool(x & old_codes))).any():
                df.loc[mask, 'dependents'] = df.loc[mask].dependents.apply(lambda x: x - old_codes)
            tbl.data = df
            changes = {fcode: err_ref_code for code in codes if (fcode := code if '!' in code else f"'{ws.id}'!{code}" )}
            ws.broadcast_changes(changes, field='code')

        # Eliminamos de las tablas involucradas los códigos a eliminar 
        tbls = ws.associated_table(err_ref, scope='parameter')
        for tbl in tbls:
            df = tbl.data
            # Nos aseguramos que no existen duplicados
            nitems = df.code.tolist().count(err_ref_code)
            if nitems == 0:
                ws.reset_link(tbl, err_ref_code, reset_params=True)
            if nitems - int(err_ref in df) == 0:
                # No se tienen códigos de error o ya han sido compactados
                continue
            mask = df.code.eq(err_ref_code)
            to_recalc = flatten_sets(df.loc[mask].dependents.tolist())
            tbl.data = df = df.drop(df.index[mask], axis=0)

            mask = df.fml.str.replace('$', '').str.contains(err_ref_code)
            error_origens = df.code[mask].tolist()
            to_recalc = to_recalc - set(error_origens) 
            if error_origens:
                # mask = df.code.isin(error_origens)
                # tbl.data.value.loc[mask] = XlErrors.REF_ERROR
                # ws.register_code(tbl, err_ref, error_origens, default_value=XlErrors.REF_ERROR)
                ws.propagate_error(XlErrors.REF_ERROR, [f"'{tbl.parent.id}'!{code}" for code in error_origens])
            if tbl.parent.id == ws.id:
                kwargs = {'tbls':[tbl], **offset_kwargs}
                ws.insert(**kwargs)
            tbl.changed.extend(to_recalc)
            tbl.recalculate(tbl.recalc)
        pass

    def delete(self, cell_slice: str=None, **kwargs):
        if not kwargs:
            excel_slice = f'{cell_slice}:{cell_slice}'.split(':', 2)[:2]
            slice_type = cell_slice.replace(':', '')
            is_numeric = slice_type.isnumeric()
            is_alpha = slice_type.isalpha()
            is_cell = not (is_numeric or is_alpha)
            if is_cell:    # insert columns / rows
                # Procesar un rango de valores
                return
            (rmin, cmin),  (rmax, cmax) = map(cell_address, self.data_rng.split(':'))
            if is_numeric:    # insert rows
                field = 'row_offset'
                fnc = int
                cmin, cmax = map(code_alpha, (cmin, cmax))
                rmin, rmax = map(int, excel_slice)
            else:                           # insert columns
                field = 'col_offset'
                fnc = code_alpha
                cmin, cmax = map(code_alpha, excel_slice)
                rmin, rmax = map(int, (rmin, rmax))

            nitems = -(fnc(excel_slice[1]) - fnc(excel_slice[0]) + 1)
            kwargs = {field: nitems}
            disc_cell = f'A{excel_slice[0]}' if is_numeric else f'{excel_slice[0]}1'
            kwargs['disc_cell'] = ExcelTable.offset_rng(disc_cell, **{field: -1})
            tbls = self.tables
            all_cells = [f'{alpha_code(col)}{row}' for col in range(cmin, cmax + 1) for row in range(rmin, rmax + 1)]
        else:
            # Se filtran de las tablas entregadas solo aquellas que pertenecen a la hoja en que 
            # se delete filas o columnas.
            tbls = [tbl for tbl in kwargs.pop('tbls', []) if tbl in self.tables]
            # Se filtran de kwargs los parametros necesarios:
            required = set(['disc_cell', 'col_offset', 'row_offset', 'cell_slice'])
            present = kwargs.keys() & required
            assert len(present) == 3 and present & {'disc_cell', 'cell_slice'}, 'Required parameters in kwargs not present'
            kwargs = {key: kwargs[key] for key in present}
            cell_slice = kwargs.pop('cell_slice')
            xlslice = XlRef(cell_slice)
            all_cells = xlslice.cells()

        ws = self
        wb = ws.parent

        to_process = {}
        for tbl in tbls:
            cells = tbl.cells_in_data_rng(all_cells)
            if not cells:
                continue
            cells = list(set(cells) & set(tbl.data.index))
            df = tbl.data
            codes = df.loc[cells].code.tolist()
            to_process[tbl] = set(codes)

        params = [cell for cell in ws.parameters() if cell in all_cells]
        for param in params:
            pcode = ws.parameter_code(param)
            tbls = ws.associated_table(param, scope='parameter')
            # Se agrega una sola vez ya que posteriormente con broadcast_changes, se borra en las otras tablas
            to_process.setdefault(tbls[0], set()).add(pcode)

        # all_dependents = self.all_dependents(to_process=to_process.copy())
        # [all_dependents[tbl].difference_update(to_process[tbl]) for tbl in to_process]

        self.process_codes_to_delete(to_process, kwargs)

        # self.insert(cell_slice, from_delete=True)
        # self.propagate_error(XlErrors.REF_ERROR)

        # for tbl, codes in all_dependents.items():
        #     data = tbl.data
        #     df = data.loc[data.code.isin(codes)]
        #     codes = df.loc[df.value.isin(list(XlErrors))].code.tolist()
        #     dependents = flatten_sets(df.dependents.tolist()) | set(df.code)
        #     changed = dependents - set(codes)
        #     if changed:
        #         tbl.changed.extend(changed)
        #         tbl.recalculate(recalc=True)

        pass

    @property
    def data_rng(self):
        if self._data_rgn is None:
            cells = list(map(cell_address, [x for x in self._param_map.keys() if not x.startswith(LINK_PREFIX)]))
            for tbl in self.tables:
                cells.extend(map(cell_address, tbl.data_rng.split(':')))
            rows, cols = zip(*cells)
            rows = [int(x) for x in rows]
            cols = [f'{x: >2s}' for x in cols]
            r_min, r_max = min(rows), max(rows)
            c_min, c_max = map(str.strip, (min(cols), max(cols)))
            self._data_rng = f'{c_min}{r_min}:{c_max}{r_max}'
        return self._data_rng
    
    def _repr_html_(self):
        (rmin, cmin),  (rmax, cmax) = map(cell_address, self.data_rng.split(':'))
        rmin, rmax = int(rmin), int(rmax)
        cmin, cmax = code_alpha(cmin), code_alpha(cmax)
        all_cells = [f'{alpha_code(col)}{row}' for row in range(rmin, rmax + 1) for col in range(cmin, cmax + 1)]
        t = pd.Series('', index=all_cells)
        if (params := [x for x in self.parameters() if not x.startswith(LINK_PREFIX)]):
            t[params] = self.parameters(*params)
        for tbl in self.tables:
            cells_rng = tbl._cell_rgn(tbl.data_rng)
            mask = tbl.data.index.isin(cells_rng)
            t[tbl.data.index[mask]] = tbl.data.value[mask]
        data = ExcelTable.excel_table(t)
        return data._repr_html_()


    def next_id(self):
        self.count += 1
        return alpha_code(self.count)

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
            old_codes = [self.parameter_code(cell) for cell in links_in_tbl]
            code_links = [self.parameter_code(cell) for cell in tbl.data.loc[links_in_tbl, 'code']]
            changes = dict(zip(old_codes, code_links))
            self.broadcast_changes(changes, field='cell')
            links = self.parent.links
            for old_code, code in changes.items():
                links[code] = links.pop(old_code)
            [self._param_map.pop(key) for key in links_in_tbl]
        pass

    def register_links(self, xtable: 'ExcelTable', to_link: list[str], default_value=0) -> tuple[list[str], list[str], list[Any]]:
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
                pvalue = default_value
                value = sheet._param_map.setdefault(cell_coord, pvalue)
                code = LINK_PREFIX + cell_coord
            code = f"'{sheet.id}'!{code}"
            cell = f"'{sheet.title}'!{cell_coord}" if sheet.id != self.id else cell_coord
            answ.append((cell, code, value))
            links[code].add(f"'{self.id}'!{xtable.id}")

        cells, code_links, values = map(list, zip(*answ))
        return cells, code_links, values
    
    def parameter_code(self, param_name:str) -> str:
        return f"'{self.id}'!Z{param_name}"

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
            values = {self.parameter_code(key): kwargs[key] for key in keys}
            self.broadcast_changes(values, field='parameter')
        # If no arguments are provided, return the list of parameters' sheet.
        else:
            return list(self._param_map.keys())

    def cell_values(self, input_cells: list[str]):
        value_items = []
        s = pd.Series(input_cells)
        sheet_groups = (
            s
            .where(s.str.contains('!'), f"'{self.title}'!" + s)
            .str.extract(tbl_pattern, expand=True)
            .set_index('cell')
            .groupby(by='sht')
            .groups
            )
        
        for sht, items in sheet_groups.items():
            n = 0
            ws = self.parent[sht]
            ws_name = ws.title
            links = items.tolist()
            while links and n < len(ws.tables):
                tbl: ExcelTable = ws.tables[n]
                if cells:=tbl.cells_in_data_rng(links):
                    value_items.extend(
                        (f"'{ws_name}'!{cell}", value)
                        for cell, value in
                        tbl.data.loc[cells, 'value'].to_dict().items()
                    )
                    links = list(set(links) - set(cells))
                n += 1
            if (parameters := set(links) & set(ws.parameters())):
                pvalues = ws.parameters(*parameters)
                value_items.extend([(f"'{ws_name}'!{cell}", value) for cell, value in zip(parameters, pvalues)])
            if (links := set(links) - set(parameters)):
                value_items.extend([(f"'{ws_name}'!{cell}", None) for cell in links])
        value_items = sorted(value_items, key=lambda x: '{2: >40s}!{0: >4s}{1}'.format(*cell_address(x[0])))
        return map(list, zip(*value_items))

    def broadcast_changes(self, changes: dict[str, Any], *, field: Literal['value', 'parameter', 'cell', 'code'] = 'value'):
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
                if not (mask:=tbl_df[tbl_addr] == 1).any():
                    continue
                cells = tbl_df[mask].index.tolist()
                ws_id, tbl_id = map(lambda x: f'#{x}', tbl_address(tbl_addr))
                ws = self.parent[ws_id]
                tbl = ws[tbl_id]
                match field:
                    case 'value' | 'parameter':
                        cells = tbl.encoder('decode', pd.Series(cells, index=cells), df=tbl.data).to_dict()
                        mapper = {key: changes[cell] for cell, key in cells.items()}
                        # tbl.set_values(mapper, field=field, recalc=True)
                        tbl.set_records(mapper, field=field)
                        tbl.recalculate(True)
                    case 'cell':  # field == 'cell'
                        mapper = {cell: changes[cell].replace(f"'{ws.title}'!", '') for cell in cells}
                        tbl.data = tbl.set_field(mapper, field='cell', data=tbl.data)
                    case 'code':  # field == 'code'
                        mapper = {cell: changes[cell] for cell in cells}
                        tbl.data = tbl.set_field(mapper, field='code', data=tbl.data)
                        for old_code, new_code in mapper.items():
                            ws.reset_link(tbl, old_code, reset_params=True)
                            links[new_code].add(tbl_addr)
                    case _:
                        pass
        pass

    def associated_table(self, cell:str, scope:Literal['cell', 'parameter']='cell') -> ExcelObject| list[ExcelObject] | None:
        match scope:
            case 'cell':
                tables: list[ExcelObject] = self._objects
                for tbl in tables:
                    if cell in tbl:
                        return tbl
            case 'parameter':
                wb = self.parent
                code = self.parameter_code(cell)
                tbl_codes = self.parent.links[code]
                tbl_pairs = [tbl_address(tbl_code) for tbl_code in tbl_codes]
                tbls = [wb[f'#{sht_id}'][f'#{tbl_id}'] for sht_id, tbl_id in tbl_pairs]
                return tbls
        return None


class ExcelTable(ExcelObject):

    def __init__(self, parent:ExcelWorksheet, tbl_name:str|None, table_rng:str, fmls: dict[str,str]|None=None, values:dict[str, Any]|None=None):

        super().__init__(tbl_name)
        self.parent: ExcelWorksheet = parent
        self.id = parent.next_id()
        self.count = 0
        self.data = self.create_dataframe()
        self.data_rng = table_rng
        self.changed = []
        parent.append(self)
        if fmls:
            self.set_formulas(fmls)
        if values:
            self.set_values(values)
        self.pack_data()
        self.recalculate(recalc=True)
        pass

    @property
    def recalc(tbl) -> bool:
        """
        Check if the table has any changed cells that require recalculation.
        """
        return bool(tbl.changed)

    @classmethod
    def create_dataframe(cls, **kwargs) -> pd.DataFrame:
        # Create empty DataFrame with specified index and columns
        structure: DataFrameStructure = {'index': [str], 'columns': TABLE_DATA_MAP}
        df = (
            pd.DataFrame(
                index=pd.Index([], name='cell', dtype=structure['index'][0]),
                columns=structure['columns'].keys()
            )
            .astype(structure['columns'])
        )
        return df
    
    def set_records(tbl, values: dict[str, str] | dict[str, Any], field:Literal['fml', 'value', 'parameter']='fml'):
        """
        Set the values of a specific field ('fml' or 'value') for the given cells in the table.
        values: A dictionary where keys are cell addresses and values are the corresponding values to set.
        field: The field to set ('fml' for formulas, 'value' for values, 'parameter' for parameters).
        """
        def cell_list(fml):
            if not fml:
                return []
            fml = fml.replace('$', '')
            it = rgn_pattern.finditer(fml)
            components = set()
            for m in it:
                sht, xlslice = tbl_pattern.match(m[0]).groups()
                prefix = f"'{sht}'!" if sht else ''
                components.update(prefix + x for x in xlslice.split(':'))
            return components

        values = excelxml.OXLLoader.plain_data(values, field=field)
            
        keys = tbl.cells_in_data_rng(values.keys()) if field != 'parameter' else list(values.keys())
        if not keys:
            return

        # Nos aseguramos que la fórmula no tenga el nombre de la hoja, ya que se asume que 
        # la fórmula es relativa a la hoja en la que se encuentra la tabla.
        values = {key: values[key] for key in keys}
        if field == 'fml':
            sht_name = f"'{tbl.parent.title}'!"
            values = {key: values[key].replace(sht_name, '') for key in values}

        # celdas aguas abajo: celdas que tienen que ser calculadas antes que una celda determinada
        # celdas aguas arriba: celdas que tienen que ser calculadas después de una celda determinada

        df = tbl.data

        # MODIFICACIONES A LA BASE DE DATOS DE LA TABLA
        # 1 - Records con fórmulas existentes que van a ser redefinidas por los valores recibidos
        if (to_convert := list(set(keys) & set(df.loc[~df.fml.eq('')].index))):
            df = tbl.clear_cells(to_convert, data=df)

        # 2 - Records no existentes que van a ser creados por los valores recibidos
        if (to_init := list(set(keys) - set(df.index))):
            df = tbl.add_empty_cells(to_init, data=df)

        # 3 - Cuando field == 'fml', se agregan records correspondientes a celdas vacías en la 
        # tabla o a enlaces externos/parámetros.
        circular_refs = {}
        if field == 'fml':
            cells_in_fmls = set(itertools.chain(*[cell_list(x) for x in values.values()]))
            cell_links = cells_in_fmls - set(df.index)
            if cell_links:
                # Se ubican las celdas vacías que enlazan las fórmulas
                empty_cells = tbl.cells_in_data_rng(list(cell_links))
                links = list(cell_links.difference(empty_cells))
                if empty_cells:
                    # Solo se crean las celdas vacías que aparecen en la fórmula para efectos de 
                    # que existan los códigos al momento de codificarlas.
                    df = tbl.add_empty_cells(empty_cells, data=df, ftype='?')
                # Enlaces externos/parámetros
                if links:
                    ws = tbl.parent
                    cell_links, code_links, value_links = ws.register_links(tbl, links, default_value=EMPTY_CELL)
                    # ref_cells: Records que existen para asegurar el código de una celda, no están enlazadas.
                    ref_cells = [cell for cell, value in zip(cell_links, value_links) if value == EMPTY_CELL]
                    df = tbl.add_empty_cells(ref_cells, data=df, ftype='?')
                    true_cells = list(set(cell_links) - set(ref_cells))  # Records de celdas que existen
                    df = tbl.add_empty_cells(true_cells, data=df)
                    df.loc[cell_links, ['code', 'value']] = list(zip(code_links, value_links))

            # Se agregan a los dependents de las variables independientes de las fórmulas el 
            # 'code' correspondiente de la celda a recibir la fórmula.
            # for code, cells in value_cells.to_dict().items():
            # all_up_codes = set()
            for cell, fml in values.items():
                up_codes = tbl.get_cells_to_calc([cell], data=df, isCoded=False)
                cells_in_fml = set(itertools.chain(*[XlRef(x).cells() for x in rgn_pattern.findall(fml)]))
                fcells = list(cells_in_fml & set(df.index))
                code = df.loc[cell].code
                bExistCircRef = cells_in_fml & set(up_codes)
                if bExistCircRef:
                    df.loc[cell, 'res_order'] = 1
                    circular_refs[code] = fcells
                    continue
                # all_up_codes.update(up_codes)
                df.loc[fcells, 'dependents'] = df.loc[fcells].dependents.apply(lambda x: x | {code} if isinstance(x, set) else {code})
            vals = {key: tbl.encoder('encode', fml, df=df) for key, fml in values.items()}
        else:   # field == 'value' | 'parameter'
            if (parameters:= values.keys() & set(tbl.parent.parameters()) and field == 'value'):
                warnings.warn(
                    f"You are modifying this parameter(s): {parameters}. "
                    f"At this level it only modifies the table in which they are set."
                    f"To be valid across the workbook, try at the worksheet level with the function "
                    f"ws.parameters(param1=value1, param2=value2)",
                    UserWarning
                )
            field = 'value'
            vals = values
        
        # Se actualizan el campo de las celdas a modificar
        df.loc[vals.keys(), field] = ese = pd.Series(vals, name=field, dtype=TABLE_DATA_MAP[field])
        
        bFlag = field == 'value' and to_convert
        df.loc[vals.keys(), 'ftype'] = '$' if field == 'fml' else '#'

        encoded_keys = ese.rename(index=df.loc[ese.index].code)
        changed = list(set(encoded_keys.index) - set(circular_refs.keys()))

        # En este punto verificamos si se tiene que recalcular el campo 'res_order' para las 
        # celdas involucradas, lo cual se da en los siguientes casos:
        # 1. Si se van a introducir fórmulas en las celdas afectadas (field == 'fml')
        # 2. Si se van a introducir valores en las celdas afectadas (field == 'value')
        # pero alguna de ellas tenía una fórmula (bFlag == True)

        if (field == 'fml' or bFlag) and changed:
            order = tbl.res_order(field, changed, df)

            df.loc[order.index.tolist(), 'res_order'] = order
            # Se establece el tipo de fórmula para los res_order modificados
            ft_df = pd.concat(
                [
                    order, 
                    # values: corresponde a los datos de entrada filtrados con keys al puro comienzo
                    pd.Series(
                        values.values(), name='fml', index=pd.Index(values.keys(), name='cell')
                    ).loc[order.index]
                ],
                axis=1
            )
            ft_df = tbl.formula_type(ft_df, res_order=set(order.tolist()))
            ftype = ft_df.ftype
            df.loc[ftype.index, 'ftype'] = ftype
            
        tbl.data = df
        # Se entra a restablecer la integridad del df.
        # 1 - Enlaces externos/Parámetros: Elimina aquellos con 'dependents' vacíos.
        if circular_refs:            
            # Se establecen las dependencias para las fórmulas que crean referencias circulares.
            df = tbl.data
            cells_in_fmls = list(flatten_sets(set(x) for x in circular_refs.values()))
            cells_in_tbls = tbl.cells_in_data_rng(cells_in_fmls)
            df.loc[cells_in_tbls, 'value'] = df.loc[cells_in_tbls].value.apply(lambda x: CIRCULAR_REF.get_instance(x))
            # Se retiran los dependents de las circular_ref ya que su valor se determina en el cálculo de las circular_refs.
            changed = list(set(changed) - set(tbl.get_cells_to_calc(circular_refs.keys())))
            for code, fcells in circular_refs.items():
                df.loc[fcells, 'dependents'] = df.loc[fcells].dependents.apply(lambda x: x | {code} if isinstance(x, set) else {code})
            changed.extend(circular_refs.keys())

        tbl.changed.extend(changed)

    def pack_data(tbl):
        """
        Pack the data in the table by removing empty cells and external links without dependents.
        """
        df = tbl.data

        # Empty cells sin dependents
        mask = df.value.isin([EMPTY_CELL]) & df.dependents.isna()
        if mask.any():
            df.drop(index=df.loc[mask].index, inplace=True)

        # Enlaces externos sin dependents
        lnks = tbl.links()
        if (mask := df.loc[lnks].dependents.astype(str).str.startswith('set()')).any():
            for cell, code in df.loc[lnks].code[mask].to_dict().items():
                df.drop(index=cell, inplace=True)
                tbl.parent.reset_link(tbl, code)

        tbl.changed = list(set(tbl.changed) - set(df.loc[lnks].code[mask]))

    def clear_cells(tbl, to_clear: list[str], data:pd.DataFrame | None = None):
        """
        Clear the formulas and values of the specified cells in the table, and update the dependents 
        of downstream cells accordingly.       
        to_clear: A list of cell addresses to clear.
        data: Dataframe to clear cells. When is None the table data is used.
        """
        df = tbl.data if data is None else data
        # Se eliminan del campo "dependent" en las celdas aguas abajo 
        # las celdas que se van a modificar.
        fmls = df.loc[to_clear].fml.apply(lambda x: tbl.encoder('decode', x, df=df)).tolist()
        excel_slice = [term.replace('$', '')  for fml in fmls for term in rgn_pattern.findall(fml)]

        to_clear_coded = df.loc[to_clear].code.tolist()
        encoded_keys = tbl._cell_rgn(excel_slice)

        df.loc[encoded_keys, 'dependents'] = (
            df.loc[encoded_keys]
            .dependents
            .apply(lambda x: x.difference_update(to_clear_coded))
        )
        # Se convierten las cells de fml a values:
        df.loc[to_clear, ['fml', 'res_order', 'ftype']] = ['', 0, '#']
        return df

    @staticmethod
    def is_cell_in_fml(cell, fml):
        code = lambda cell: '{1:0>6s}|{0:_>3s}'.format(*cell_pattern.match(cell).groups()[1:]).split('|')
        if not fml:
            return False
        components = [x.replace('$', '').split(':') for x in set(rgn_pattern.findall(fml))]
        if not any(x == cell for x in itertools.chain(*components)):
            cell_code = code(cell)
            rng_limits = (
                [code(x) for x in pair]
                for pair in components if len(pair) == 2
            )
            # mask = [t[0] <= cell_code <= t[1] for t in rng_limits]
            # print(f'{cell_code=}, {mask=}')
            bflag = any(t[0][0] <= cell_code[0] <= t[1][0] and t[0][1] <= cell_code[1] <= t[1][1] for t in rng_limits)
            return bflag
        return True


    def add_empty_cells(tbl, to_init: list[str], data=None, ftype='#'):
        """
        Add empty cells to the table's data, initializing their formulas, dependents, and other fields.
        to_init: A list of cell addresses to add as empty cells.
        """
    
        value_rec = dict(fml='', dependents=set(), res_order=0, ftype=ftype, value=EMPTY_CELL, code='')

        bflag = data is None
        df = tbl.data if bflag else data
        if to_init:
            df = pd.concat(
                [
                    df,
                    pd.DataFrame([value_rec.values()], columns=value_rec.keys(), index=pd.Index(to_init, name='cell'))
                ]
            )

            # Se asignan los códigos a las nuevos registros
            df.loc[to_init, 'code'] = [f'{tbl.next_id()}' for _ in range(len(to_init))]

            # Se entra a calcular los 'dependents' siempre que existan las fórmulas y los registros a
            # agregar no sean referencias (ftype = '?')
            if ftype == '#':
                mask = ~df.fml.eq('')
                if any(mask):
                    fmls = (
                        df.loc[mask, ['fml', 'code']]
                        .assign(fml=lambda db: db.fml.apply(lambda x: tbl.encoder('decode', x if isinstance(x, str) else '' , df=df)))
                    )
                    for cell in to_init:
                        mask = fmls.fml.apply(lambda fml: tbl.is_cell_in_fml(cell, fml))
                        df.loc[[cell], 'dependents'] = [set(fmls.loc[mask].code)]
        if bflag:
            tbl.data = df
        else:
            return df

    def next_id(self):
        self.count += 1
        return f'{self.id}{self.count}'

    def normalize_data(self, df):
        df = (
            df
            .assign(code=lambda db: [self.next_id() for n in range(1, len(db.index) + 1)])
            .assign(fml= lambda db: self.encoder('encode', db['fml'], df=db))
            .assign(dependents=lambda db: db.dependents.apply(lambda x: {db.code[term] for term in x} if x is not np.nan else x))
        )
        return df

    def set_formulas(self, fmls: dict[str, str], recalc: bool = False):
        """
        Set the formulas for the specified cells in the table and recalculate if needed.
        fmls: A dictionary where keys are cell addresses and values are the corresponding formulas to set.
        recalc: A boolean indicating whether to recalculate the table after setting the formulas.
        """
        self.set_records(fmls, field='fml')
        self.recalculate(recalc)

    def cells_in_data_rng(self, cells: list[str]) -> list[str]:
        """
        Returns a list of cells that are within the data range of the table.
        cells: A list of cell addresses to check.
        return: A list of cell addresses that are within the data range of the table.

        Examples:

        wb = ExcelWorkbook('Book')
        ws = ExcelWorksheet(wb, 'sht1')
        tbl_range = 'D20:M25'
        tbl = ExcelTable(ws, 'tbl', tbl_range)

        cells = [*XlRef('E19:F51').cells()]
        tbl.cells_in_data_rng(cells) == [*XlRef('E20:F25')]
 
        cells = [*XlRef(f"'sht2'!{tbl_range}").cells()]
        tbl.cells_in_data_rng(cells) == []

        cells  = [*XlRef('E19:F51').cells(), *XlRef(f"'sht2'!{tbl_range}").cells()]
        tbl.cells_in_data_rng(cells) == [*XlRef('E20:F25')]

        (['d20']) => ['D20'] /=> ValueError()
        (['D$20']) => ['D20'] /=> ValueError()
        (['$D20']) => ['D20'] /=> []

        (['D20:Z21']) /=> ['D20:Z21']
        (['C20:E21']) /=> []


        """
        (rmin, cmin),  (rmax, cmax) = map( cell_address, self.data_rng.split(':'))
        ws_name = self.parent.title
        s = pd.Series(cells)
        s = s.where(s.str.contains('!'), f"'{ws_name}'!" + s)
        db = s.str.extract(cell_pattern, expand=True)
        mask = (db.sht == f'{ws_name}') & db.row.astype(int).between(int(rmin), int(rmax)) & db.col.between(cmin, cmax)
        if not mask.any():
            return []
        answ = pd.Series(cells, index=mask).loc[True]
        return [answ] if len(cells) == 1 else answ.tolist()
    
    def links(self):
        """
        Return any cell address that don't belong to the table range but 
        are referenced in at least a formula in the table.
        """
        cells = self.data.index.tolist()
        cells_in_rgn = self.cells_in_data_rng(cells)
        cell_links = list(set(cells) - set(cells_in_rgn))
        return cell_links

    def __del__(tbl):
        ws = tbl.parent
        wb = ws.parent
        cell_slice = tbl.data_rng
        ws.delete(cell_slice)
        ndx = ws.index(tbl)
        ws._objects.pop(ndx)

    @staticmethod
    def set_field(changes: dict[str, Any], *, field: Literal['code', 'cell']= 'code', with_dependents: bool = True, data=None):
        """
        Update the specified field ('code' or 'cell') in the provided DataFrame based on the given changes.
        changes: A dictionary where keys are the original values and values are the new values to set.
        field: The field to update ('code' for codes, 'cell' for cell addresses).
        with_dependents: A boolean indicating whether to also update dependents when updating codes.
        """
        assert isinstance(data, pd.DataFrame)
        df = data
        if field == 'code':
            if not (gmask := df.code.isin(list(changes.keys()))).any():
                return df
            direct_desc = set(df.loc[gmask].code)
            if with_dependents:
                direct_desc |= flatten_sets(df.loc[gmask].dependents)
            mask = df.code.isin(direct_desc)
            def replacement(m):
                if ':' in m[0]:
                    m = tbl_pattern.match(m[0])
                    prefix = f"'{m[1]}'!" if m[1] else ''
                    linf, lsup = map(lambda x: f"{prefix}{x}", m[0].replace(prefix, '').replace('$', '').split(':'))
                    linf = changes.get(linf, linf)
                    lsup = changes.get(lsup, lsup)
                    lsup = lsup.split('!')[-1]
                    linf, lsup = map(lambda tpl: pass_anchors(*tpl), zip(m[0].split(':'), (linf, lsup)))
                    sub_str = ':'.join([linf, lsup])
                    return sub_str
                key = m[0].replace('$', '')
                return pass_anchors(m[0], changes[key]) if key in changes else m[0]

 
            df.loc[mask, 'fml'] = df[mask].fml.str.replace(
                rgn_pattern_grp,
                replacement,
                regex=True
            )            
            # Modificación de códigos
            df.loc[gmask, 'code'] = [changes[key] for key in df.loc[gmask].code]

            if with_dependents and (mask := df.dependents.apply(lambda x: bool(x & changes.keys()))).any():
                # Modificación de dependdents
                fnc = lambda x: set(changes.get(y, y) for y in x)
                df.loc[mask, 'dependents'] = df.loc[mask].dependents.apply(fnc)
        else:  # field == 'cell'
            mask = df.code.isin(changes.keys())
            keys = df.loc[mask, 'code'].apply(changes.get)
            df.rename(index=keys, inplace=True)
        return df

    def set_values(tbl, values: dict[str, Any], recalc:bool = False):
        tbl.set_records(values, field='value')
        tbl.recalculate(recalc)

    def set_parameters(tbl, values: dict[str, Any], recalc:bool = False):
        tbl.set_records(values, field='parameter')
        tbl.recalculate(recalc)

    def get_cells_to_calc(
            tbl: 'ExcelTable', 
            codes:list[str], 
            data:'ExcelTable' | None = None, 
            isCoded:bool=True
            ) -> list[str]:
        """
        Returns a list of cells that need to be recalculated based on the provided codes.
        codes: List of codes to check for recalculation.
        data: DataFrame containing the table data. If not provided, uses the table's data
        isCoded: Boolean indicating whether the provided codes are already encoded or not.
        returns: List of cell addresses that need to be recalculated, sorted in row order when 
        codes is not coded.
        """
        ws = tbl.parent
        to_process = set()
        changed = codes
        data = data if data is not None else tbl.data
        if not isCoded:
            changed = (
                data
                .loc[changed]
                .code
                .to_list()
            )
        params = set(changed) & set(ws.parameter_code(x) for x in ws.parameters())
        changed = [code.replace(f"'{ws.id}'!", '') for code in set(changed) - params]
        # changed = tbl.cells_in_data_rng(changed)
        changed.extend(params)
        changed_cells, f_changed = tbl.cells_to_calc(changed, data=data), None
        while True:
            try:
                _, changed = changed_cells.send(f_changed)
            except StopIteration:
                break
            f_changed = list(set(changed) - to_process)
            to_process.update(f_changed)
        if not isCoded:
            mask = data["code"].isin(to_process)
            to_process = data.loc[mask].index.tolist()
        return sorted(to_process, key=lambda x: '{0: >4s}{1}'.format(*cell_address(x)))

    def ordered_formulas(self, order, feval=False):
        assert self.data is not None, 'Table not initialized'
        gmask = (self.data.code.isin(order) & (self.data.ftype != '#')).tolist()
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
                cells = cells.sort_values('cell', key=ndx_sorter)
                mask = cells.cell.tolist()
                row, col = map(set, zip(*[cell_address(cell) for cell in mask]))
                ftype, cell_range = (row.pop(), sorted(col)) if len(row) == 1 else (col.pop(), sorted(row))
                formulas.append(
                    (ftype, cell_range, cells.iloc[0].fml)
                )
            pass

        pyfmls = []
        for frst_item, scnd_item, fml in formulas:
            try:
                pyfml = self.formula_translation(frst_item, scnd_item, fml, feval=feval)
            except Exception as e:
                pyfml = ''
            pyfmls.extend(pyfml)

        return pyfmls

    def formula_translation(self, frst_item, scnd_item, fml, table_name='tbl', feval=False):
        '''
        Traduce una fórmula de Excel a una fórmula de Python
        :param fml: str. Fórmula de Excel
        :param table_name: str. Nombre de la tabla que contiene la fórmula
        :return: str. Fórmula de Python

        :Examples:
            ('12', 'D', '=SUM(A1:D20)', feval=True) => [('D12', "xlf.sum(tbl['A1:D20'].values)")]
            ('12', 'D', '=SUM(A1:D20)', feval=False) => ["D12=xlf.sum(tbl['A1:D20'].values)"]
            ('12', list('DEFG'), '=SUM(A1:D20)') => ["D12=xlf.sum(tbl['A1:D20'].values)",
                                                    "E12=xlf.sum(tbl['A1:E20'].values)",
                                                    "F12=xlf.sum(tbl['A1:F20'].values)",
                                                    "G12=xlf.sum(tbl['A1:G20'].values)"]

            ('A', ['1', '5', '20'], '=SUM(A1:D20)') => ["A1=xlf.sum(tbl['A1:D20'].values)",
                                                        "A5=xlf.sum(tbl['A5:D20'].values)",
                                                        "A20=xlf.sum(tbl['A20:D20'].values)"]

            ('A', ['1', '5', '20'], '=SUM(A11:D20)') /=> ["A1=xlf.sum(tbl['A11:D20'].values)",
                                                        "A5=xlf.sum(tbl['A55:D20'].values)",
                                                        "A20=xlf.sum(tbl['A2020:D20'].values)"]

            ('12', list('DEF'), '=SUM(A1:DA20)') /=> ["D12=xlf.sum(tbl['A1:DA20'].values)",
                                                    "E12=xlf.sum(tbl['A1:EA20'].values)",
                                                    "F12=xlf.sum(tbl['A1:FA20'].values)"]

        '''

        match (frst_item, scnd_item, fml):
            case (str(frst_item), str(scnd_item), fml):  # cell fml
                # to_pythonize = f'{scnd_item}{frst_item}{fml}'
                mask = [(f'{scnd_item}{frst_item}', scnd_item)]
                # to_pythonize = f'{scnd_item}{frst_item}{fml}'
                mask = [(f'{scnd_item}{frst_item}', scnd_item)]
                axis = None            # To avoid FutureWarning
            case (str(frst_item), list(scnd_item), fml) if frst_item.isnumeric():  # row fml
                # to_pythonize = f'{scnd_item[0]}{frst_item}{fml}'
                scnd_item = sorted(scnd_item)
                mask = [(f'{col}{frst_item}', col) for col in scnd_item]
                # to_pythonize = f'{scnd_item[0]}{frst_item}{fml}'
                scnd_item = sorted(scnd_item)
                mask = [(f'{col}{frst_item}', col) for col in scnd_item]
                axis = 0
            case (str(frst_item), list(scnd_item), fml) if frst_item.isalpha():  # column fml
                # to_pythonize = f'{frst_item}{scnd_item[0]}{fml}'
                scnd_item = sorted(scnd_item, key=lambda x: int(x))
                mask = [(f'{frst_item}{row}', row) for row in scnd_item]
                # to_pythonize = f'{frst_item}{scnd_item[0]}{fml}'
                scnd_item = sorted(scnd_item, key=lambda x: int(x))
                mask = [(f'{frst_item}{row}', row) for row in scnd_item]
                axis = 1
            case _:  # (list(frst_item), list(scnd_item), fml)    # range fml
                to_pythonize = ''
                mask = None
                axis = None
        py_fml = pythonize_fml(fml, table_name=table_name, axis=axis)
        frst_item = mask[0][1]
        if feval:
            py_fml = py_fml.lstrip('=+')
            return [(cell, py_fml.replace(frst_item, x)) for cell, x in mask]
        return [cell + py_fml.replace(frst_item, x) for cell, x in mask]

    def formula_type(self, df, res_order:list[int]|None=None):
        mask = df.res_order > 0
        if res_order is not None:
            mask = mask & (df.res_order.isin(res_order))

        order = (
            df
            .loc[mask, ['fml', 'res_order']]
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

            ngrp = 0
            for k in range(2):
                for row in sorted(rmap.keys(), key=lambda x: len(rmap[x])):
                    columns = rmap.pop(row)
                    while columns:
                        frst_col, *columns = sorted(columns, key=lambda x:(len(x), x))
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
                            ngrp += 1
                            sgrp = f'{"RC"[k]}{ngrp}'
                            mask = [frst_col] + mask
                            formulas.extend([(f'{col}{row}' if col.isalpha() else f'{row}{col}', sgrp) for col in mask])
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
        ftype = pd.Series(dict(formulas), dtype=TABLE_DATA_MAP['ftype'])
        # df.loc[:, 'ftype'] = '#'     # Se asume que toda celda por defecto es celda de valor.
        df.loc[ftype.index, 'ftype'] = ftype     # Se cambia el tipo para las celdas con fórmula.
        return df

    def cells_to_calc(
            self, 
            init_changed: list[str], 
            data: 'ExcelTable' | None=None
            ) -> Generator[tuple[int, list[str]], list[str]|None, None]:
        """
        Generador que devuelve las celdas a recalcular según res_order.

        init_changed: lista de códigos de celdas que han cambiado y que se van a recalcular.
        data: DataFrame de la tabla a recalcular. Si es None, se usa self.data.

        returns: tupla (res_order, changed) donde res_order es el orden de recalculo y changed es 
        la lista de códigos de celdas a recalcular en ese orden.
        En cada iteración es posible filtrar que celdas se van a recalcular devolviendo una lista de 
        códigos de celdas. Si se devuelve None, se recalculan todas las celdas.

        """
        init_changed = list(set(init_changed))
        df = self.data if data is None else data 
        df = (
            df
            .reset_index()
            .set_index('code')
        )
        to_report = {key: set(grp) for key, grp in df.loc[init_changed, ['res_order']].groupby(by='res_order').groups.items()}
        while True:
            try:
                k_min = min(to_report.keys())
            except ValueError:
                break
            changed = list(to_report.pop(k_min))
            changed_returned = (yield (k_min, changed))
            changed = changed if changed_returned is None else changed_returned
            if changed:
                dependents = df.loc[changed].dependents
                mask = ~dependents.isna()
                if mask.any():
                    changed = list(flatten_sets(dependents[mask]))
                    grouped_changed = {key: set(grp) for key, grp in df.loc[changed, ['res_order']].groupby(by='res_order').groups.items()}
                    [to_report.setdefault(k, set()).update(v) for k, v in grouped_changed.items()]

    def evaluate(tbl, *formulas, pythonize=True) -> list[Any] | Any:
        """
        Evaluate the provided formulas in the context of the table and return their results.
        formulas: A variable number of formulas to evaluate.
        pythonize: A boolean indicating whether to convert the formulas to Python syntax before evaluation.
        Returns: A list of evaluated results for each formula, or a single result if only one formula is provided and pythonize is True.
        """
        if pythonize:
            formulas = [
                pythonize_fml(fml, table_name='tbl', axis=None).lstrip('=+')
                for fml in formulas
            ]
        answer = []
        for py_fml in formulas:
            try:
                value = eval(py_fml, globals(), locals())
            except ZeroDivisionError as e:
                value = XlErrors.DIV_ZERO_ERROR
            except TypeError as e:
                value = XlErrors.VALUE_ERROR
            except NameError as e:
                value = XlErrors.NAME_ERROR
            except ValueError as e:
                if 'could not be broadcast' in str(e):
                    value = XlErrors.NULL_ERROR
                else:
                    value = XlErrors.NUM_ERROR
            except Exception as e:
                value = XlErrors.UNKNOWN_ERROR

            match value:
                case np.ndarray():
                    value = value.flatten()[0]
                case pd.DataFrame():
                    value = value.iat[0, 0]
                case _:
                    pass
            answer.append(value)
        return answer[0] if pythonize and len(formulas) == 1 else answer

    def recalculate(tbl, recalc:bool=False):
        """
        Recalculate the table based on the current formulas and values, propagating any changes to 
        dependent cells.
        recalc: A boolean indicating whether to perform the recalculation.
        """
        if recalc:
            ws = tbl.parent
            df = tbl.data
            # Se obtienen los códigos de las celdas que tienen un error
            error_origens = set()
            if (mask := df.index.isin([x.code for x in list(XlErrors)])).any():
                [error_origens.update(x) for x in df[mask].dependents.to_list()]
            changed, tbl.changed = tbl.changed, []
            values = {}
            changed_cells, f_changed = tbl.cells_to_calc(changed), None
            while True:
                try:
                    res_order, changed = changed_cells.send(f_changed)
                except StopIteration:
                    break
                mask = tbl.data.code.isin(changed)
                s0 = pd.Series({
                    code: value 
                    for code, value in tbl.data.loc[mask, ['code', 'value']].values
                    if code in error_origens or not isinstance(value, XlErrors)
                }, dtype=TABLE_DATA_MAP['value'])
                changed = f_changed = s0.index.tolist()
                if not f_changed:
                    continue
                if res_order == 0:
                    values.update(s0.to_dict())
                    continue
                cells, formulas = map(list, zip(*tbl.ordered_formulas(changed, feval=True)))
                vals = tbl.evaluate(*formulas, pythonize=False)
                s1 = pd.Series(vals, index=tbl.data.loc[cells].code, dtype=TABLE_DATA_MAP['value'])     # New values
                circular_ref_mask = s1 == CIRCULAR_REF
                errors_mask = s1.isin(list(XlErrors))
                mask = ~circular_ref_mask & ~errors_mask & (s0 != s1[s0.index])
                if circular_ref_mask.any():
                    circular_refs = s1[circular_ref_mask].index.tolist()
                    ws.propagate_circular_ref(tbl, circular_refs)
                if errors_mask.any():
                    # Celdas que al recalcular se generan nuevos errores
                    errors = s1[errors_mask]
                    to_propagate = {}
                    [
                        to_propagate.setdefault(y, []).append(f"'{tbl.parent.id}'!{x}")
                        for x, y in errors.items()
                    ]
                    for err, codes in to_propagate.items():
                        tbl.parent.propagate_error(err, codes=codes)
                    pass
                vals = s1[mask]
                f_changed = vals.index.tolist()
                if not f_changed:
                    continue
                # Celdas con error que al recalcular se elimina el error
                if (error_clear := vals[vals.index.isin(error_origens)].index.tolist()):
                    errors = s0[error_clear].to_dict()
                    xl_errors = {}
                    [xl_errors.setdefault(y, []).append(f"'{tbl.parent.id}'!{x}") for x, y in errors.items()]
                    for err, codes in xl_errors.items():
                        tbl.parent.propagate_error(err, codes=codes, reg_value=XlFlags.ERROR_CLEAR)
                values.update(vals.to_dict())
                vals.rename(index=lambda x: tbl.encoder('decode', x, df=tbl.data), inplace=True)
                tbl.data.loc[vals.index, 'value'] = vals

            ws = tbl.parent
            values = dict((f"'{ws.id}'!{key}", value) for key, value in values.items())
            ws.broadcast_changes(values, field='value')

    def fml_dependents(self, df):
        """
        Calculate the independent and dependent cells for each formula in the DataFrame.
        df: DataFrame containing the formulas and other relevant information.
        Returns: DataFrame with additional columns for independents and dependents.
        """
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
                    pairs.update([f'{ws_name}{alpha_code(col)}{row}'
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
                pd.DataFrame(dependents.items(), columns=['coord', 'dependents'], dtype=TABLE_DATA_MAP['dependents']).set_index('coord')
            ],
            axis=1
        )
        return df

    def res_order(tbl, field:Literal['fml', 'value', 'parameter'], changed: list[str], df) -> pd.Series:
        """
        Calculate the recalculation order for the formulas in the DataFrame based on their dependencies.
        df: DataFrame containing the formulas and their dependencies.
        Returns: DataFrame with an additional column for recalculation order.
        """

        up_codes = set(tbl.get_cells_to_calc(changed, data=df))
        # up_codes

        mask = (
            ~df.code.isin(up_codes) &   # codigos not in upcodes
            df.dependents.apply(        # codigos en los cuales aparezca un up_code en los 'dependents'
                lambda x: True if isinstance(x, set) and (x & up_codes) else False
                )
        )
        order0_map = (
            df.loc[mask, ['code', 'res_order']]
            .set_index('code')
            .res_order
            .to_dict()
        )
        down_codes = set(order0_map.keys())
        # order0_map, down_codes

        # Se construyen los pairs
        mask = df.code.isin(up_codes | down_codes)
        independents = (
            df.loc[mask, ['code', 'dependents']]
            .set_index('code')
            .dependents
            .to_dict()   
        )
        # independents

        dependency_map = collections.defaultdict(set)
        pairs = set(
            (dep, key, dependency_map[dep].add(key))[:2] for key, deps in independents.items()
            if isinstance(deps, set) and (fdeps := deps & up_codes)
            for dep in fdeps
        )
        # Se define acá que con "res_order" = 0 solo las celdas con campo "fml" = None
        # Con valor "res_order" = 1 solo celdas con 'fml' con términos independientes
        # del tipo 'res_order' = 0 o celdas con 'fml' sin términos independientes (tipo: cell = 2 + 5) 
        k_order = 1 if field == 'fml' else 0
        order0_map.update((key, k_order) for key in set(changed) - dependency_map.keys())
        # order0_map

        # pairs, dependency_map

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

        assert not (set(res_order[0]) - order0_map.keys()), 'The first batch of formulas does not contain all the formulas'
        order_map = order0_map.copy()
        # mask = []
        for batch in res_order[1:]:
            # mask.extend(batch)
            for item in batch:
                order = max(order_map[key] for key in dependency_map[item])
                order_map[item] = order + 1
        order = (
            pd.Series(order_map, name='res_order', dtype=TABLE_DATA_MAP['res_order'])
            .loc[changed]
            .rename(index=lambda x: tbl.encoder('decode', x, df=df))
        )
        # order
        return order

    def _cell_rgn(self, excel_slice: tuple[str, ...] | str) -> list[str]:
        """
        Returns a list of cell addresses that are within the specified Excel slice.
        excel_slice: A tuple or string representing the Excel slice (e.g., ('A1', 'B2') or 'A1:B2').
        Returns: A list of cell addresses within the specified range.
        """
        ref_range = f"'{self.parent.title}'!{self.data_rng}"
        return cells_in_range(ref_range, excel_slice)

    @classmethod
    def excel_table(cls, data:pd.Series, fill_value=''):
        """
        Convert a pandas Series with cell addresses as the index into a DataFrame that represents an Excel table.
        data: A pandas Series where the index contains cell addresses and the values are the corresponding data.
        fill_value: The value to use for missing cells in the resulting DataFrame.
        Returns: A pandas DataFrame representing the Excel table, with rows and columns corresponding to the cell addresses.
        """
        return (
            data
            .set_axis(
                pd.MultiIndex.from_tuples(
                    [cell_address(cell)[:2] for cell in data.index],
                    names=['row', 'col']
                )
            )
            .unstack(level=-1, fill_value=fill_value)
            .sort_index(key=lambda x: x.str.extract(r'(\d+)', expand=False).astype(int))
        )

    @classmethod
    def encoder(cls, action: Literal['decode', 'encode'], in_fmls: str | Sequence[str] | pd.Series, df: pd.DataFrame) -> pd.Series:
        """
        Encode or decode formulas in the provided DataFrame based on the specified action.
        action: A string indicating whether to 'decode' or 'encode' the formulas.
        in_fmls: The formulas to encode or decode.
        df: DataFrame containing the formulas and their dependencies.
        Returns: A pandas Series with the encoded or decoded formulas.
        """
        def translate(m, field: Literal['cell', 'code']):
            key = m[0].replace('$', '')
            var = df.loc[key, field]
            if m[2] != LINK_PREFIX + ERR_CODE_PREFIX: 
                return pass_anchors(m[0], var)
            k = int(m[3])
            return list(XlErrors)[k].value

        if b_str := isinstance(in_fmls, str):
            in_fmls = pd.Series([in_fmls])
        if action == 'decode':
            df = df.reset_index().set_index('code')
            field = 'cell'
        else:
            field = 'code'
        if (b_lst := isinstance(in_fmls, Sequence)):
            in_fmls = pd.Series(in_fmls)
        out_fmls = in_fmls.str.replace(
            cell_pattern,
            lambda m: translate(m, field),
            regex=True
        )
        if b_lst:
            out_fmls = out_fmls.tolist()
        return out_fmls.iloc[0] if b_str else out_fmls

    def get_formula(self, *excel_slice: tuple[str, ...] | None) -> pd.Series | str:
        """
        Retrieve the formulas from the specified Excel slice.
        excel_slice: A tuple or string representing the Excel slice (e.g., ('A1', 'B2') or 'A1:B2').
        Returns: A pandas Series with the formulas from the specified range.
        """
        excel_slice = excel_slice or self.data_rng
        cell_rgn = self._cell_rgn(excel_slice)
        mask = self.data.index.isin(cell_rgn) & (self.data.fml.str.len() > 0)
        coded_fmls = self.data.loc[mask, 'fml']
        fmls = self.encoder('decode', coded_fmls, df=self.data)
        mask = self.data.index.isin(cell_rgn) & (self.data.ftype == '#')
        vals = self.data.loc[mask].value
        answ = pd.concat([fmls, vals])
        if len(cell_rgn) == 1:
            return answ.iloc[0]
        return self.excel_table(answ)
    
    @staticmethod
    def offset_rng(cells: str | list[str], col_offset: int = 0, row_offset: int = 0, 
                   disc_cell: str | None = None, tbl: Optional['ExcelTable'] = None) -> str | dict[str, str]:
        """
        Calculate the offset for a range of cells in the Excel table.
        cells: A string or list of strings representing the cell addresses.
        col_offset: The number of columns to offset.
        row_offset: The number of rows to offset.
        disc_cell: The cell address to use as a reference for the offset.
        tbl: The ExcelTable instance containing the cells.
        Returns: A string or dictionary representing the offset cell addresses.
        """
        if bflag := isinstance(cells, str):
            cells = [cells]

        disc_sht = [None, tbl.parent.title] if tbl else [None,]
        predicate = lambda x: True
        # Cuando se eliminen celdas, se debe asegurar que el offset no sobrepase los límites de la tabla
        rmin, cmin = MIN_ROWS, code_alpha(MIN_COLS)
        if disc_cell:
            sht, disc_cell = tbl_address(disc_cell)
            row, col = cell_address(disc_cell)
            rmin = int(row)
            cmin = code_alpha(col)
            if sht not in disc_sht:
                disc_sht = [sht]
            if col_offset == 0 and row_offset:
                predicate = lambda x: int(cell_address(x)[0]) >= int(cell_address(disc_cell)[0])
            if col_offset and row_offset == 0:
                predicate = lambda x: code_alpha(cell_address(x)[1]) >= code_alpha(cell_address(disc_cell)[1])
            else:
                predicate = lambda x: '{0: >4s}{1}'.format(*cell_address(x)) >= '{0: >4s}{1}'.format(*cell_address(disc_cell))

        try:
            filter_rng, filter_sht, filter_cells = zip(
                *[
                    (x, *tbl_addr) for x in cells 
                    if (tbl_addr := tbl_address(x)) and tbl_addr[0] in disc_sht 
                ]
            )
        except ValueError:
            return cells[0] if bflag else {}
        ndx = pd.Index(
            [x for x in itertools.chain(*[y.split(':') for y in filter_cells]) if predicate(x)]
        )
        db = ndx.str.extract(cell_pattern, expand=True).set_index(ndx)

        mask = ~db.row.str.contains('$', regex=False)
        fnc = lambda x: str(max(rmin, int(x.strip('$')) + row_offset))
        db.loc[mask, 'row'] = db.loc[mask, 'row'].apply(fnc)
        
        mask = ~db.col.str.contains('$', regex=False)
        fnc = lambda x: alpha_code(max(cmin, code_alpha(x) + col_offset))
        db.loc[mask, 'col'] = db.loc[mask, 'col'].apply(fnc)

        db['cell'] = db.col + db.row
        cells_map = db.cell.to_dict()
        values = [':'.join(map(lambda x: cells_map.get(x, x), key.split(':'))) for key in filter_cells]
        values = [(f"'{sht}'!" if sht else '') + value for sht, value in zip(filter_sht, values)]
        answ = dict(zip(filter_rng, values))
        return answ.get(cells[0], cells[0]) if bflag else answ
    
    def direct_dependents(self, cells, data=None, is_code=False) -> pd.DataFrame:
        """
        Retrieve the direct dependents of the specified cells.
        cells: A string or list of strings representing the cell addresses.
        data: A pandas DataFrame containing the cell data.
        is_code: A boolean indicating whether the cells are specified by code.
        Returns: A pandas DataFrame with the direct dependents of the specified cells.
        """
        ws = self.parent
        wb = ws.parent
        tbl = self
        df = tbl.data if data is None else data
        # Se limita el alcance a las cells solicitadas
        df = df.loc[df.code.isin(cells)] if is_code else df.loc[cells]
        # Se da la posibilidad de que en las cells puedan venir parámetros
        codes = ['!'.join(f"'{ws.id}'!{code}".split('!')[-2:]) for code in df.code]
        fmls = []
        mask = ~df.dependents.isna()
        dep_df = df.loc[mask, ['code', 'dependents']]
        dep_pairs = [(f"'{ws.id}'!{code}", f"'{ws.id}'!{x}") for _, (code, dep) in dep_df.iterrows() for x in dep]
        fml_df = pd.DataFrame(dep_pairs, columns=['code', 'dependent'])
        fmls.append(fml_df)

        external_links = wb.links.keys() & set(f"'{ws.id}'!{x}" for x in df.code)
        etables = flatten_sets([wb.links[code] for code in external_links])
        for etbl in etables:
            sht_id, tbl_id = tbl_address(etbl)
            df = wb['#' + sht_id]['#' + tbl_id].data

            dep_df = df.loc[df.code.isin(codes), ['code', 'dependents']]
            dep_pairs = [(code, f"'{sht_id}'!{x}") for _, (code, dep) in dep_df.iterrows() for x in dep]
            fml_df = pd.DataFrame(dep_pairs, columns=['code', 'dependent'])
            fmls.append(fml_df)
        fmls = (
            pd.concat(fmls)
            .drop_duplicates()
            .set_index('code')
            .sort_index()
        )
        return fmls
    
    def external_dependents(tbl, codes, fmls=None) -> pd.DataFrame:
        """
        Retrieve the external dependents of the specified codes in the table.
        tbl: The ExcelTable instance containing the codes.
        codes: A list of codes for which to retrieve external dependents.
        fmls: A list of pandas DataFrames containing the formulas.
        Returns: A pandas DataFrame with the external dependents of the specified codes.
        """
        ws = tbl.parent
        wb = ws.parent
        fmls = fmls or [pd.DataFrame(columns=['code', 'dependent'])]
        external_links = [key for key in set(wb.links.keys()) & set(f"'{ws.id}'!{code}" for code in codes)]
        etables = flatten_sets([wb.links[code] for code in external_links])
        for etbl in etables:
            sht_id, tbl_id = tbl_address(etbl)
            df = wb['#' + sht_id]['#' + tbl_id].data

            dep_df = df.loc[df.code.isin(external_links), ['code', 'dependents']]
            dep_pairs = [(code, f"'{sht_id}'!{x}") for _, (code, dep) in dep_df.iterrows() for x in dep]
            fml_df = pd.DataFrame(dep_pairs, columns=['code', 'dependent'])
            fmls.append(fml_df)
        fmls = (
            pd.concat(fmls)
            .drop_duplicates()
            .set_index('code')
            .sort_index()
        )
        return fmls

    def __getitem__(self, excel_slice: str|list[str]) -> pd.DataFrame:
        """
        Retrieve the data for the specified Excel slice.
        excel_slice: A string or list of strings representing the cell addresses.
        Returns: A pandas DataFrame with the data for the specified slice.
        """
        cell_rgn = self._cell_rgn(excel_slice)
        # ws = self.parent
        # cell_rgn = cells_in_range(f"'{ws.title}'!{ws.data_rng}", excel_slice)
        value_cells = list(set(cell_rgn) & set(self.data.index))
        empty_cells =list(set(cell_rgn) - set(value_cells))
        data = pd.concat([
            self.data.loc[value_cells, 'value'],
            pd.Series(
                EMPTY_CELL, 
                index=pd.Index(empty_cells, name='cell'), 
                dtype=TABLE_DATA_MAP['value'],
                name='value'
            )
        ])
        df = self.excel_table(data)
        return df

    def __setitem__(self, excel_slice, value):
        """
        Set the value for the specified Excel slice.
        excel_slice: A string or list of strings representing the cell addresses.
        value: The value to set for the specified cells.
        """
        cell_rgn = list(set(self._cell_rgn(excel_slice)) & set(self.data.index))
        self.data.loc[cell_rgn, 'value'] = np.array(value).flatten()
        codes = self.encoder('encode', cell_rgn, df=self.data)
        self.changed.extend(codes)
        codes = self.encoder('encode', cell_rgn, df=self.data)
        self.changed.extend(codes)

    def __contains__(self, item) -> bool:
        """
        Check if the specified item is contained in the table.
        item: The item to check for.
        Returns: A boolean indicating whether the item is contained in the table.
        """
        row, col, sh_name = (cell_address(item) + (None, ))[:3]
        sh_name = sh_name or self.parent.title
        if (bflag := sh_name == self.parent.title):
            (rmin, cmin),  (rmax, cmax) = map( lambda x: (int((tpl:=cell_address(x))[0]), f'{tpl[1]: >2s}'), self.data_rng.split(':'))
            bflag = (rmin <= int(row) <= rmax) and (cmin <= f'{col.upper(): >2s}' <= cmax)
        return bflag

    def minimum_table(self, any_row=True) -> pd.DataFrame:
        """
        Retrieve the minimum table with the specified row filter.
        any_row: A boolean indicating whether to include rows with any non-zero values.
        Returns: A pandas DataFrame with the minimum table.
        """
        # excel_slice = self.cells_in_data_rng(self.data.index.tolist())
        # mask = (self.data.index.isin(cell_rgn)) & (self.data.value != 0)
        # excel_slice = self.data.loc[mask, :].index.tolist()
        excel_slice = self.data_rng
        df = self[excel_slice]
        if not any_row:
            mask = df.any(axis=1)
            df = df[mask]
        return df

    def _repr_html_(self):
        """
        Retrieve the HTML representation of the table.
        Returns: A string containing the HTML representation of the table.
        """
        cells_rng = self._cell_rgn(self.data_rng)
        t = pd.Series('', index=cells_rng)
        mask = self.data.index.isin(cells_rng)
        t[self.data.index[mask]] = self.data.value[mask]
        data = self.excel_table(t, fill_value='')
        return data._repr_html_()


def load_workbook(filename):
    wbxlm = excelxml.load_workbook(filename)
    excel_wb = ExcelWorkbook(filename)
    for ws_name in wbxlm.sheetnames:
        excel_wb.create_worksheet(ws_name)
    for k, ws_name in enumerate(wbxlm.sheetnames, 1):
        wsxlm = wbxlm[ws_name]
        wsheet = excel_wb[ws_name]
        ws_range = wsxlm.ws_range
        cells = wsxlm[ws_range]
        fmls = {adr: f"={fml.lstrip('=+')}" for adr, cell in cells.items() if (fml := cell.formula_)}
        values = {adr: cell.value for adr, cell in cells.items() if adr not in fmls}
        _ = ExcelTable(wsheet, f'sheet{k}_tbl', ws_range, fmls, values)
    return excel_wb

def main():
    # from tests.fixtures import dynamic_workbook
    # dynamic_workbook = dynamic_workbook.__wrapped__
    # wb = dynamic_workbook()

    fname = r'C:\Users\agmontesb\Documents\GitHub\excel\tests\files\excel_module_test.xlsx'
    excel_wb = load_workbook(fname)
    pass

if __name__ == '__main__':
    main()