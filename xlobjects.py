from __future__ import annotations

from typing import Dict, Iterator, List, TypedDict
from enum import Enum, Flag
import re

from xlpatterns import (
    MIN_COLS, MIN_ROWS, 
    MAX_COLS, MAX_ROWS,
    alpha_code, code_alpha,
    cell_pattern, cell_address,
    tbl_pattern, tbl_address,
)


__all__ = [
    'alpha_code', 
    'TABLE_DATA_MAP', 'ColumnSpec', 'DataFrameStructure', 'structure', 
    'XlFlags', 'Inmutable', 'XlErrors', 'CircularRef', 'EmptyCell',
    'CIRCULAR_REF', 'EMPTY_CELL'
]


TABLE_DATA_MAP = {
    'fml': str, 'dependents': object, 'res_order': int, 
    'ftype': str, 'value': object, 'code': str
}

class ColumnSpec(TypedDict):
    column_name: type

# Define a TypedDict for the structure of the DataFrame
class DataFrameStructure(TypedDict):
    index: List[type]  # List of types for the index
    columns: Dict[str, ColumnSpec]  # Dictionary of column names and their types

# Example structure definition
structure: DataFrameStructure = {
    'index': [str],  # Example: Index should be integers
    'columns': TABLE_DATA_MAP
}


class XlFlags(Flag):
    ERROR_CLEAR = 1
    EMPTY_CELL = 2
    VALUE_CELL = 4


class Inmutable:
    def __add__(self, other):
        return self

    def __radd__(self, other):
        return self

    def __sub__(self, other):
        return self

    def __rsub__(self, other):
        return self

    def __mul__(self, other):
        return self

    def __rmul__(self, other):
        return self

    def __truediv__(self, other):
        return self

    def __rtruediv__(self, other):
        return self

    def __eq__(self, other: object) -> bool:
        return True if isinstance(other, self.__class__) else False

    def __ne__(self, other: object) -> bool:
        return not self.__eq__(other)

    def __hash__(self):
        return super().__hash__()


class XlErrors(Inmutable, Enum):
    REF_ERROR = "#REF!"
    VALUE_ERROR = "#VALUE!"
    DIV_ZERO_ERROR = "#DIV/0!"
    NAME_ERROR = "#NAME?"
    NUM_ERROR = "#NUM!"
    NULL_ERROR = "#NULL!"
    GETTING_DATA_ERROR = "#GETTING_DATA!"
    SPILL_ERROR = "#SPILL!"
    UNKNOWN_ERROR = "#UNKNOWN!"
    # Add more Excel error types here as needed

    def __str__(self):
        return self.value

    @property
    def code(self):
        id = list(self.__class__._value2member_map_.keys()).index(self.value)
        return f'Z{id}'
  
    def __hash__(self):
        return super().__hash__()
    

class CircularRef(Inmutable):
    value = 0

    @classmethod
    def get_instance(cls, value=None):
        self = cls()
        if isinstance(value, cls):
            return value
        self.value = value if value is not None or isinstance(value, EmptyCell) else 0 
        # self.value = value or 0
        return self
    
    def __str__(self):
        return str(self.value)
    

class EmptyCell:
    def __init__(self):
        self.value = 0

    def __add__(self, other):
        if isinstance(other, str):
            return other
        else:
            return self.value + other

    def __radd__(self, other):
        return self.__add__(other)  # Delegate to __add__

    def __sub__(self, other):
        return self.value - other

    def __rsub__(self, other):
        return -self.__sub__(other)

    def __mul__(self, other):
        return self.value * other

    def __rmul__(self, other):
        return self.__mul__(other)

    def __truediv__(self, other):
        return self.value / other

    def __rtruediv__(self, other):
        return self.__truediv__(other)

    def __eq__(self, value):
        if (value is None) or isinstance(value, EmptyCell):
            return True
        return False

    def __hash__(self):
        return id(self)

    def __bool__(self):
        return False

    def __str__(self):
        return ""
    

EMPTY_CELL = EmptyCell()
CIRCULAR_REF = CircularRef()


class XlRef:

    def __init__(self, xlref: str):
        sheet, xlslice = tbl_address(xlref)
        self.sheet = sheet
        self.adr = xlslice = xlslice.upper()

        xlslice = xlslice.replace('$', '')
        tl_ref, br_ref = f'{xlslice}:{xlslice}'.split(':', 2)[:2]

        if all(x.isnumeric() for x in (tl_ref, br_ref)):  # row(s): '5:20', '10'
            tl_row, br_row = map(int, (tl_ref, br_ref))
            tl_col, br_col = MIN_COLS, MAX_COLS
        elif all(x.isalpha() for x in (tl_ref, br_ref)):    # column(s): 'A:C', 'Z'
            tl_col, br_col = map(code_alpha, (tl_ref, br_ref))
            tl_row, br_row = MIN_ROWS, MAX_ROWS
        else:   # general: 'A5', 'D1:Z20'
            (tl_row, tl_col),  (br_row, br_col) = map(cell_address, (tl_ref, br_ref))
            tl_row, br_row = map(int, (tl_row, br_row))
            tl_col, br_col = map(code_alpha, (tl_col, br_col))
        self.coords = (tl_row, tl_col, br_row, br_col)

    def isCell(self) -> bool:
        """
        Determines if the XlRef object represents a single cell reference.

        ej:
        XlRef('B5').isCell() == True
        XlRef('B5:B5').isCell() == False

        """
        adr = self.adr.replace('$', '')
        return ':' not in adr and not adr.isalpha() and not adr.isnumeric()

    def cells(self, column_wise: bool = True) -> Iterator[str]:
        """
        Generates cell references within the range defined by the XlRef object.

        Args:
            column_wise (bool): If True, generates cells column-wise; if False, generates cells row-wise.

        Yields:
            str: Cell reference in the format 'A1', 'B2', etc.

        ej:
        list(XlRef('B2:C4').cells()) == ['B2', 'C2', 'B3', 'C3']
        list(XlRef('B2:C4').cells(column_wise=False)) == ['B2', 'B3', 'C2', 'C3']

        """
        tl_row, tl_col, br_row, br_col = self.coords
        prefix = f"'{self.sheet}'!" if self.sheet else ''
        if column_wise:
            for r in range(tl_row, br_row + 1):
                for c in range(tl_col, br_col + 1):
                    yield f"{prefix}{alpha_code(c)}{r}"
        else:
            for c in range(tl_col, br_col + 1):
                for r in range(tl_row, br_row + 1):
                    yield f"{prefix}{alpha_code(c)}{r}"

    def intersects(self, other: 'XlRef') -> 'XlRef | None':
        """
        Determines the intersection of this XlRef object with another XlRef object.

        Args:
            other (XlRef): Another XlRef object to check for intersection.

        Returns:
            XlRef | None: A new XlRef object representing the intersection range, or None if there is no intersection.

        ej:
        XlRef('A1:D5').intersects(XlRef('C3:F10')) == XlRef('C3:D5')
        XlRef('A1:D5').intersects(XlRef('F10')) == None

        """
        if not isinstance(other, XlRef) or self.sheet != other.sheet:
            return None

        tl_row1, tl_col1, br_row1, br_col1 = self.coords
        tl_row2, tl_col2, br_row2, br_col2 = other.coords

        # Calculate the intersection coordinates
        inter_tl_row = max(tl_row1, tl_row2)
        inter_tl_col = max(tl_col1, tl_col2)
        inter_br_row = min(br_row1, br_row2)
        inter_br_col = min(br_col1, br_col2)

        # Check if there is an intersection
        if inter_tl_row <= inter_br_row and inter_tl_col <= inter_br_col:
            xlslice = f"{alpha_code(inter_tl_col)}{inter_tl_row}:{alpha_code(inter_br_col)}{inter_br_row}"
            xlref = f"'{self.sheet}'!{xlslice}" if self.sheet else xlslice
            return XlRef(xlref)
        
        return None

    def offset(self, row_offset: int = 0, col_offset: int = 0) -> 'XlRef' | XlErrors.REF_ERROR:
        """
        Returns a new XlRef object that is offset by the specified number of rows and columns.

        Args:
            row_offset (int): Number of rows to offset. Positive values move down, negative values move up.
            col_offset (int): Number of columns to offset. Positive values move right, negative values move left.

        Returns:
            XlRef: A new XlRef object representing the offset range.

        ej:
            XlRef('A3').offset(1, 1) == XlRef('B4')
            XlRef('$A3').offset(1, 1) == XlRef('$A4')
            XlRef('$A$3').offset(1, 1) == XlRef('$A$3')
            XlRef('$A$3').offset(1, 1) == XlRef('$A$3')

            XlRef('B4:$C$6').offset(2, 1) == XlRef('C6:$C$6')
            XlRef('B4:$C$6').offset(3, 0) == XlRef('B$6:$C7')
            XlRef('B4:$C$6').offset(0, 3) == XlRef('$C4:E$6')
            XlRef('B4:$C$6').offset(3, 3) == XlRef('$C$6:E7')

            XlRef('A3').offset(-1, 0) == XlErrors.REF_ERROR
            XlRef('A3').offset(0, -1) == XlErrors.REF_ERROR
            XlRef('A3').offset(-1, -1) == XlErrors.REF_ERROR

        """
        tl_cell, br_cell = self.xlslice_corners()
        coords = list(self.coords)
        answ = []
        cells = (tl_cell, ) if self.isCell() else (tl_cell, br_cell)
        for k, cell in enumerate(cells):
            row, col = cell_address(cell)
            if row.isnumeric():
                if not MIN_ROWS <= (coord := coords[k * 2] + row_offset) <= MAX_ROWS:
                    return XlErrors.REF_ERROR
                coords[k * 2] = coord
                row = str(coords[k * 2])
            if col.isalpha():
                if not MIN_COLS <= (coord := coords[k * 2 + 1] + col_offset) <= MAX_COLS:
                    return XlErrors.REF_ERROR
                coords[k * 2 + 1] = coord
                col = alpha_code(coords[k * 2 + 1])
            answ.extend(((coords[k * 2], row), (coords[k * 2 + 1], col)))
            
        rows = sorted(answ[::2], key=lambda x:x[0])
        cols = sorted(answ[1::2], key=lambda x:x[0])
        xlslice = ':'.join(f'{col[1]}{row[1]}' for row, col in zip(rows, cols))
        xlref = f"'{self.sheet}'!{xlslice}" if self.sheet else xlslice
        return XlRef(xlref)

    def xlslice_corners(self):
        return f'{self.adr}:{self.adr}'.split(':', 2)[:2]

    def r1c1_xlref(self, xlref: str | 'XlRef' | None = None) -> str | None:
        """
        Returns a str object that represents the xlref in R1C1 format.

        Args:
            xlref (XlRef | str): xlref to be converted.

        Returns:
            str: str object representing the XlRef in absolute (xlref=None) or relative (xlref != None)

        ej:

        XlRef('B5').r1c1_xlref() == 'R5C2'
        XlRef('$B$5').r1c1_xlref() == 'R5C2'
        
        XlRef('B5').r1c1_xlref(XlRef('A1')) == 'R[-4]C[-1]'
        XlRef('B5').r1c1_xlref(XlRef('A$1')) == 'R1C[-1]'
        XlRef('B5').r1c1_xlref(XlRef('$A1')) == 'R[-4]C1'
        XlRef('B5').r1c1_xlref(XlRef('$A$1')) == 'R1C1'

        XlRef('B5').r1c1_xlref('A5') == 'RC[-1]'
        XlRef('B5').r1c1_xlref('$A$5') == 'R5C1'

        """

        if xlref is None:
            prefix = f"'{self.sheet}'!" if self.sheet else ''
            pairs = list(zip(self.coords[::2], self.coords[1::2]))
            if self.isCell(): pairs = pairs[:1]
        else:
            if isinstance(xlref, str):
                try:
                    xlref = XlRef(xlref)
                except Exception:
                    return None
            elif not isinstance(xlref, XlRef):
                return None

            prefix = f"'{xlref.sheet}'!" if xlref.sheet else ''
            xlref_coords = xlref.coords
            tl_cell, br_cell = xlref.xlslice_corners()
            cells = (tl_cell, ) if xlref.isCell() else (tl_cell, br_cell)

            coords = 2*self.coords[:2]
            pairs = []
            for k, cell in enumerate(cells):
                row, col = cell_address(cell)
                row = xlref_coords[2 * k] if '$' in row else f'[{xlref_coords[2 * k] - coords[2 * k]}]'
                col = xlref_coords[2 * k + 1] if '$' in col else f'[{xlref_coords[2 * k + 1] - coords[2 * k + 1]}]'
                pairs.append((row, col))
        
        fnc = lambda x: 'R{}C{}'.format(*x)
        xlslice = ':'.join(map(fnc, pairs)).replace('[0]', '')
        return f'{prefix}{xlslice}'

    def a1_xlref(self, r1c1xlref:str):
        """
        Returns a str object that represents the r1c1xlref in A1 format.

        Args:
            r1c1xlref (str): xlref to be converted.

        Returns:
            str: str object representing the r1c1xlref in A1 format.

        ej:

        XlRef('B5').a1_xlref('R[-4]C[-1]') == 'A1' 
        XlRef('B5').a1_xlref('R1C[-1]') == 'A$1'
        XlRef('B5').a1_xlref('R[-4]C1') == '$A1'
        XlRef('B5').a1_xlref('R1C1') == '$A$1'

        XlRef('B5').a1_xlref('RC[-1]') == 'A5'
        XlRef('B5').a1_xlref('R5C1') == '$A$5'

        """
        sheet, r1c1 = f'!{r1c1xlref}'.split('!')[-2:]
        prefix = f"{sheet}!" if sheet else ''
        deltas = re.findall(r'[RC]([0-9\]\-\[]*)', r1c1)
        coords = self.coords[:2]
        pairs = []
        for k in (0, 1):
            fnc = [str, alpha_code][k]
            for delta in deltas[k::2]:
                delta = delta or '[0]'
                pfix = '' if delta[0] == '[' else '$'
                dmy = int(delta) if pfix else (coords[k] + int(delta[1:-1]))
                if dmy < 0: return XlErrors.REF_ERROR
                pairs.append((dmy, pfix + fnc(dmy)))
        # Se ordenan los pares para tener tl_cell < rb_cell
        k = len(deltas) // 2
        rows, cols = pairs[:k], pairs[k:]
        if k == 2:
            rows, cols = map(lambda x: sorted(x, key=lambda y: y[0]), (rows, cols))
        xlslice = ':'.join(f'{col[1]}{row[1]}' for row, col in zip(rows, cols))
        return f'{prefix}{xlslice}'

    def __contains__(self, item: str | 'XlRef') -> bool:
        """
        Checks if a given cell reference or XlRef object is within the range defined by this XlRef object.

        Args:
            item (str | XlRef): The cell reference (e.g., 'A1') or another XlRef object to check.

        Returns:
            bool: True if the item is within the range; False otherwise.

        ej:
            XlRef('D5') in 'A1:F10' == True
            XlRef('D5') in 'A1:C10' == False
            XlRef('D5:E6') in 'A1:D5' == False

        """
        if isinstance(item, str):
            item = XlRef(item)
        if not isinstance(item, XlRef):
            return False

        xlref = self.intersects(item)
        return xlref is not None and xlref.coords == item.coords  # Ensure both are XlRef objects

    def __str__(self):
        return f"'{self.sheet}'!{self.adr}" if self.sheet else self.adr

    def __repr__(self):
        xlslice = str(self)
        return f"XlRef('{xlslice}')"

    def __eq__(self, other: object) -> bool:
        """
        Checks if two XlRef objects are equal based on their sheet and address.
        """
        if isinstance(other, str):
            return str(self) == other
        if not isinstance(other, XlRef):
            return False
        return self.sheet == other.sheet and self.adr == other.adr


def main():

    xlrefstr = ['F', 'Z:AD', '5', '5:20', 'd50', "'sheet1'!a10:d20", 'a10:D20', 'A10:D20']
    for refstr in xlrefstr:
        print(f'testing {refstr}')
        sheet, xlslice = tbl_address(refstr)
        prefix = f"'{sheet}'!" if sheet else ''
        anchors = f'{xlslice.upper()}:{xlslice.upper()}'.split(':', 2)[:2]
        if all(x.isnumeric() for x in anchors):
            rows = anchors
            cols = map(alpha_code, (MIN_COLS, MAX_COLS))
        elif all(x.isalpha() for x in anchors):
            rows = map(str, (MIN_ROWS, MAX_ROWS))
            cols = anchors
        else:
            rows, cols = zip(*map(cell_address, anchors))


        xlref = XlRef(refstr)
        assert str(xlref) == f'{prefix}{xlslice.upper()}'
        assert xlref.sheet == sheet
        assert xlref.adr == xlslice.upper()
        assert xlref.coords[::2] == tuple(map(int, rows))
        assert xlref.coords[1::2] == tuple(map(code_alpha, cols))


    assert XlRef('A1').isCell()
    xlrefs = ['A1:B10', 'A1:A1', 'A:C', '1:10', 'C', '10']
    assert all(not XlRef(x).isCell() for x in xlrefs)

    
    assert list(XlRef('A1:B3').cells()) == ['A1', 'B1', 'A2', 'B2', 'A3', 'B3']
    assert list(XlRef('A1:B3').cells(column_wise=False)) == ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
    assert list(XlRef('A1:B5').cells()) == [f'{col}{row}' for row in range(1, 6) for col in 'AB']
    assert list(XlRef('A1:B5').cells(column_wise=False)) == [f'{col}{row}' for col in 'AB' for row in range(1, 6)]

    pairs = [('A1:C10', 'B5:C10'), ('C10:C20', 'C10:C15'),
             ('A3:B5', 'B5:B5'), ('D4:E5', 'D5:D5'), ('D15:E16', 'D15:D15'), ('A15:B16', 'B15:B15'),
             ('C8:D9', 'C8:D9'),
             ('A1', None), ('E16', None)
             ]
    xlref_base = XlRef('B5:D15')
    for xlref, answ in pairs:
        assert str(xlref_base.intersects(XlRef(xlref))) == str(answ), f"Failed for {xlref}: expected {answ}, got {xlref_base.intersects(XlRef(xlref))}"

    for xlref, answ in pairs:
        assert xlref != answ or (XlRef(xlref) in xlref_base), f"Failed for {xlref}: expected (XlRef({xlref}) in {xlref_base}) == True"

    pairs = [
        ('A1:C5', 3, 5, 'F4:H8'), ('$A1:C5', 3, 5, '$A4:H8'), ('$A$1:C5', 3, 5, '$A1:H8'), 
        ('A1:$C$5', 3, 5, '$C4:F$5'),
    ]

    xlref_off = """A3:$C$6	B3:$C$6	C3:$C$6	$C3:D$6	$C3:E$6
    A4:$C$6	B4:$C$6	C4:$C$6	$C4:D$6	$C4:E$6
    A5:$C$6	B5:$C$6	C5:$C$6	$C5:D$6	$C5:E$6
    A6:$C$6	B6:$C$6	C6:$C$6	$C6:D$6	$C6:E$6
    A$6:$C7	B$6:$C7	C$6:$C7	$C$6:D7	$C$6:E7"""

    xlref_offsets = [x.strip().split('\t') for x in xlref_off.split('\n')]
    xlref_base = XlRef('B4:$C$6')
    for row_off in range(-1, 4):
        for col_off in range(-1, 4):
            print(xlref_base, row_off, col_off, xlref_offsets[row_off + 1][col_off + 1])
            assert f'{xlref_base.offset(row_off, col_off)}' == xlref_offsets[row_off + 1][col_off + 1]

    assert xlref_base.offset(-MAX_ROWS, 0) == XlErrors.REF_ERROR
    assert xlref_base.offset(MAX_ROWS, 0) == XlErrors.REF_ERROR
    assert xlref_base.offset(0, -MAX_COLS) == XlErrors.REF_ERROR
    assert xlref_base.offset(0, MAX_COLS) == XlErrors.REF_ERROR

    dmy = xlref_base.r1c1_ref()

    pass




if __name__ == '__main__':
    main()