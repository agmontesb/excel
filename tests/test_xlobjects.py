import pytest
import sys
import os
import re

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from xlobjects import (
    XlRef, XlErrors,
    tbl_address, alpha_code, MIN_ROWS, MAX_ROWS, MIN_COLS, MAX_COLS, cell_address, code_alpha,
)


class Test_XlRef:

    @pytest.mark.parametrize("refstr", [
        'F', 'Z:AD', '5', '5:20', 'd50', "'sheet1'!a10:d20", 'a10:D20', 'A10:D20'
    ])
    def test_xlref_attribs(self, refstr):
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

    @pytest.mark.parametrize("xlrefstr, answer", [
        ('A1', True),
        ('A1:B10', False), ('A1:A1', False), ('A:C', False), ('1:10', False), 
        ('C', False), ('10', False),
    ])
    def test_iscell(self, xlrefstr, answer):
        assert XlRef(xlrefstr).isCell() == answer

    @pytest.mark.parametrize("xlrefstr, column_wise, answer", [
        ('A1:B3', True, ['A1', 'B1', 'A2', 'B2', 'A3', 'B3']),
        ('A1:B3', False, ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']),
        ('A1:B5', True, [f'{col}{row}' for row in range(1, 6) for col in 'AB']),
        ('A1:B5', False, [f'{col}{row}' for col in 'AB' for row in range(1, 6)]),
    ])
    def test_cells(self, xlrefstr, column_wise, answer):
        assert list(XlRef(xlrefstr).cells(column_wise)) == answer

    @pytest.mark.parametrize("xlrefstr_base, xlrefstr, answer", [
        ('B5:D15', 'A1:C10', 'B5:C10'), 
        ('B5:D15', 'C10:C20', 'C10:C15'),
        ('B5:D15', 'A3:B5', 'B5:B5'), 
        ('B5:D15', 'D4:E5', 'D5:D5'), 
        ('B5:D15', 'D15:E16', 'D15:D15'), 
        ('B5:D15', 'A15:B16', 'B15:B15'),
        ('B5:D15', 'C8:D9', 'C8:D9'),
        ('B5:D15', 'A1', None), 
        ('B5:D15', 'E16', None),        
    ])
    def test_intersects(self, xlrefstr_base, xlrefstr, answer):
        xlref_base = XlRef(xlrefstr_base)
        assert xlref_base.intersects(XlRef(xlrefstr)) == answer, f"Failed for {xlrefstr}: expected {answer}, got {xlref_base.intersects(XlRef(xlrefstr))}"

    @pytest.mark.parametrize("xlrefstr_base, xlrefstr, answer", [
        ('B5:D15', 'A1:C10', False), 
        ('B5:D15', 'C10:C20', False),
        ('B5:D15', 'A3:B5', False), 
        ('B5:D15', 'D4:E5', False), 
        ('B5:D15', 'D15:E16', False), 
        ('B5:D15', 'A15:B16', False),
        ('B5:D15', 'C8:D9', True),
        ('B5:D15', 'A1', False), 
        ('B5:D15', 'E16', False),        
    ])
    def test_contains(self, xlrefstr_base, xlrefstr, answer):
        xlref_base = XlRef(xlrefstr_base)
        assert (XlRef(xlrefstr) in xlref_base) == answer, f"Failed for {xlrefstr}: expected (XlRef({xlrefstr}) in {xlref_base}) == True"

    @pytest.mark.parametrize("row_off, col_off, answer", [
        (-1, -1, 'A3:$C$6'),
        (-1, 0, 'B3:$C$6'),
        (-1, 1, 'C3:$C$6'),
        (-1, 2, '$C3:D$6'),
        (-1, 3, '$C3:E$6'),
        (0, -1, 'A4:$C$6'),
        (0, 0, 'B4:$C$6'),
        (0, 1, 'C4:$C$6'),
        (0, 2, '$C4:D$6'),
        (0, 3, '$C4:E$6'),
        (1, -1, 'A5:$C$6'),
        (1, 0, 'B5:$C$6'),
        (1, 1, 'C5:$C$6'),
        (1, 2, '$C5:D$6'),
        (1, 3, '$C5:E$6'),
        (2, -1, 'A6:$C$6'),
        (2, 0, 'B6:$C$6'),
        (2, 1, 'C6:$C$6'),
        (2, 2, '$C6:D$6'),
        (2, 3, '$C6:E$6'),
        (3, -1, 'A$6:$C7'),
        (3, 0, 'B$6:$C7'),
        (3, 1, 'C$6:$C7'),
        (3, 2, '$C$6:D7'),
        (3, 3, '$C$6:E7'),
        (-MAX_ROWS, 0, XlErrors.REF_ERROR),
        (MAX_ROWS, 0, XlErrors.REF_ERROR),
        (0, -MAX_COLS, XlErrors.REF_ERROR),
        (0, MAX_COLS, XlErrors.REF_ERROR),
    ])
    def test_offset(self, row_off, col_off, answer):
        xlref_base = XlRef('B4:$C$6')
        assert xlref_base.offset(row_off, col_off) == answer

    @pytest.mark.parametrize("xlrefstr, xlref_base, answer", [
        ('B5', None, 'R5C2'),
        ('$B$5', None, 'R5C2'),
        ('B5', XlRef('A1'), 'R[4]C[1]'),
        ('$B$5', XlRef('A1'), 'R5C2'),
        ('B$5', XlRef('A1'), 'R5C[1]'),
        ('$B5', XlRef('A1'), 'R[4]C2'),
        ('B5', XlRef('A5'), 'RC[1]'),
        ('A15', XlRef('A5'), 'R[10]C'),
        ('$B$5', XlRef('A5'), 'R5C2'),
    ])
    def test_r1c1_ref(self, xlrefstr, xlref_base, answer):
        assert XlRef(xlrefstr).r1c1_ref(xlref_base) == answer

    @pytest.mark.parametrize("xlref_base, xlref, answer", [
        ('B5', None, 'R5C2'),
        ('$B$5', None, 'R5C2'),
        ('B5', 'A1', 'R[-4]C[-1]'),
        ('B5', '$A1', 'R[-4]C1'),
        ('B5', 'A$1', 'R1C[-1]'),
        ('B5', '$A$1', 'R1C1'),
        ('B5', 'A5', 'RC[-1]'),
        ('B5', '$A$5', 'R5C1'),
    ])
    def test_r1c1_xlref(self, xlref_base, xlref, answer):
        assert XlRef(xlref_base).r1c1_xlref(xlref) == answer

    @pytest.mark.parametrize("xlref_base, answer, r1c1_xlref", [
        ('B5', 'A1', 'R[-4]C[-1]'),
        ('B5', '$A1', 'R[-4]C1'),
        ('B5', 'A$1', 'R1C[-1]'),
        ('B5', '$A$1', 'R1C1'),
        ('B5', 'A5', 'RC[-1]'),
        ('B5', '$A$5', 'R5C1'),
    ])
    def test_a1_xlref(self, xlref_base, answer, r1c1_xlref):
        assert XlRef(xlref_base).a1_xlref(r1c1_xlref) == answer