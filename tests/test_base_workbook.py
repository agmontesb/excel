import sys
import os
sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pytest
import openpyxl as px
import pandas as pd
import itertools
import functools

from excel_workbook import (
    ExcelWorkbook, ExcelTable, 
    cell_address, cell_pattern, 
    data_in_range, tbl_address, rgn_pattern 
    )


class TableComparator:
    sort_fnc = staticmethod(lambda x: x.map(lambda x: '{0}{2: >4s}{1: >3s}'.format(*cell_pattern.match(x).groups())))

    def __init__(self, df):
        self.df = (
            df
            .assign(dependents=lambda db: df.dependents.apply(lambda items: str(items)))
        )

    def difference(self, other, fields=None):
        df = self.__sub__(other, fields=fields)
        return df
    
    def symmetric_difference(self, other, fields=None):
        return self.__xor__(other, fields=fields)

    def __or__(self, other, fields=None):
        df = (
            pd.concat([self.df, other.df])
            .drop_duplicates(subset=fields)
            .assign(dependents=lambda db: db.dependents.apply(lambda x: eval(x) if x != 'nan' else set()))
            .sort_index(key=self.sort_fnc)
        )
        df.loc[:, 'dependents'] = df.dependents.where(~df.dependents.isna(), set())
        return df
    
    def __and__(self, other):
        x = self ^ other
        df = self.df.loc[~self.df.index.isin(x.index)]
        df.loc[:, 'dependents'] = df.dependents.where(~df.dependents.isna(), set())
        df = df.sort_index(key=self.sort_fnc)
        return df
    
    def __sub__(self, other, fields=None):
        x = self.__xor__(other, fields=fields)
        df = x.loc[x.index.isin(self.df.index)]
        df = df.drop_duplicates(subset=['code'], keep='first')
        df.loc[:, 'dependents'] = df.dependents.where(~df.dependents.isna(), set())
        df = df.sort_index(key=self.sort_fnc)
        return df
    
    def __xor__(self, other, fields=None):
        df = (
            pd.concat([self.df, other.df])
            .drop_duplicates(subset=fields, keep=False)
            .assign(dependents=lambda db: db.dependents.apply(lambda x: eval(x) if x != 'nan' else set()))
            .sort_index(key=self.sort_fnc)
        )
        df.loc[:, 'dependents'] = df.dependents.where(~df.dependents.isna(), set())
        return df
    
    def __eq__(self, other):
        df = self ^ other
        return df.empty


@pytest.fixture
def base_workbook():
    case = 1
    # Create a base workbook for testing
    filename = r"C:\Users\agmontesb\Downloads\excel_module_test.xlsx"
    wb = px.load_workbook(filename)

    excel_wb = ExcelWorkbook('excel_module_test')

    ws_name = "No links, No parameters"
    ws = wb[ws_name]
    wsheet = excel_wb.create_worksheet(ws_name)

    # Tabla 1
    ws_range = ["G4:I9", "F3:I9"][case]
    fmls, values = data_in_range(ws, ws_range)
    sh1_tbl1 = ExcelTable(wsheet, 'sh1_tbl1', ws_range, fmls, values, recalc=True)
    # m_sh1_tbl1 = sh1_tbl1.minimun_table()

    # Tabla 2
    ws_range = ["G11:H15", "F12:H15"][case]
    fmls, values = data_in_range(ws, ws_range)
    sh1_tbl2 = ExcelTable(wsheet, 'sh1_tbl2', ws_range, fmls, values, recalc=True)
    m_sh1_tbl2 = sh1_tbl2.minimun_table()

    ws_name = "Parameters and inner links"
    ws = wb[ws_name]
    wsheet = excel_wb.create_worksheet(ws_name)

    # Tabla 1
    ws_range = ["F4:H9", "E3:H9"][case]
    fmls, values = data_in_range(ws, ws_range)
    sh2_tbl1 = ExcelTable(wsheet, 'sh2_tbl1', ws_range, fmls, values, recalc=True)
    # m_sh2_tbl1 = sh2_tbl1.minimun_table()

    # Tabla 2
    ws_range = ["F13:H17", "E12:H17"][case]
    fmls, values = data_in_range(ws, ws_range)
    sh2_tbl2 = ExcelTable(wsheet, 'sh2_tbl2', ws_range, fmls, values, recalc=True)
    # m_sh2_tbl2 = sh2_tbl2.minimun_table()

    ws_name = "Outer links, outer parameter"
    ws = wb[ws_name]
    wsheet = excel_wb.create_worksheet(ws_name)

    # Tabla 1
    ws_range = ["F3:H8", "E2:H8"][case]
    fmls, values = data_in_range(ws, ws_range)
    sh3_tbl1 = ExcelTable(wsheet, 'sh3_tbl1', ws_range, fmls, values, recalc=True)
    # m_sh3_tbl1 = sh3_tbl1.minimun_table()

    wb.close()
    return excel_wb


class TestBaseWorkbook:

    def test_base_workbook(self, base_workbook):

        assert base_workbook.title == 'excel_module_test'
        assert base_workbook.parent is None

        sheet_names = ['No links, No parameters', 'Parameters and inner links', 'Outer links, outer parameter']
        assert base_workbook.sheetnames == sheet_names

        sheet_ids = ['sheet', 'sheet1', 'sheet2']
        assert sheet_ids == [sheet.id for sheet in base_workbook.sheets]

        assert all(
            base_workbook.index(base_workbook.sheets[k]) == k 
            for k in range(len(base_workbook.sheets))
            )
        
        assert all(base_workbook[f'#{id}'].title == sheet_name for id, sheet_name in zip(sheet_ids, sheet_names))
        assert all(base_workbook[sheet_name].id == id for id, sheet_name in zip(sheet_ids, sheet_names))

    def test_base_worksheet(self, base_workbook):
        sheet = base_workbook['No links, No parameters']
        assert sheet.title == 'No links, No parameters'
        assert sheet.parent is base_workbook
        assert sheet.id == 'sheet'

        table_names = ['sh1_tbl1', 'sh1_tbl2']
        assert sheet.tablenames == table_names
        assert sheet.tables == [sheet[sheet_name] for sheet_name in table_names]

        assert all(sheet.index(sheet.tables[k]) == k for k in range(len(sheet.tables)))

        sheet_ids = [sheet.id for sheet in sheet.tables]
        assert all(sheet[f'#{id}'].title == table_name for id, table_name in zip(sheet_ids, table_names))

    def test_base_table(self, base_workbook):
        table = base_workbook['No links, No parameters']['sh1_tbl1']
        assert table.title == 'sh1_tbl1'
        assert table.parent is base_workbook['No links, No parameters']
        assert table.id == 'A'

        assert table.data_rng == 'G4:I9'
        assert table.needUpdate is False
        assert table.changed == []

        assert table.parent.index(table) == 0
        assert table.parent.objectnames() == ['sh1_tbl1', 'sh1_tbl2']

    def test_ws_parameters(self, base_workbook):
        ws = base_workbook['Parameters and inner links']
        assert ws.parameters() == ['G2']
        assert ws.parameters('G2') == [0]

        associated_tables = ws.associated_table('G2', scope='parameter')
        assert sorted(tbl.title for tbl in associated_tables) == ['sh2_tbl2', 'sh3_tbl1']

        param_code = ws.parameter_code('G2')
        dependents = [
            tbl.encoder('decode', tbl.get_cells_to_calc([param_code]))
            for tbl in associated_tables
            ]
        old_values = [
            tbl.data.loc[dependent, 'value'].tolist() for 
            tbl, dependent in zip(associated_tables, dependents)
            ]
        
        ws.parameters(G2=42420)

        new_values = [
            tbl.data.loc[dependent, 'value'].tolist() for 
            tbl, dependent in zip(associated_tables, dependents)
            ]
        
        assert all(42420 in value for value in new_values), 'Change parameter not reported'
        assert all(old != new for old, new in zip(old_values, new_values)), 'Change parameter not applied'


    def test_ws_cell_values(self, base_workbook):
        ws = base_workbook['Outer links, outer parameter']
        tbl = ws['sh3_tbl1']

        links = tbl.links()
        links.append('A1')  # Not in the table range nor in the links
        links, values = ws.cell_values(links)

        # All input cells are sheet discriminated
        assert all('!' in x for x in links)
        # For cells not in the table range or in the links, the value is None
        assert [tbl_address(cell)[-1] for cell, value in zip(links, values) if value is None] == ['A1']


class TestHelperMethods:

    def test_offset_rng(self, base_workbook):
        wb = base_workbook
        ws = wb['Parameters and inner links']
        tbl = ws['sh2_tbl1']

        # Offset single cell
        assert tbl.offset_rng('A1', row_offset=1) == 'A2', 'Offset row'
        assert tbl.offset_rng('A1', row_offset=1, col_offset=1) == 'B2', 'Offset row and col'
        assert tbl.offset_rng('$A1', row_offset=1, col_offset=1) == '$A2', 'Offset row and col'
        assert tbl.offset_rng('A$1', row_offset=1, col_offset=1) == 'B$1', 'Offset row and col'
        assert tbl.offset_rng('$A$1', row_offset=1, col_offset=1) == '$A$1', 'Offset row and col'

        # Offset range
        assert tbl.offset_rng('A1:C5', col_offset=1) == 'B1:D5', 'Offset col'

        # The disc_cell parameter: Cells be offseted if cell > = disc_cell
        # The disc_cell worksheet name if not present to be consider as tbl parent worksheet title
        assert tbl.offset_rng('D4', row_offset=1, disc_cell='A5') == 'D4', 'D4 < A5 => D4'
        assert tbl.offset_rng('B5', row_offset=1, disc_cell='A5') == 'B6', 'B5 > A5 => B6'
        # The disc_cell worksheet name offset only cells in the same worksheet
        assert tbl.offset_rng("'sheet12'!C5", row_offset=1, disc_cell="'sheet12'!A5") == "'sheet12'!C6", f'"sheet12" != "{ws.title}"'
        assert tbl.offset_rng('C5', row_offset=1, disc_cell="'sheet12'!A5") == 'C5', f'"sheet12" != "{ws.title}"'

        # Equivalentes
        r1 = tbl.offset_rng('B5', row_offset=1, disc_cell="A5")
        r2 = tbl.offset_rng('B5', row_offset=1, disc_cell=f"'{ws.title}'!A5")
        r3 = tbl.offset_rng(f"'{ws.title}'!B5", row_offset=1, disc_cell="A5")
        r4 = tbl.offset_rng(f"'{ws.title}'!B5", row_offset=1, disc_cell=f"'{ws.title}'!A5")
        assert r1 == r2 == r3.split('!')[-1] == r4.split('!')[-1], f'{r1=} != {r2=} != {r3=} != {r4=}'

        # Offset list of cells = offset each cell
        predicate = lambda x, disc_cell: '{0: >4s}{1}'.format(*cell_address(x)) >= '{0: >4s}{1}'.format(*cell_address(disc_cell))
        cells = ['A1', 'B10', 'C5', 'D4']
        kwargs = dict(row_offset=1, col_offset=1, disc_cell='A5')
        fnc = lambda x: tbl.offset_rng(x, **kwargs)
        cells_map = tbl.offset_rng(cells, **kwargs)
        assert all(fnc(key) == value for key, value in cells_map.items()), 'Offset list of cells'
        assert len(cells) >= len(cells_map), 'Offset list of cells'
        assert all(predicate(x, kwargs['disc_cell']) for x in (set(cells) - cells_map.keys()))


class TestModTables:

    def test_insert_rows(self, base_workbook):
        reduce = lambda items: functools.reduce(lambda t, e: t.union(e) or t, items, set())

        wb = base_workbook
        ws = wb['Parameters and inner links']
        tbl = ws['sh2_tbl1']

        ins_slice = '6'
        ins_rng = [f"'{ws.title}'!A6"]

        all_dep = pd.concat([
            tbl.all_dependents(cells)
            for tbl in ws.tables
            if (cells := tbl.cells_in_data_rng(tbl.data.index.tolist()))
        ])

        codes = set(all_dep.index)
        dep_parts = all_dep.dependent.apply(lambda x: cell_pattern.match(x).groups()).tolist()
        dep_df = (
            pd.DataFrame([(f"'{sht}'!{tbl}", f'{tbl}{num}') for sht, tbl, num in dep_parts], columns=['tbl', 'code'])
            .set_index('tbl')
        )

        fmls = []
        for tbl_path, codes_df in dep_df.groupby(level=0):
            sht_id, tbl_id = tbl_address(tbl_path)
            tbl = wb['#' + sht_id]['#' + tbl_id]
            codes = codes_df.code.to_list()
            fml_df = (
                tbl.data.loc[tbl.data.code.isin(codes), ['fml', 'code']]
                .set_index('code')
                .rename(index= lambda x: f"'{sht_id}'!{x}")
            )
            fmls.append(fml_df)

        fmls = pd.concat(fmls).drop_duplicates()

        fmls1 = []
        for tbl_code, fml in fmls.fml.items():
            sht_id, tbl_id, _ = cell_pattern.match(tbl_code).groups()
            ltbl = wb['#' + sht_id]['#' + tbl_id]
            fmls1.append(ltbl.encoder('decode', fml))

        # Insert a row
        ws.insert(ins_slice)

        fmls2 = []
        for tbl_code, fml in fmls.fml.items():
            sht_id, tbl_id, _ = cell_pattern.match(tbl_code).groups()
            ltbl = wb['#' + sht_id]['#' + tbl_id]
            fmls1.append(ltbl.encoder('decode', fml))


        flags = []
        for tbl_code, *fml_pair in zip(fmls.index, fmls1, fmls2):
            sht_id, tbl_id, _ = cell_pattern.match(tbl_code).groups()
            ltbl = wb['#' + sht_id]['#' + tbl_id]

            lst1, lst2 = map(rgn_pattern.findall, fml_pair)

            flags.extend(
                [ltbl.offset_rng(x, row_offset=1, disc_cell=ins_rng[0]) == y for x, y in zip(lst1, lst2) if x != y]
            )
        assert all(flags)

        pass


class TestSetRecordsTable:

    def test_field_fml(self, base_workbook):
        field = 'fml'
        # values = dict(F13='=+F15+F12', G12='=+E12&F12')   # Fórmula con referencia a celda vacía.
        # values = dict(F13='=F5 + G13', F15='=+F14+G14')   # Creación de nuevo enlace externno
        # values = dict(H13='=+$G$2*F13 + 75', F17='=+F16+F15+2*F14', G16='=F9 + F13 + F14')  # Cambiar fórmulas en celdas existentes  
        # values = dict(F12='=1+2', G12='=G14+F17', H12='=F12+G12')  # Asignar fórmulas en celdas vacías 
        # values = dict(F13='=1+2', G14='=F13+G15')  # Asignar fórmulas en celdas con valores

    def test_field_value1(self, base_workbook):
        field = 'value'
        wb = base_workbook
        ws = wb['Parameters and inner links']
        tbl = ws['sh2_tbl2']

        values = dict(F13=55, F15=80, G15=22)  # Cambiar valores en celdas existentes
        cells = list(values.keys())
        codes = [tbl.encoder('encode', x) for x in cells]

        df1 = TableComparator(tbl.data)
        tbl.set_records(values, field=field)
        tbl.recalculate(recalc=True)
        df2 = TableComparator(tbl.data)

        df = tbl.data

        flds = df.columns.tolist()
        flds.remove('value')
        diff = df2.symmetric_difference(df1, fields=flds)
        assert diff.empty, 'Esta peración solo cambia valore en el campo "value"'

        diff = df2 ^ df1
        assert set(diff.index.unique()) & set(cells) == set(cells), 'Al menos las celdas reportadas cambian'

        all_dependents = set(tbl.get_cells_to_calc(codes))
        all_codes = set(diff.code.to_list())
        assert all_codes.issubset(all_dependents), 'Faltan las "cell" que no cambian de valor porque por ejemplo en la fórmula se multiplica por cero'

        pass

    def test_field_value2(self, base_workbook):
        field = 'value'
        wb = base_workbook
        ws = wb['Parameters and inner links']
        tbl = ws['sh2_tbl2']

        values = dict(F12=35, G12=80, H12=22)  # Asignar valores en celdas vacías 

        df1 = TableComparator(tbl.data)
        tbl.set_records(values, field=field)
        tbl.recalculate(recalc=True)
        df2 = TableComparator(tbl.data)

        cells = list(values.keys())
        codes = [tbl.encoder('encode', x) for x in cells]

        
        # values = dict(G16=333, H13=222)  # Asignar valores en celdas con fórmulas
