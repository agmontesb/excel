import sys
import os
sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from pathlib import Path
import pytest
import pandas as pd

from utilities import tbl_data, wb_from_excelfile, TableComparator
from excel_workbook import ExcelWorkbook, ExcelTable, EMPTY_CELL

wb_structure = [
    # (ws_name, tbl_name, tbl_range)
    ('No links, No parameters', 'sht1_tbl1', 'F3:I9'),
    ('No links, No parameters', 'sht1_tbl2', 'F12:H15'),
    ('Parameters and inner links', 'sht2_tbl1', 'E3:H9'),
    ('Parameters and inner links', 'sht2_tbl2', 'E12:H17'),
    ('Outer links, outer parameter', 'sht3_tbl1', 'E3:H8'),
]

@pytest.fixture(scope='module')
def empty_workbook():
    excel_wb = ExcelWorkbook('excel_module_test')
    return excel_wb

@pytest.fixture()
def static_workbook():
    excel_wb = ExcelWorkbook('excel_module_test')

    # Create a base workbook for testing
    for ws_name, tbl_name, tbl_range in wb_structure:
        if not ws_name in excel_wb.sheetnames:
            wsheet = excel_wb.create_worksheet(ws_name)
        fmls, values, tblv = tbl_data(tbl_name, tbl_range)
        sht_tbl = ExcelTable(wsheet, tbl_name, tbl_range, fmls, values)

        # Se verifica que se reconstuyen los valores de la tabla
        df = sht_tbl.data
        assert (df.loc[tblv.index].value.map(str) == tblv).all()
    return excel_wb


@pytest.fixture
def dynamic_workbook():
    # Create a base workbook for testing
    filename = Path(__file__).parent / 'files' / 'excel_module_test.xlsx'
    return wb_from_excelfile(filename, wb_structure)


def test_Form2517():
    form2517_structure = [
            ('H2 (ESF - Patrimonio)', 'esf_tbl', 'G9:K193'), # Estado de Situación Financiera
            ('H3 (ERI - Renta Liquida)', 'pat_tbl', 'G196:K220'), # Patrimonio
            ('H7 (Resumen ESF-ERI)', 'rlc_tbl', 'H9:L375'), # Renta líquida cedular
            ('H8 (Resumen ESF-ERI)', 'edr_tbl', 'H376:N457'), # Estado de Resultados
    ]
    filename = Path(__file__).parent / 'files' / 'Reporte_Conciliación_Fiscal_F2517V6_AG2023_v1.0.1-2024.xlsm'
    return wb_from_excelfile(filename, form2517_structure)


@pytest.fixture
def base_tables():
    value_rec = value_rec = dict(fml='', dependents=set(), res_order=0, ftype='#', value=EMPTY_CELL, code='')
    
    df1 = (
        pd.DataFrame([value_rec.values()], columns=value_rec.keys(), index=pd.Index(['A1', 'A2', 'A3'], name='cell'))
        .assign(code=lambda db: db.index.str.replace('A', 'M'))
        .assign(dependents=lambda db: [set(), set(), set(['M1','M2'])])
        .sort_index()    
    )

    df2 = (
        pd.concat(
            [
                df1,
                pd.DataFrame([value_rec.values()], columns=value_rec.keys(), index=pd.Index(['B5'], name='cell'))
            ]
        )
        .sort_index()
    )
    df2.loc['A1', 'value'] = 756
    df2.loc['B5', 'code'] = 'M32'

    tdf1, tdf2 = map(TableComparator, (df1, df2))
    yield tdf1, tdf2



