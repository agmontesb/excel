import sys
import os
sys.path.append(os.path.abspath(os.path.dirname(__file__)))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import pandas as pd

from utilities import TableComparator
from fixtures import base_tables, static_workbook, empty_workbook


class TestTableComparator:

    def test_base_codes(self, static_workbook, empty_workbook):
        base_tbl = static_workbook.sheets[0].tables[0]
        tbl1 = empty_workbook.create_worksheet('sheet1').create_table(base_tbl.title, base_tbl.data_rng)

        df1 = base_tbl.data
        tdf1 = TableComparator(df1)

        values = df1.value
        fmls = df1.loc[df1.fml != ''].fml.apply(lambda x: base_tbl.encoder('decode', x, df=base_tbl.data))
        value_cells = list(set(values.index) - set(fmls.index))
        values = values.loc[value_cells].to_dict()
        fmls = fmls.to_dict()

        tbl1.set_records(fmls, field='fml')
        tbl1.set_records(values, field='value')
        tbl1.recalculate(recalc=True)
        tdf2 = TableComparator(tbl1.data)

        tbl1.set_records({'I3': 10, 'G3': 20, 'H3': 30}, field='value')
        tbl1.recalculate(recalc=True)

        df2 = tbl1.data
        tdf2 = TableComparator(df2)

        dmy = tdf1 ^ tdf2

        pass

    def test_union_operator(self, base_tables):
        tdf1, tdf2 = base_tables
        df = tdf1 | tdf2

        lst = ['A1', 'A1', 'A2', 'A3', 'B5']
        assert df.index.tolist() == lst
        # Existe diferencia entre los records 0 y 1 con 'indice 'A1
        assert not (df.iloc[0] == df.iloc[1]).all()
        # La diferencia entre los records 0 y 1 es un campo y ese campo es el campo value.
        assert sum(df.iloc[0].apply(str) != df.iloc[1].apply(str)) == 1 and not (df.iloc[0] == df.iloc[1]).value
        # No existe diferencia entre los campos pra los records diferentes a 'A1
        lst = set(tdf1.df.index) & set(tdf2.df.index)
        assert all((tdf1.df.loc[x].apply(str) == tdf2.df.loc[x].apply(str)).all() for x in lst if x != 'A1')
        pass

    def test_intersection_operator(self, base_tables):
        tdf1, tdf2 = base_tables
        df = tdf1 & tdf2

        lst = ['A2', 'A3']
        assert df.index.tolist() == lst
        # No existe difrencia entre los campos de los records en lst.
        assert all((tdf1.df.loc[x].apply(str) == tdf2.df.loc[x].apply(str)).all() for x in lst)
        # Cualquier otro record existente tanto en tdf1.df como en tdf2.td presenta al 
        # menos un campo diferente.
        clst = (set(tdf1.df.index.tolist()) & set(tdf2.df.index.tolist())) - set(lst)
        assert not all((tdf1.df.loc[x] == tdf2.df.loc[x]).all() for x in clst)
        pass

    def test_xor_operator(self, base_tables):   # symetric_difference    
        tdf1, tdf2 = base_tables
        df = tdf1 ^ tdf2

        lst = ['A1', 'A1', 'B5']
        counter = {key: lst.count(key) for key in set(lst)}
        # Registros con counter[key] == 1 existe solo en una de las df,
        assert all(
            (key in tdf1.df.index and key not in tdf2.df.index) or 
            (key not in tdf1.df.index and key in tdf2.df.index) 
            for key in lst if counter[key] == 1
            )
        # registros con counter[key] == 2 existen en ambas df.
        assert all(
            (key in tdf1.df.index and key in tdf2.df.index)
            for key in lst if counter[key] == 2
            )
        # registros con counter[key] == 2 difieren en al menos un campo.
        assert not all((tdf1.df.loc[x] == tdf2.df.loc[x]).all() for x in lst if counter[x] == 2)
        pass

    def test_difference_operator(self, base_tables):
        tdf1, tdf2 = base_tables
        
        df1 = tdf1 - tdf2
        lst1 = ['A1']
        assert df1.index.tolist() == lst1
        # lst contiene registros que existiendo en tdf1 no existen tdf2 o registros que 
        # existiendo en ambas tdf, los registros difieren en al menos un campo.
        only_tdf1 = set(tdf1.df.index) - set(tdf2.df.index)
        assert not tdf2.df.index.isin(only_tdf1).all()
        cmmn = set(tdf1.df.index) & set(tdf2.df.index) & set(lst1)
        assert not all((tdf1.df.loc[x].apply(str) == tdf2.df.loc[x].apply(str)).all() for x in cmmn)

        df2 = tdf2 - tdf1
        lst2 = ['A1', 'B5']
        assert df2.index.tolist() == lst2
        only_tdf2 = set(tdf2.df.index) - set(tdf1.df.index)
        assert not tdf1.df.index.isin(only_tdf2).all()
        cmmn = set(tdf1.df.index) & set(tdf2.df.index) & set(lst2)
        assert not all((tdf1.df.loc[x].apply(str) == tdf2.df.loc[x].apply(str)).all() for x in cmmn)
        # Equivalencia con el operador "^" para la unión de las diferencias.
        assert TableComparator(TableComparator(df1) ^ TableComparator(df2)) == TableComparator(tdf1 ^ tdf2)
        pass

    def test_difference(self, base_tables):
        tdf1, tdf2 = base_tables
        df1 = tdf1.difference(tdf2)
        df2 = tdf1 - tdf2
        # Equivalencia con el operador "-" para field=None
        assert TableComparator(df1) == TableComparator(df2)
        fields = list(set(tdf1.df.columns.tolist()) - set(['value']))
        # Cuando la diferencia se hace sobre campos en que existe equivalencia, 
        # la base de datos resultante es vacía.
        assert tdf1.difference(tdf2, fields=fields).empty
        # Equivalencia field = None y field = df.columns
        assert TableComparator(tdf1.difference(tdf2, fields=None)) == TableComparator(tdf1.difference(tdf2, fields=fields + ['value']))
        pass

    def test_symmetric_difference(self, base_tables):
        tdf1, tdf2 = base_tables
        df = tdf1.symmetric_difference(tdf2)
        # Equivalencia con el operador "^" para fields igual a None.
        assert TableComparator(df) == TableComparator(tdf1 ^ tdf2)
        fields = ['code']
        df = tdf1.symmetric_difference(tdf2, fields=fields)
        assert df.index.tolist() == ['B5']
        pass

    def test_eq(self, base_tables):
        tdf1, tdf2 = base_tables
        assert tdf1 == tdf1
        assert not tdf1 == tdf2
        pass

    def test_ne(self, base_tables):
        tdf1, tdf2 = base_tables
        assert tdf1 != tdf2
        assert not tdf1 != tdf1
        pass

