import pandas as pd

from excel_workbook import cell_pattern


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
    
    def __ne__(self, value: object) -> bool:
        return not self.__eq__(value)
    

def excel_df(tbl_str, row_ini, col_ini):
    '''
    Convierte el copy/paste de una tabla de excel a un pd.Series donde el index son las
    direcciones de las celdas y el valor es el str(value_cell)
    tbl_str: str. Copy/paste tabla excel.
    row_ini: str. Fila (row) superior izquierda de la tabla.
    col_ini: str. Columna (col) superior izquierda de la tabla.
    '''
    lst = [x.split('\t') for x in tbl_str.strip().split('\n')]
    ndx1, ndx2 = int(row_ini), int(row_ini) + len(lst)   
    col1, col2 = ord(col_ini), ord(col_ini) + len(lst[0])
    df = (
        pd.DataFrame(lst, index=pd.RangeIndex(ndx1, ndx2), columns=[chr(x) for x in range(col1, col2)])
        .unstack()
        .to_frame()
        .rename(columns={0:'value'})
        .assign(cell=lambda db: db.index.map(lambda x: x[0] + str(x[1])))
        .set_index('cell')
    )
    return df.value
