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
