from typing import Dict, List, TypedDict
from enum import Enum, Flag

__all__ = [
    'alpha_code', 
    'TABLE_DATA_MAP', 'ColumnSpec', 'DataFrameStructure', 'structure', 
    'XlFlags', 'Inmutable', 'XlErrors', 'CircularRef', 'EmptyCell',
    'CIRCULAR_REF', 'EMPTY_CELL'
]




def alpha_code(id, nbase=26):
    answ = []
    while id:
        res = id % nbase
        id = id // nbase
        if res == 0:
            res = nbase
            id -=1
        answ.append(chr(ord('A') + res - 1))
    return ''.join(answ[::-1])


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
        self.value = value or 0
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

    def __str__(self):
        return ""
    

EMPTY_CELL = EmptyCell()
CIRCULAR_REF = CircularRef()
