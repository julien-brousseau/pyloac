# Imports
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, FORMULA

from apso_utils import xray, mri, msgbox

# --------------------------------------------------------------------------------------------------------
# Class CELL
# Interface allowing more control over a cell reference

class Cell:

    def __init__(self, address, sheet):
        self.__sheet = sheet.Instance
        self.__range = self.__sheet.getCellByPosition(address[0], address[1]) if type(address) == list else self.__sheet.getCellRangeByName(address)

    # Change the relative position of the Cell
    def offset(self, x = 0, y = 0):
        self.__range = self.__sheet.getCellByPosition(self.__x() + x, self.__y() + y)
        return self
    
    # Move Cell pointer to new coords range
    def move(self, x = None, y = None):
        mx = x if x != None else self.__x()
        my = y if y != None else self.__y()
        self.__range = self.__sheet.getCellByPosition(mx, my)

    # Coordinates
    def __x(self):
        return self.__range.CellAddress.Column 
    def __y(self):
        return self.__range.CellAddress.Row
    def coords(self):
        return [self.__x(), self.__y()]
    def address(self):
        return self.__letter() + str(self.__y() + 1)
    
    # Setters
    def setValue(self, value):
        if type(value) is str: self.__range.String = value
        else: self.__range.Value = value

    # Getters
    def value(self, decimal = False):
        if self.__range.CellContentType.value == 'TEXT': return self.__range.String
        elif self.__range.CellContentType.value == 'VALUE': return '{0:.2f}'.format(self.__range.Value) if decimal else int(self.__range.Value)
        elif self.__range.CellContentType.value == 'EMPTY': return None
        else: return 'NOT IN RANGE: ' + self.__range.CellContentType.value
    def toString(self):
        return str(self.value())
    
    # Clear the contents of the cell
    def Clear(self):
        self.__range.clearContents(VALUE + DATETIME + STRING + FORMULA)
        
    # Helpers
    def __letter(self, offset = 0):
        col = self.__x() + offset
        if not type(col) is int: return None
        if col <= 25: return chr(col + 65)  # single letter
        else: return "A" + chr(col + 65 - 26) # multiple letters TODO: EXPAND
