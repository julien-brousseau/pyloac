# Utilities
from Utils import ClearContent
import datetime

# Debugging tools
from apso_utils import xray, mri, msgbox

# --------------------------------------------------------------------------------------------------------
# Class CELL
# Interface allowing more control over a cell reference

class Cell:

  # The address parameter can be either [x, y] or 'range'
  # The sheet parameter is a Sheet object
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

  def __dateStrToCalcDate(self, dateStr):
    delta = datetime.date(*map(lambda x: int(x), dateStr.split('-'))) - datetime.date(1899, 12, 30)
    date = float(delta.days) + (float(delta.seconds) / 86400)
    return date
  # def __dateToCalcDate(self, d):
    # blop = datetime.strptime( d, '%Y-%m-%d' )
    # msgbox(str(d))
    # date = float(delta.days) + (float(delta.seconds) / 86400)
    # return 1#date
 
  # Setters
  def setValue(self, value, _type = 'String'):
    if _type in ['Amount', 'Integer']:
      self.__range.Value = value
    elif _type == 'Date':
      date = self.__dateStrToCalcDate(value)
      self.__range.Value = date
    elif _type == 'DateTime':
      self.__range.String = value
    else: 
      self.__range.String = str(value)
    # if type(value) is str: self.__range.String = value
    # self.__range.Value = value
  def Clear(self):
    ClearContent(self.__range)

  # Getters
  def value(self, decimal = False):

    if self.__range.CellContentType.value == 'TEXT': return self.__range.String
    elif self.__range.CellContentType.value == 'VALUE': return '{0:.2f}'.format(self.__range.Value) if decimal else self.__range.Value # int(self.__range.Value)
    elif self.__range.CellContentType.value == 'EMPTY': return None
    else: return 'NOT IN RANGE: ' + self.__range.CellContentType.value
  def toString(self):
    return str(self.__range.String)

  # Helpers
  def __letter(self, offset = 0):
    col = self.__x() + offset
    if not type(col) is int: return None
    if col <= 25: return chr(col + 65)  # single letter
    else: return "A" + chr(col + 65 - 26) # multiple letters TODO: EXPAND
