# Math
from math import floor

# Date 
from datetime import datetime, timedelta 

# Filtering extensions
from com.sun.star.sheet import TableFilterField
from com.sun.star.sheet.FilterOperator import EQUAL

# Cell values
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, FORMULA

# PropertyValue objects management
from com.sun.star.beans import PropertyValue
import uno

# Debugging tools
# from apso_utils import xray, mri, msgbox 

 

# Return the array index of a capital letter (with ASCII offset)
def ColumnIndex(letter = None):

  # Parameter must exist and be a string
  if not type(letter) is str:
    return None

  # Single letter
  if len(letter) == 1: return ord(letter) - 65
  
  # Multiple letters
  l1, l2 = list(letter)
  return 26 + (ord(l1) - 65) * 26 + (ord(l2) - 65)

# Return the array index of a capital letter (with ASCII offset)
def ColumnLabel(col = None):

  # Parameter must be a integer
  if not type(col) is int:
    return None

  # Single letter
  if col <= 25:
    return chr(col + 65)

  # Multiple letters
  l1 = chr(floor(col / 26) - 1 + 65)
  l2 = chr(col % 26 + 65)
  return l1 + l2

# Remove all contents from the specified cell range
def ClearContent(rng):
  rng.clearContents(VALUE + DATETIME + STRING + FORMULA)
  return False

# Returns the standard string format (YYYY-MM-DD) from a LibreOffice date format (integer)
def LODateToString(loDate):
  return (datetime.strptime( "1900", '%Y' ) + timedelta(days = int(loDate) - 2)).strftime("%Y-%m-%d")
 
# Used in Sheet filtering
def FilterField(val):
  field = TableFilterField()
  field.Field = 0
  field.IsNumeric = False
  field.Operator = EQUAL
  field.StringValue = val
  return field

# Return a dict {name: value} as a PropertyValue object
def PropValue(values, any = False):
    ps = tuple([PropertyValue(Name = n, Value = v) for n, v in values.items()])
    if any: ps = uno.Any('[]com.sun.star.beans.PropertyValue', ps)
    return ps