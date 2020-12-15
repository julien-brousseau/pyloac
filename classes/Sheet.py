# Classes
from Cell import Cell 

# Utilities
from Utils import ClearContent

# --------------------------------------------------------------------------------------------------------
# Class SHEET
# 

class Sheet:

  def __init__(self, name, documentContext):
    self.Instance = documentContext.Sheets[name]
    self.Name = name

  # Returns the CellRange reference to the "range name" arg
  def Range(self, rng):
    return self.Instance.getCellRangeByName(rng) 

  # Clear the contents of the cell range
  def Clear(self, rng):
    ClearContent(self.Range(rng))

  # Returns the number of the next empty row in the sheet
  def NextEmptyRow(self):
    cursor = self.Instance.createCursor()
    cursor.gotoEndOfUsedArea(True)
    return cursor.Rows.Count + 1

  # Returns a list of all cells below a named range
  def GetRangeAsList(self, rng):
    a = []
    c = Cell(rng, self)
    while c.value():
      a.append(c.value())
      c.offset(0, 1)
    return a[1:]