# Utilities
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, FORMULA

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
  def Clear(self, range):
    self.Range(range).clearContents(VALUE + DATETIME + STRING + FORMULA)

  # Returns the number of the next empty row in the sheet
  def NextEmptyRow(self):
    cursor = self.Instance.createCursor()
    cursor.gotoEndOfUsedArea(True)
    return cursor.Rows.Count + 1
