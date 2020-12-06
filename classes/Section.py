
from Sheet import Sheet 
from Cell import Cell 
from apso_utils import xray, mri, msgbox

# --------------------------------------------------------------------------------------------------------
# Class SECTION

class Section:

  def __init__(self, name, documentContext):
    self.Name = name
    self.__documentContext = documentContext
    self.__sheet = Sheet(self.Name, self.__documentContext)
    self.__dataSheet = Sheet(self.Name + 'Data', self.__documentContext)

    self.__firstRow = Cell('FirstRow', self.__dataSheet).value()
    self.__dataFirstRow = Cell('DataFirstRow', self.__dataSheet).value()
    self.refreshData()
    self.__columns = [*map(lambda col: {'label': col['label'], 'index': col['column']}, self.Data)]
      
  # 
  def __ModelFromData(self):
    # Fetch column headers
    headers = []
    cell = Cell([0, self.__dataFirstRow - 2], self.__dataSheet)
    while cell.value(): 
        headers.append(cell.toString().lower())
        cell.offset(1, 0)
    # Loop through every row and add them as dict to array
    arr = []
    cell = Cell([0, self.__dataFirstRow - 1], self.__dataSheet)
    while cell.value(): # rows 
      r = {}
      for header in headers: #cols
        r[header] = cell.value() 
        cell.offset(1, 0) 
      arr.append(r)
      cell.offset(len(headers) * -1, 1)
    # msgbox(str(arr))
    return arr
  #
  def refreshData(self):
    self.Data = self.__ModelFromData()

  #
  def BuildColumnHeaders(self):
    self.refreshData()
    row = self.__firstRow
    self.__sheet.Clear(f"A{row}:Z{row}")
    cell = Cell([0, row - 1], self.__sheet)
    for column in self.__columns:
      cell.move(column['index'] - 1)
      cell.setValue(column['label'].upper())
    return None