# Imports
from Sheet import Sheet 
from Cell import Cell 
from datetime import datetime

from apso_utils import xray, mri, msgbox 
 
# --------------------------------------------------------------------------------------------------------
# Class SECTION
# Wrapper for document context and associated data sheet
# Requirements:
#   - Sheet in the current document with the same name as constructor argument (ex: 'Transactions')
#   - Sheet in the current document for section settings, named with appended "Data" (ex: 'TransactionsData')
#     containing the following named cells:
#     - 'Error'         => Error log for the section
#     - 'FirstRow'      => First row of the list section in main sheet
#     - 'DataFirstRow'  => First row of the Model in the data sheet

 
class Section:

  def __init__(self, name, documentContext):
    self.Name = name
    self.__documentContext = documentContext

    # Associated Sheets
    self.__sheet = Sheet(self.Name, self.__documentContext)
    self.__dataSheet = Sheet(self.Name + 'Data', self.__documentContext)

    # Custom settings from Data Sheet
    self.__firstRow = Cell('FirstRow', self.__dataSheet).value()
    self.__dataFirstRow = Cell('DataFirstRow', self.__dataSheet).value()

    # Section data model
    self.refreshModel()
    self.__columns = [*map(lambda col: {'label': col['label'], 'index': col['column']}, self.Model)]
      
  # Fetch column info on the current section Data Sheet
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
  
  # Externally callable data refresh 
  def refreshModel(self):
    self.Model = self.__ModelFromData()

  # Clear and rebuild column headers for current section
  def BuildColumnHeaders(self):
    self.refreshModel()
    row = self.__firstRow
    self.__sheet.Clear(f"A{row}:Z{row}")
    cell = Cell([0, row - 1], self.__sheet)
    for column in self.__columns:
      cell.move(column['index'] - 1)
      cell.setValue(column['label'].upper())

  # Add an error to Data sheet's Error log (or clear log with no argument)
  def Error(self, msg = None):
    cell = Cell('Error', self.__dataSheet)
    if not msg: cell.Clear()
    else: 
      log = str(cell.value()) if cell.value() != None else ''
      ts = datetime.now().strftime("%H:%M:%S")
      cell.setValue(ts + ': ' + str(msg) + '\n' + log)
            
  def SaveForm(self):
    msgbox("blop")