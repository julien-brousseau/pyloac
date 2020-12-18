# Classes
from Sheet import Sheet 
from Cell import Cell 
from Form import Form
from Utils import ColumnLabel
 
# Utilities 
import uno
from datetime import datetime
from com.sun.star.beans import PropertyValue
from com.sun.star.util import SortField

# Debugging tools
from apso_utils import xray, mri, msgbox 
 
# --------------------------------------------------------------------------------------------------------
# Class SECTION
# Wrapper for document context and associated data sheet
# Requirements:
#   - Sheet in the current document with the same name as constructor argument (ex: 'Transactions')
#   - Sheet in the current document for section settings, named with appended "Data" (ex: 'TransactionsData')
#     containing the following named cells:
#     - 'Error'           => Error log for the section (preferably multiline merged cells)
#     - 'FirstRow'        => First row of the list section in main sheet
#     - 'DataFirstRow'    => First row of the Model in the data sheet
#     - 'NextId'          => Automatic ID auto-increment (MAX() + 1 of transactions ids)
 
class Section:

  def __init__(self, name, documentContext):
    self.Name = name
    self.__documentContext = documentContext

    # Associated Sheets
    self.__sheet = Sheet(self.Name, self.__documentContext)
    self.__dataSheet = Sheet(self.Name + 'Data', self.__documentContext)

    # Custom settings from Data Sheet
    self.__firstRow = int(Cell('FirstRow', self.__dataSheet).value())
    self.__dataFirstRow = int(Cell('DataFirstRow', self.__dataSheet).value())

    # Section data model
    self.Model = []
    self.refreshModel()
    self.__columns = [*map(lambda col: {'label': col['label'], 'index': col['column']}, self.Model)]

    self.Today = Cell('Today', Sheet('Settings', self.__documentContext)).toString()
      
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
    return arr
  
  # Externally callable data refresh 
  def refreshModel(self):
    self.Model = self.__ModelFromData()
 
  # Returns possible values for a list field
  # The field must have a named range ('Values_' + fieldName) as list header
  def ListFieldValues(self, fieldName):
    return self.__dataSheet.GetRangeAsList('Values_' + fieldName)
 
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

  # Show a dialog form based on the Section's Model
  def OpenForm(self):
    form = Form(self)
    form.Open() 

  # Write a new line in section Sheet with 
  def AddNewLine(self, data):

    # Metadata - Id and Timestamp (autovalue)
    meta = list(filter(lambda f: f['field'] in ['Id', 'TS'], self.Model))
    meta[0]['value'] = Cell('NextId', self.__dataSheet).value()
    meta[1]['value'] = datetime.now().strftime("%Y-%d-%m %H:%M:%S")
    completeData = [*data, *meta]

    # Next row index
    row = self.__sheet.NextEmptyRow()

    # Write fields
    for field in completeData:
      coords = [field['column'] - 1, row - 1]
      Cell(coords, self.__sheet).setValue(field['value'], field['type'])
    
    # Sort rows
    self.SortRange()
    # self.Error('New row added: ' + str(data))
 
  # Return a list of all Sections' field names
  def FieldNames(self):
    return list(map(lambda col: col['field'], self.Model))

  # Sorting Columns
  def SortRange(self, col1 = 0, asc1 = False, col2 = 25, asc2 = False):
    cellRange = self.__sheet.Range("A4:Z999")
    sortFields = []
    sortFields.append(NewSortField(col1, asc1))
    # if col2:
    #   sortFields.append(NewSortField(col2, asc2))
    sortDesc = [PropertyValue()]
    sortDesc[0].Name = "SortFields"
    sortDesc[0].Value = uno.Any('[]com.sun.star.util.SortField', sortFields)
    cellRange.sort(sortDesc)

  # Clear the data (non-header) content of the sheet
  def ClearSheet(self):
    self.__sheet.Clear('A' + str(self.__firstRow + 1) + ':Z9999')
    Cell('Error', self.__dataSheet).setValue('')
    self.Error('Sheet reset')

# >> Used in SortRange
def NewSortField(col, asc):
    field = SortField()
    field.Field = col
    field.SortAscending = asc
    return field
