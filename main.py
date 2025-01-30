# pylint: disable=F0401
# Module management
import sys, os, glob, importlib, platform, uno

# Create APSO instance to prevent apso_utils loading errors
ctx = XSCRIPTCONTEXT.getComponentContext()
ctx.ServiceManager.createInstance('apso.python.script.organizer.impl')
  
# Debugging tools 
from apso_utils import xray, mri, msgbox 
 
# Reference to current document
def This(): 
  return XSCRIPTCONTEXT.getDocument()

# Fetch scripts directory from Settings Sheet
SCRIPTS_DIRECTORY = This().Sheets['Settings'].getCellRangeByName(platform.system() + 'ScriptsDirectory').String
  
# Force module reloading to clear cache (or else modules are not rebuilt on edit)
classes_dir = SCRIPTS_DIRECTORY + '/classes'
sys.path.append(classes_dir)  
for src_file in glob.glob(os.path.join(classes_dir, '*.py')):
  name = os.path.basename(src_file)[:-3]  
  importlib.import_module(name)  
  importlib.reload(sys.modules[name])   
  
# Import classes 
from Section import Section
from Sheet import Sheet 
from Cell import Cell 
from Utils import LODateToString, PropValue, ColumnLabel


# -------------------------------------------------------------------
# Global helpers

# Returns the coordinates of currently selected cell
def CurrentSelection():
  selection = This().CurrentSelection.CellAddress
  return [selection.Column, selection.Row] 

# Test button  
def blop(self = None):
  pass

settings = Sheet('Settings', This())


# -------------------------------------------------------------------
# Transactions helpers

# Filter sheet by transaction type
def ApplyFilter(self):
  field = 'Type'
  section = Section('Transactions', This())
  value = section.Form().getByName('FldFilter').SelectedValue
  if not value: return None
  columnIndex = [*filter(lambda f: f['label'] == field, section.Columns)][0]['index']
  section.Sheet.Filter(section.FirstRow + 1, columnIndex - 1, value)

# Make all rows visible and reset filter field
def ClearFilters(self):
  section = Section('Transactions', This())
  section.Sheet.ToggleVisibleRows()
  section.Form().getByName('FldFilter').SelectedValue = ''
 
# Auto-refresh values for filter fields based on section data (called on field focus)
def RefreshFilterValues(self):
  section = Section('Transactions', This())  
  values = section.ListFieldValues('Type')
  section.Form().getByName('FldFilter').StringItemList = values

# Button - Open New transaction form dialog
def OpenTransactionsForm(self): 
  section = Section('Transactions', This())  
  section.OpenForm()       
 
# Button - Generate transactions sheet (column headers, etc)
def GenerateTransactionsSheet(self):
  section = Section('Transactions', This())  
  section.ClearSheet()
  section.BuildColumnHeaders()

# Button - Toggle Settings
def ToggleTransactionsSettings(self):
  settings = Sheet('TransactionsData', This())  
  settings.toggle()

 
# -------------------------------------------------------------------
# Planification helpers

# Record the currently selected planned transaction into Transactions sheet and flags it as recorded 
# by prepending an underscore, preventing it from being calculated as an active transaction in Soldes sheet
def SaveCellAsTransaction(self):
  section = Section('Transactions', This())
  sheet = Sheet('Planification', This())

  # This character is appended when a cell is saved as transaction
  ignoreFirstCharacter = '_'
 
  # Fetch selected cell
  cell = Cell(CurrentSelection(), sheet)
  cellContent = cell.value()

  # Ignore Empty cells, Cells containing already copied data and Cells out of range (date column or headers)
  if not cellContent or str(cellContent)[0] == ignoreFirstCharacter or cell.coords()[0] < 1 or cell.coords()[1] < 13:
      section.Error('Invalid selection')
      return None

  # Add an underscore to the cell value to prevent it from being calculated in Soldes
  cell.setValue(ignoreFirstCharacter + str(cellContent))

  # Fetch fields and values from header and merge them in dict as field:value
  firstFieldRow = '3'
  nbFields = 11
  values = list(map(lambda v: v or '', sheet.GetRangeAsList(cell.address()[0] + firstFieldRow , nbFields)))
  fields = dict(list(map(lambda f: (f, values.pop(0)), sheet.GetRangeAsList('A' + firstFieldRow , nbFields))))
 
  fields['Amount'] = float(cellContent)

  # Set correct date instead of day
  fields['Date'] = LODateToString(Cell('A' + str(cell.coords()[1] + 1), sheet).value())
  
  # Map values to section model, then save as new transaction
  model = list(filter(lambda f: not f['autovalue'], section.Model))
  data = list(map(lambda f: dict(f, value = fields[f['field']] if f['field'] in fields else ' '), model))
  section.AddNewLine(data)
     
# Show/hide transaction details header in Planification
def TogglePlanificationHeaders(self):
  sheet = Sheet('Planification', This())
  visible = sheet.Range('A2').getRows()
  sheet.ToggleVisibleRows('A3:A16', not visible)


# -------------------------------------------------------------------
# Soldes helpers

# Soldes - Compile cell values as numbers for all rows up to selected row
def CompileSoldes(self = None):
  sheet = Sheet('Soldes', This())
  selectedCell = Cell(CurrentSelection(), sheet)

  firstColumn = 1
  labelsRow = 3
  
  # Break if selected row is part of header
  if selectedCell.coords()[1] < labelsRow + 1: 
    return msgbox('Invalid selection')

  # Loop through all rows
  lastColumn = sheet.NextEmptyColumn() - 1; 
  for row in range(labelsRow + 2, selectedCell.coords()[1] + 2):

    # Skip to the first un-rendered row
    if not Cell([firstColumn, row - 1], sheet).containsFormula(): continue

    # Loop through all columns, and compile formulas into numbers
    for col in range(firstColumn, lastColumn):
      cell = Cell(ColumnLabel(col) + str(row), sheet)
      cellValue = cell.value()
      cell.setValue(cellValue, 'Integer')
    
  return


# -------------------------------------------------------------------
# Factures helpers

# TODO: change PropertyValue for new functions everywhere
# TODO: change settings for global object?

def PrintFacture(self = None):
  facturesDirectory = 'Factures'
  facturePrintRange = 'A1:F47'

  facturesSheet = Sheet('Factures', This())
  row = int(Cell(CurrentSelection(), facturesSheet).address()[1])

  # Check if row is valid
  clientCode = Cell('B' + str(row), facturesSheet).value();
  if (row < 2) or (not clientCode): 
    return msgbox('Invalid selection')
  
  # Check if facture exists 
  offset = int(Cell('ClientRowOffset', Sheet('Settings', This())).strval())
  sheetName = str(row - offset) + '.' + clientCode;
  if not sheetName in This().Sheets.ElementNames: 
    return msgbox('Sheet not found: ' + sheetName)

  # Set facture name
  factureSheet = Sheet(sheetName, This());
  num = Cell('A' + str(row), facturesSheet).strval()[2:]
  date = Cell('D' + str(row), facturesSheet).strval()
  factureName = num + '-' + date + '_' + clientCode + '.pdf'

  # Build facture path - had to use workaround to avoid pathlib parent bug
  factureUrl = '/'.join(uno.fileUrlToSystemPath(This().URL).split('/')[:-1]) + '/' + facturesDirectory + '/' + factureName

  # Set sent date
  Cell('K' + str(row), facturesSheet).setValue(date)

  # Render specific fields in facture sheet so their values become static
  # Client info fields
  for i in range(9, 14):
    Cell('B' + str(i), factureSheet).toString()
  # Date and number
  Cell('D9', factureSheet).toString()
  Cell('E9', factureSheet).toString()
  # Prices
  for i in range(17, 45): 
    Cell('E' + str(i), factureSheet).toString('Amount')
 
  # Save as pdf
  This().storeToURL('file://' + factureUrl, PropValue({
      'FilterName': 'calc_pdf_Export',
      'FilterData': PropValue({ 'Selection': factureSheet.Range(facturePrintRange) }, True),
  }))
 
def CreateFacture(self = None):
  facturesSheet = Sheet('Factures', This())

  # Get client data and generate sheet name
  client = Cell('FactureClient', facturesSheet).strval()
  clientCode = Cell('FactureClientCode', facturesSheet).strval()
  num = Cell('NextFactureNumber', Sheet('Settings', This())).strval()
  factureSheetName = num + '.' + clientCode;

  # Duplicate template and set client field
  This().Sheets.copyByName('FactureTemplate', factureSheetName, len(This().Sheets))
  factureSheet = Sheet(factureSheetName, This())
  Cell('I9', factureSheet).setValue(client)

  # Add new client facture to the list
  clientRow = int(Cell('ClientRowOffset', settings).value() + Cell('NextFactureNumber', settings).value())
  Cell('C' + str(clientRow), facturesSheet).setValue(client)

# 
def factureToTransactions(self = None):
  facturesSheet = Sheet('Factures', This())
  row = int(Cell(CurrentSelection(), facturesSheet).address()[1])

  # Check if row is valid
  clientCode = Cell('B' + str(row), facturesSheet).value();
  if (row < 2) or (not clientCode): 
    return msgbox('Invalid selection')
  
  # Check if facture exists 
  offset = int(Cell('ClientRowOffset', Sheet('Settings', This())).strval())
  sheetName = str(row - offset) + '.' + clientCode;
  if not sheetName in This().Sheets.ElementNames: 
    return msgbox('Sheet not found: ' + sheetName)
  
  section = Section('Transactions', This())
  model = list(filter(lambda f: not f['autovalue'], section.Model))

  date = Cell('D' + str(row), facturesSheet).value()
  num = Cell('A' + str(row), facturesSheet).value()
  totalAmount = Cell('F' + str(row), facturesSheet).value()
  taxesAmount = Cell('G' + str(row), facturesSheet).value() + Cell('I' + str(row), facturesSheet).value()

  today = Cell('Today', Sheet('Settings', This())).strval()

  # Create payment transaction
  fields = { 'Date': today, 'Amount': totalAmount, 'Type': 'Dépôt', 'Description': 'Paiement facture #' + num, 'File': 'Comptes', 'FromAccount': '', 'ToAccount': 'The Source', 'User': 'Julien', 'Shared': '' }
  data = list(map(lambda f: dict(f, value = fields[f['field']] if f['field'] in fields else ' '), model))
  section.AddNewLine(data)

  # Create taxes transaction
  fields = { 'Date': today, 'Amount': taxesAmount, 'Type': 'Transfert', 'Description': 'TR taxes #' + num, 'File': 'Comptes', 'FromAccount': 'The Source', 'ToAccount': 'The Taxes', 'User': 'Julien', 'Shared': '' }
  data = list(map(lambda f: dict(f, value = fields[f['field']] if f['field'] in fields else ' '), model))
  section.AddNewLine(data)

  # Mark facture as paid
  Cell('M' + str(row), facturesSheet).setValue(today)

  # Hide facture sheet
  Sheet(sheetName, This()).hide()