# Module management
import sys, os, glob, importlib 

# Debugging tools 
from apso_utils import xray, mri, msgbox 
 
# Reference to current document
def This(): 
  return XSCRIPTCONTEXT.getDocument()

# Fetch scripts directory from Settings Sheet
SCRIPTS_DIRECTORY = This().Sheets['Settings'].getCellRangeByName('ScriptsDirectory').String
  
# Force module reloading to clear cache (or else modules are not rebuilt on edit)
classes_dir = SCRIPTS_DIRECTORY + "/classes"
sys.path.append(classes_dir)  
for src_file in glob.glob(os.path.join(classes_dir, '*.py')):
  name = os.path.basename(src_file)[:-3]  
  importlib.import_module(name)  
  importlib.reload(sys.modules[name])   
  
# Import classes 
from Section import Section
from Sheet import Sheet 
from Cell import Cell 
from Utils import LODateToString 
 
# -------------------------------------------------------------------
  
# Returns the coordinates of currently selected cell
def CurrentSelection(celladdress = False):
  selection = This().CurrentSelection.CellAddress
  return [selection.Column, selection.Row]

# Test button 
def blop(self):
  return 0
    
# -------------------------------------------------------------------
    
# Filter sheet by transaction type
def ApplyFilter(self):
  field = 'Type'
  section = Section('Transactions', This())
  value = section.Form().getByName('FldFilter').SelectedValue
  if not value: return None
  columnIndex = [*filter(lambda f: f['label'] == field, section.Columns)][0]['index']
  section.Sheet.Filter(section.FirstRow + 1, columnIndex - 1, value)
        
# Make all rows visible and reset 
def ClearFilters(self):
  section = Section('Transactions', This())
  section.Sheet.ToggleVisibleRows()
  section.Form().getByName('FldFilter').SelectedValue = ''
 
# Button - Open New transaction form dialog
def OpenTransactionsForm(self): 
  section = Section('Transactions', This())  
  section.OpenForm()       
 
# Button - Generate transactions sheet (column headers, etc)
def GenerateTransactionsSheet(self):
  section = Section('Transactions', This())  
  section.ClearSheet()
  section.BuildColumnHeaders()

# Planification - Record the currently selected planned transaction into Transactions sheet
# and flags it as recorded by prepending an underscore, preventing it from being calculated
# as an active transaction in Soldes sheet
def SaveCellAsTransaction(self):
  section = Section('Transactions', This())
  sheet = Sheet('Planification', This())

  # This character is appended when a cell is saved as transaction
  ignoreFirstCharacter = '_'

  # Fetch selected cell
  cell = Cell(CurrentSelection(), sheet)
  cellContent = cell.value()

  # Ignore 1. Empty cells / 2. Cells containing already copied data / 3. Cells out of range (date column and headers)
  if not cellContent or str(cellContent)[0] == ignoreFirstCharacter or cell.coords()[0] < 1 or cell.coords()[1] < 13:
      msgbox("Invalid selection")
      return None

  # Add an underscore to the cell value to prevent it from being calculated in Soldes
  cell.setValue(ignoreFirstCharacter + str(cellContent))

  # Fetch fields and values from header and merge them in dict as field:value
  values = list(map(lambda v: v or '', sheet.GetRangeAsList(cell.address()[0] + '1', 7)))
  fields = sheet.GetRangeAsList('A1', 7)
  fields = dict(list(map(lambda f: (f, values.pop(0)), fields)))

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
  sheet.ToggleVisibleRows('A2:A11', not visible)

# Auto-refresh values for filter fields based on section data (called on field focus)
def RefreshFilterValues(self):
  section = Section('Transactions', This())  
  values = section.ListFieldValues('Type')
  section.Form().getByName("FldFilter").StringItemList = values
