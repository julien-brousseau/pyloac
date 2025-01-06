# pylint: disable=F0401
# Module management
import sys, os, glob, importlib 

# Create APSO instance to prevent apso_utils loading errors
# ctx = XSCRIPTCONTEXT.getComponentContext()
# ctx.ServiceManager.createInstance("apso.python.script.organizer.impl")
  
# Debugging tools 
# from apso_utils import xray, mri, msgbox 
 
# Reference to current document
def This(): 
  return XSCRIPTCONTEXT.getDocument()

# Fetch scripts directory from Settings Sheet
SCRIPTS_DIRECTORY = This().Sheets['Settings'].getCellRangeByName('ScriptsDirectory').String
  
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
from Utils import LODateToString 

# -------------------------------------------------------------------
# Global helpers

# Returns the coordinates of currently selected cell
def CurrentSelection():
  selection = This().CurrentSelection.CellAddress
  return [selection.Column, selection.Row] 

# Test button  
def blop():
  SaveFactureAsTransaction()
  pass
    
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
  section.Form().getByName("FldFilter").StringItemList = values

# Button - Open New transaction form dialog
def OpenTransactionsForm(self): 
  section = Section('Transactions', This())  
  section.OpenForm()       
 
# Button - Generate transactions sheet (column headers, etc)
def GenerateTransactionsSheet(self):
  section = Section('Transactions', This())  
  section.ClearSheet()
  section.BuildColumnHeaders()

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
      section.Error("Invalid selection")
      return None

  # Add an underscore to the cell value to prevent it from being calculated in Soldes
  cell.setValue(ignoreFirstCharacter + str(cellContent))

  # Fetch fields and values from header and merge them in dict as field:value
  values = list(map(lambda v: v or '', sheet.GetRangeAsList(cell.address()[0] + '1', 10)))
  fields = dict(list(map(lambda f: (f, values.pop(0)), sheet.GetRangeAsList('A1', 10))))
  fields["Amount"] = float(cellContent)

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
  sheet.ToggleVisibleRows('A2:A13', not visible)

# Soldes - Compile cell values as numbers for all rows up to selected row
def CompileSoldes(self = None):
  sheet = Sheet('Soldes', This())
  selectedCell = Cell(CurrentSelection(), sheet)

  # Reference column to determine if the row has already been compiled (contains formula)
  firstColumnIndex = 1
  # Max column to compile 
  maxColumnIndex = 40
  # First row containing data
  firstRowIndex = 4

  # Break if selected row is part of header
  if selectedCell.coords()[1] < firstRowIndex: return None

  # Loop through rows up to selected row
  cell = Cell([firstColumnIndex, firstRowIndex], sheet)
  while (cell.coords()[1] <= selectedCell.coords()[1]):
    
    # Loop through all cells if not already compiled
    if cell.containsFormula():
      for colIndex in range(firstColumnIndex, maxColumnIndex + 1):

        # Replace formula with compiled value
        cellValue = cell.value()
        cell.setValue(cellValue, 'Integer')
        cell.offset(1, 0)   
      
      # Move cell pointer to first column 
      cell.move(firstColumnIndex, None)  

    # Move cell pointer to next row
    cell.offset(0, 1) 

# -------------------------------------------------------------------
def SaveFactureAsTransaction(self):
  sheet = Sheet('Soldes', This())
  selectedCell = Cell(CurrentSelection(), sheet)
  # msgbox(selectedCell.text())
