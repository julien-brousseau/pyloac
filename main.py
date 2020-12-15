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
    
# -------------------------------------------------------------------   
    
# Generate Sheet
SECTION = Section('Transactions', This())

# Test button  
def blop(self):  
  OpenTransactionsForm(self)           
 
# -------------------------------------------------------------------    

# Button - Open transactions form 
def OpenTransactionsForm(self): 
  SECTION.OpenForm()       
                 
# Button - Generate transactions sheet (column headers, etc)
def GenerateTransactionsSheet(self):
  SECTION.ClearSheet()
  SECTION.BuildColumnHeaders()

# Button - Save transactions form as new line
def SaveTransactionsForm():
  SECTION.Error('SAVE!')      

 

