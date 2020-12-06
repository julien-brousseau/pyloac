# Module management
import sys, os, glob, importlib 

# Debug tools
from apso_utils import xray, mri, msgbox 

# test
import uno
import unohelper
from com.sun.star.awt import XActionListener

FORM_PADDING = 10
LABEL_SPACING = 5
FIELD_SPACING = 5

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

# App       
def blop(self): 
  OpenTransactionsForm(self)

# Generate Sheet   
SECTION = Section('Transactions', This())    


class MyActionListener( unohelper.Base, XActionListener ):
  def __init__(self, dialogObj):
    self.__dialogObj = dialogObj
    # self.__dialogObj.Controls[0].Model.Label
 
  def actionPerformed(self, actionEvent):
    xray(self.__dialogObj.Controls[1])
    for blop in self.__dialogObj.Controls:
      SECTION.Error(blop.Model.Type)
    SaveTransactionsForm()

# Interface button call
def GenerateTransactionsSheet(self):  
  SECTION.BuildColumnHeaders()

# Temporary called as test
def OpenTransactionsForm(self):
    # Dialog layout
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dialogModel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialogModel.PositionX = 100
    dialogModel.PositionY = 100
    dialogModel.Width = 200
    dialogModel.Height = 300
    dialogModel.Title = "Add a transaction"

    y = 0
    # Fields
    for model in SECTION.Model:

      fieldH = 14
      labelH = 10

      fieldWidth = model['width']
      columnIndex = model['column']
      fieldName = model['field']

      label = dialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel" )
      label.PositionX = FORM_PADDING
      label.PositionY = FORM_PADDING + (y * (fieldH + FIELD_SPACING) ) + ((y - 1) * labelH)
      label.Width  = fieldWidth * 15 
      label.Height = labelH
      label.Name = fieldName + 'Label'
      label.Label = model['label']

      field = dialogModel.createInstance("com.sun.star.awt.UnoControlEditModel" )
      field.PositionX = FORM_PADDING
      field.PositionY = FORM_PADDING + (y * (fieldH + FIELD_SPACING) ) + (y * labelH)
      field.Width  = fieldWidth * 15 
      field.Height = fieldH
      field.Name = fieldName
      field.TabIndex = columnIndex

      dialogModel.insertByName( fieldName + 'Label', label)
      dialogModel.insertByName( fieldName, field)
      y += 1

    # Button
    button = dialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel" )
    button.PositionX = 50
    button.PositionY  = 30
    button.Width = 50
    button.Height = 14
    button.Name = "Submit"
    button.TabIndex = 99 
    button.Label = "Save"
    dialogModel.insertByName( "Submit", button)

    # create the dialog control and set the model
    DialogObj = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    DialogObj.setModel(dialogModel)

    # add the action listener
    DialogObj.getControl("Submit").addActionListener(MyActionListener(DialogObj))
    # DialogObj.getControl("myButtonName").addActionListener( MyActionListener( DialogObj.getControl( "myLabelName" ), labelModel.Label ))

    # create a peer ??
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.ExtToolkit", ctx)       
    DialogObj.setVisible(False)       
    DialogObj.createPeer(toolkit, None)

    # execute it
    DialogObj.execute()
    DialogObj.dispose()
    
def SaveTransactionsForm():
  SECTION.Error('SAVE!')