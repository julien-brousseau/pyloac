# Imports
from Sheet import Sheet 
from Cell import Cell 
from datetime import datetime
from com.sun.star.util import Date

# test
import uno
import unohelper
from com.sun.star.awt import XActionListener

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
    self.Model = []
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

  # Show a dialog form based on the Section's Model
  def OpenForm(self):

    # Dialog layout
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dialogModel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialogModel.PositionX = 100
    dialogModel.PositionY = 100
    dialogModel.Width = 200
    dialogModel.Height = 300
    dialogModel.Title = "Submit"

    # Fields
    for index, model in enumerate(self.Model):

      # Convert types to UNO objects
      types = {
        'String': 'UnoControlEditModel',
        'Integer': 'UnoControlEditModel',
        'Date': 'UnoControlDateFieldModel',
        'Amount': 'UnoControlCurrencyFieldModel',
        'CheckBox': 'UnoControlCheckBoxModel'
      }
      
      # Position variables
      FORM_PADDING = 10
      FIELD_SPACING = 5
      FIELD_HEIGHT = 14
      LABEL_HEIGHT = 10

      # Field label
      label = dialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel" )
      label.PositionX = FORM_PADDING
      label.PositionY = FORM_PADDING + (index * (FIELD_HEIGHT + FIELD_SPACING) ) + ((index - 1) * LABEL_HEIGHT)
      label.Width  = model['width'] * 15 
      label.Height = LABEL_HEIGHT
      label.Name = model['field'] + 'Label'
      label.Label = model['label']

      # Field control
      field = dialogModel.createInstance("com.sun.star.awt." + types[model['type']] )
      field.PositionX = FORM_PADDING
      field.PositionY = FORM_PADDING + (index * (FIELD_HEIGHT + FIELD_SPACING) ) + (index * LABEL_HEIGHT)
      field.Width  = model['width'] * 15 
      field.Height = FIELD_HEIGHT
      field.Name = model['field']
      field.TabIndex = model['column']

      # Type-dependant properties
      if model['type'] == 'Date':
        field.Dropdown = True
        field.Spin = True
        field.Date = Date(5, 10, 2020)
        field.Text = "2020-10-05"
        field.DateFormat = 11
      if model['type'] == 'Amount':
        field.Spin = True
        field.Value = 0

      # Insert fields
      dialogModel.insertByName( model['field'] + 'Label', label)
      dialogModel.insertByName( model['field'], field)

    # Button
    button = dialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel" )
    button.PositionX = 100
    button.PositionY  = 30
    button.Width = 50
    button.Height = 14
    button.Name = "Submit"
    button.TabIndex = 99 
    button.Label = "Save"
    dialogModel.insertByName("Submit", button)

    # create the dialog control and set the model
    DialogObj = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    DialogObj.setModel(dialogModel)

    # add the action listener
    DialogObj.getControl("Submit").addActionListener(Submit(DialogObj, self))

    # create a peer ??
    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.ExtToolkit", ctx)       
    DialogObj.setVisible(False)       
    DialogObj.createPeer(toolkit, None)

    # execute it
    DialogObj.execute()
    DialogObj.dispose()
    

class Submit(unohelper.Base, XActionListener):

  def __init__(self, dialogObj, section):
    self.__dialogObj = dialogObj
    self.__section = section

  def actionPerformed(self, actionEvent): 
    # Allowed fields (TODO: !Hardcoded)
    columns = ['Date', 'Amount', 'Type', 'Description', 'File', 'FromAccount', 'ToAccount', 'Deductible', 'GigId', 'Id', 'User', 'TS']
    # Filter all allowed fields in dialog
    filteredControls = filter(lambda c: c.Model.Name in columns, self.__dialogObj.Controls)
    # Fetch array of values from dialog
    formValues = list(map(lambda c: c.Text, filteredControls))
    # Create a copy of the Section's Model and add form values 
    blop = list(map(lambda model, value: dict(model, value = value), self.__section.Model, formValues))

    msgbox(str(blop))
    # self.__section.Error(blop)



    # self.__section.Error([*output])
    # self.__section.Error(self.__dialogObj.Controls[2].Model.Value)
    # self.__section.validateForm()
    # for blop in self.__dialogObj.Controls:
    #   SECTION.Error(blop)
    # self.__dialogObj.Controls[0].Model.Label
