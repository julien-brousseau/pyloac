# Imports
from com.sun.star.util import Date

# test
import uno
import unohelper
from com.sun.star.awt import XActionListener

from apso_utils import xray, mri, msgbox 
 
# --------------------------------------------------------------------------------------------------------
# Class FORM
 
class Form:

  def __init__(self, section):
    self.__section = section
    self.Data = None

  #  
  def Open(self):
    
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
    for index, model in enumerate(self.__section.Model):

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
  
  #
  def Save(self, formData):
    # Allowed fields (TODO: !Hardcoded)
    columns = ['Date', 'Amount', 'Type', 'Description', 'File', 'FromAccount', 'ToAccount', 'Deductible', 'GigId', 'Id', 'User', 'TS']
    # Filter all allowed fields in dialog
    filteredControls = filter(lambda c: c.Model.Name in columns, formData)
    # # Fetch array of values from dialog
    formValues = list(map(lambda c: c.Text, filteredControls))
    # Create a copy of the Section's Model and add form values 
    self.Data = list(map(lambda model, value: dict(model, value = value), self.__section.Model, formValues))
  
    # Validate form values, then add new line
    if self.Validate():
      self.__section.AddNewLine(self.Data)
    else:
      self.__section.Error('Validation failed')
  
  #
  def Validate(self):
    return True

# Listeners
class Submit(unohelper.Base, XActionListener):

  def __init__(self, dialogObj, form):
    self.__dialogObj = dialogObj
    self.__form = form

  def actionPerformed(self, actionEvent): 
    self.__form.Save(self.__dialogObj.Controls)

