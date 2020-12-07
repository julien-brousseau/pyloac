# Utilities
from com.sun.star.util import Date

# Dialog control
import uno
import unohelper
from com.sun.star.awt import XActionListener

# Debugging tools
from apso_utils import xray, mri, msgbox 
 
# --------------------------------------------------------------------------------------------------------
# Class FORM
 
class Form:

  def __init__(self, section):
    self.__section = section
    self.Data = None

  # Show form dialog based on section's model
  def Open(self):
    
    # Dialog layout
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dialogModel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialogModel.PositionX = 100
    dialogModel.PositionY = 100
    dialogModel.Width = 200
    dialogModel.Height = 400
    dialogModel.Title = "Submit"

    # Fields
    for index, model in enumerate(filter(lambda m: not m['autovalue'], self.__section.Model)):
    # for index, model in enumerate(self.__section.Model):

      # Convert types to UNO objects
      types = {
        'String': 'UnoControlEditModel',
        'Integer': 'UnoControlEditModel',
        'Date': 'UnoControlDateFieldModel',
        'DateTime': 'UnoControlDateFieldModel',
        'Amount': 'UnoControlNumericFieldModel',
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
        dateStr = self.__section.Today
        field.Dropdown = True
        field.Spin = True
        field.Date = Date(*map(lambda x: int(x), dateStr.split('-')[::-1]))
        field.Text = dateStr
        field.DateFormat = 11
      elif model['type'] == 'Amount':
        field.Spin = True
        field.Value = 0
        field.DecimalAccuracy = 2
      else:
        field.Text = model['default']

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
  
  # Validate form and handle success/failure
  def Save(self, formDialog):
    # Allowed fields
    columns = self.__section.FieldNames()
    # Filter all allowed fields in dialog
    filteredControls = filter(lambda c: c.Model.Name in columns, formDialog.Controls)
    # Fetch array of values from dialog
    formValues = list(map(lambda c: c.Text, filteredControls))
    # Create a copy of the Section's Model and add form values 
    self.Data = list(map(lambda model, value: dict(model, value = value), self.__section.Model, formValues))
    
    # Reset label color
    labels = list(filter(lambda e: e.Model.Name[-5:] == 'Label', formDialog.Controls))
    for label in labels:
      label.Model.TextColor = 0x000000

    # Validate fields and return errors if any
    errors = self.__validateForm()
    if errors:
      # Change label color for error fields and show errors in log
      for error in errors:
        self.__section.Error(error['error'])
        formDialog.getControl(error['field'].capitalize() + 'Label').Model.TextColor = 0xFF0000
      # Return error to dialog
      return errors
    
    # Add a new line if validation passes
    self.__section.AddNewLine(self.Data)
    return False
  
  # Validate fields against their required property
  def __validateForm(self):
    # Filter out metadata fields
    ignoreFields = ['Id', 'User', 'TS']
    data = list(filter(lambda f: f['field'] not in ignoreFields, self.Data))
    # Validate each remaining field
    validated = list(map(lambda d: dict(d, error = d['label'] + ' field is required') if (d['required'] and d['value'] == '') else d, data))
    # Filter out valid fields
    errorFields = list(filter(lambda f: f.get('error'), validated))
    # Return error fields array or None if no errors
    return errorFields if len(errorFields) else None

# Listeners
class Submit(unohelper.Base, XActionListener):
  def __init__(self, dialogObj, form):
    self.__dialogObj = dialogObj
    self.__form = form
  def actionPerformed(self, actionEvent): 
    # Save form and close dialog if success
    errors = self.__form.Save(self.__dialogObj)
    if not errors: self.__dialogObj.endDialog(1)

