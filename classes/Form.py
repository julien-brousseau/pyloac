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
# Section sub-wrapper for dialogs and fields management

class Form:

  def __init__(self, section):
    self.__section = section
    self.Data = None

  # Returns a field value depending its type
  def __getFieldValue(self, field):
    if str(field).find('UnoListBoxControl') != -1:
      return field.SelectedItem 
    elif str(field).find('UnoCheckBoxControl') != -1:
      return field.State or ''
    else: 
      return field.Text
  
  # Show form dialog based on section's model
  def Open(self):
    
    # Dialog layout
    # TODO: Make dialog width and height automatic 
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dialogModel = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    dialogModel.PositionX = 100
    dialogModel.PositionY = 100
    dialogModel.Width = 250
    dialogModel.Height = 225
    dialogModel.Title = "Submit"

    # Convert field type to UNO objects
    types = {
      'String': 'UnoControlEditModel',
      'Enum': 'UnoControlListBoxModel',
      'Integer': 'UnoControlEditModel',
      'Date': 'UnoControlDateFieldModel',
      'DateTime': 'UnoControlDateFieldModel',
      'Amount': 'UnoControlNumericFieldModel',
      'CheckBox': 'UnoControlCheckBoxModel' }
    
    # Position variables
    FORM_PADDING = 15
    FIELD_SPACING = 10
    FIELD_HEIGHT = 15
    LABEL_HEIGHT = 10
    yOffset = 0
    xOffset = 0
    currentSection = 1

    # Build all fields except metadata
    fields = filter(lambda m: not m['autovalue'], self.__section.Model)
    for model in fields:

      if currentSection != model['section']: 
        xOffset = 0
        yOffset += (FIELD_HEIGHT + LABEL_HEIGHT + FIELD_SPACING)
        currentSection = model['section']
      _w = model['width'] * 15 

      # Field label
      label = dialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel" )
      label.PositionX = FORM_PADDING + xOffset
      label.PositionY = FORM_PADDING + yOffset
      label.Width  = _w
      label.Height = LABEL_HEIGHT
      label.Name = model['field'] + 'Label'
      label.Label = model['label']

      # Field control
      field = dialogModel.createInstance("com.sun.star.awt." + types[model['type']] )
      field.PositionX = FORM_PADDING + xOffset
      field.PositionY = FORM_PADDING + yOffset + LABEL_HEIGHT
      field.Width  = _w
      field.Height = FIELD_HEIGHT
      field.Name = model['field']
      field.TabIndex = model['column']

      # 
      xOffset += _w + FIELD_SPACING
       
      # Type-dependant properties
      if model['type'] == 'Date':
        field.Dropdown = True
        field.Spin = True
        field.DateFormat = 11
        dateStr = self.__section.Today
        field.Date = Date(*map(lambda x: int(x), dateStr.split('-')[::-1]))
        field.Text = dateStr
      elif model['type'] == 'Amount': 
        field.Spin = True
        field.DecimalAccuracy = 2
        field.Value = model['default']
      elif model['type'] == 'Enum':
        field.LineCount = 10
        field.Dropdown = True
        values = self.__section.ListFieldValues(model['field'])
        if not model['required']:
          values = [''] + values
        field.StringItemList = values
        field.SelectedItems = [values.index(model['default'])] if model['default'] else [0]
      elif model['type'] == 'CheckBox':
        field.State = model['default'] or 0
      else:
        field.Text = model['default'] or ''
 
      # Insert fields
      dialogModel.insertByName( model['field'] + 'Label', label)
      dialogModel.insertByName( model['field'], field)

    # Button
    button = dialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel" )
    button.PositionX = FORM_PADDING
    button.PositionY  = yOffset + 60
    button.Width = 75
    button.Height = 20
    button.BackgroundColor = 0x4CA71B
    button.TextColor = 0xFFFFFF
    button.Name = "Submit"
    button.TabIndex = 99 
    button.Label = "Add"
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

    # Filter all allowed fields in dialog
    columns = self.__section.FieldNames()
    filteredControls = filter(lambda c: c.Model.Name in columns, formDialog.Controls)

    # Fetch array of values from dialog fields
    formValues = list(map(lambda c: self.__getFieldValue(c), filteredControls))

    # Create a copy of the Section's Model (with autovalue fields filtered out) and merge with form values 
    model = filter(lambda f: not f['autovalue'] , self.__section.Model)
    self.Data = list(map(lambda model, value: dict(model, value = value), model, formValues))
    
    # Reset all labels colors
    labels = list(filter(lambda e: e.Model.Name[-5:] == 'Label', formDialog.Controls))
    for label in labels:
      label.Model.TextColor = 0x000000
    
    # Validate fields and manage errors
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
    ignoreFields = ['Id', 'TS']
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

