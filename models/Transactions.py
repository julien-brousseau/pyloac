# Model for transaction sheet columns and form

# FieldName:  (Str)   Form and column name
  # col:      (Int)   Column index
  # type':    (Str)   'Value' | 'String' | 'Date' | 'Checkbox'
  # reqd':    (Bool)  Form field value is required
  # label':   (Str)   Column header,
  # default': (Str)   'Default value' | 'CLEAR' | 'PREVIOUS'
  # noform':  (Bool)  Optional - No associated form field (default False)
# }

Transactions = {
  'Date': {
    'col': 0,
    'type': 'Date',
    'reqd': 1,
    'label': 'Date',
    'default':  'PREVIOUS'
  },
  'Amount': {
      'col': 1,
      'type': 'Value',
      'reqd': 1,
      'label': 'Amount',
      'default': 'CLEAR'
  },
  'Type': {
      'col': 2,
      'type': 'String',
      'reqd': 1,
      'label': 'Catégorie',
      'default': 'PREVIOUS'
  },
  'Description': {
      'col': 3,
      'type': 'String',
      'reqd': 0,
      'label': 'Decription',
      'default': ''
  },
  'File': { 
      'col': 4,
      'type': 'String',
      'reqd': 1,
      'label': 'Fichier',
      'default': 'PREVIOUS'
  },
  'FromAccount': { 
      'col': 5,
      'type': 'String',
      'reqd': 0,
      'label': 'From account',
      'default': 'PREVIOUS'
  },
  'ToAccount': {
      'col': 6,
      'type': 'String',
      'reqd': 0,
      'label': 'Compte crd',
      'default': ''
  },
  'Deductible': {
      'col': 8,
      'type': 'CheckBox',
      'reqd': 0,
      'label': 'Déductible',
      'default': ''
  },
  'GigId': { 
      'col': 7,
      'type': 'String',
      'reqd': 0,
      'label': 'Gig ID',
      'default': ''
  },
  'ID': {
    'col': 20, 
    'type': 'Value',
    'reqd': 1,
    'label': 'ID',
    'default': 'CLEAR',
    'noform': 1
  },
  'Timestamp': {
    'col': 21,
    'type': 'String',
    'reqd': 1,
    'label': 'Timestamp',
    'default': 'CLEAR',
    'noform': 1
  },
}