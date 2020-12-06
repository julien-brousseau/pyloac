# --------------------------------------------------------------------------------------------------------
# Class FIELD, which associates DATA objects with their respective form fields. Allows to validate their
# data and set field values

# import sys
# from ..models.Transactions import Transactions 

# MODELS = {
#     'Transactions': Transactions
# }

class Field:

    def __init__(self, fieldName, sheet):
        self.Name = "blop"

    #     self.Name = fieldName
    #     self.SheetName = sheet.Name

    #     # Field element
    #     self.Instance = sheet.Instance.DrawPage.getForms().getByIndex(0).getByName(self.Name)

    #     # Field type (ListBox, CheckBox, Edit, DateField)
    #     self.Type = self.Instance.ServiceName.split('.')[-1]

    #     # Data model
    #     self.Data = MODELS[self.SheetName][self.Name] if MODELS[self.SheetName].get(self.Name) else {}

    # # Rebuild the item list for listbox objects
    # def RefreshValues(self): 
 
    #     # Bypass if field is not listbox type
    #     if not self.Type == "ListBox":
    #         return False

    #     Items = []
    #     Column = None

    #     # Find the corresponding column in DATA
    #     for t in range(0,15):
    #         if self.Name == LISTS_SHEET.Instance.getCellByPosition(t,0).String:
    #             Column = t

    #     # Exit if no column matches field name 
    #     if Column is None:
    #         return False

    #     # Add empty value at beginning if the field is not required
    #     if not self.Data["reqd"]:
    #         Items.append("")

    #     # Add all string values below the column header
    #     for x in range(1, 50):
    #         a = LISTS_SHEET.Instance.getCellByPosition(Column, x).String
    #         if not a:
    #             break
    #         Items.append(a)

    #     # Replace the list values with Items
    #     self.Instance.StringItemList = Items

    # return None
