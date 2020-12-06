# Imports
from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, FORMULA

# --------------------------------------------------------------------------------------------------------
# Class SHEET
# 

class Sheet:

    def __init__(self, name, documentContext):
        self.Instance = documentContext.Sheets[name]
        self.Name = name

    # Returns the CellRange reference to the "range name" arg
    def Range(self, range):
        return self.Instance.getCellRangeByName(range) 
    
    # Clear the contents of the cell range
    def Clear(self, range):
        self.Range(range).clearContents(VALUE + DATETIME + STRING + FORMULA)

    # # Last used row
    # def LastRow(self):
    #     Curs = self.Instance.createCursor()
    #     Curs.gotoEndOfUsedArea(True)
    #     return Curs.Rows.Count

    #
    # def buildColumnHeaders(self):
    #     Data = Sheet(This().Sheets.getByName(sheetName + 'Data'))
    #     headers = fetchColumnHeaders(sheet)
    #     headersRow = Cell('FirstRow', sheet).value()
    #     for col in headers:
    #         msgbox(col)
    #         # Cell([0, firstRow], CURRENT_SHEET).setValue(row)