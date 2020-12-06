from com.sun.star.sheet.CellFlags import VALUE, DATETIME, STRING, FORMULA

# --------------------------------------------------------------------------------------------------------
# Class SHEET

class Sheet:

    def __init__(self, name, documentContext):
        self.Instance = documentContext.Sheets[name]
        self.Name = name

    # Returns the CellRange object corresponding to the "range" arg
    def Range(self, range):
        return self.Instance.getCellRangeByName(range) 
    
    #
    def Clear(self, range):
        self.Range(range).clearContents(VALUE + DATETIME + STRING + FORMULA)  # or 5
        # return None

    # Write an error in the ErrorCell or clears it if None is passed as parameter
    def Error(self, msg = "---", cell = None):
        # Default cell name
        cell = cell or "Error"
        # Clear if None passed as parameter
        if not msg:
            self.Instance.getCellRangeByName(cell).String = ""
            return None
        self.Instance.getCellRangeByName(cell).String = msg
        return None

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