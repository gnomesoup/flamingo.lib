from pyrevit import clr
clr.AddReferenceByName(
    "Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
)
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal
from os import path

def OpenWorkbook(excelPath, createNew = True):
    """
    Check to see if excel is already running. Open if not.
    Get a list of currently open workbooks. Connect to
    the correct workbook if open, otherwise open the workbook
    specified in `excelPath`.

    args:
        excelPath (str): location of excel file to load
        createNew (boolean): if true, a new workbook will be opened
    
    returns:
        Microsoft.Office.Interop.Excel.Workbook
    """
    try:
        excel = Marshal.GetActiveObject("Excel.Application")
    except:
        excel = Excel.ApplicationClass()

    excel.Visible = True
    excel.DisplayAlerts = False

    workbook = None
    for wb in excel.Workbooks:
        if (wb.Fullname).lower() == excelPath.lower():
            workbook = wb
            break
    if workbook is None:
        if path.exists(excelPath):
            workbook = excel.Workbooks.Open(excelPath)
        elif createNew:
            workbook = excel.Workbooks.Add()
        else:
            workbook = False
    return workbook

def GetWorksheetData(worksheet, group=False, skip=0):
    """
    Grab data from an excel worksheet into a dictionary with row
    numbers as keys. If `group` is set to `True`, then any rows that
    have only the first column filled out will get added as sorting
    value added to the first element of the following row dictionaries
    until another row with only a single element in the first column
    is reached.

    args:
        worksheet(Microsoft.Office.Interop.Excel.Workbook.Worksheet)
        [group(boolean)] Defaults to False
        [skip(int)] Defaults to 0
    returns:
        dict
    """

    lastRow = worksheet.UsedRange.Rows.Count
    lastColumn = worksheet.UsedRange.Columns.Count
    groupSort = 0
    groupName = None
    sheetValues = {
        "rowCount": lastRow,
        "columnCount": lastColumn
    }
    for i in range(skip + 1, lastRow + 1):
        metaValues = {
            "Sort Name": groupName,
            "Sort Number": groupSort,
            "Row Number": i,
        }
        rowValues = []
        sortRow = False
        cellValue = None
        for j in range(1, lastColumn + 1):
            cellValue = worksheet.Cells[i,j].Text
            if cellValue not in [None, ""]:
                if j == 1 and group:
                    sortRow = True
                else:
                    sortRow = False
            rowValues.append(cellValue)
        if sortRow:
            groupSort = groupSort + 1
            groupName = rowValues[0]
            continue
        sheetValues[i] = {
            "meta": metaValues,
            "data": rowValues,
        }
    return sheetValues
