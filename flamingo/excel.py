from pyrevit import clr
from pyrevit.coreutils.logger import get_logger

clr.AddReferenceByName(
    "Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
)
from Microsoft.Office.Interop import Excel
from System.Runtime.InteropServices import Marshal
from os import path

LOGGER = get_logger(__name__)


def OpenWorkbook(excelPath=None, createNew=True):
    """
    Check to see if excel is already running. Open if not.
    Get a list of currently open workbooks. Connect to
    the correct workbook if open, otherwise open the workbook
    specified in `excelPath`.

    args:
        excelPath (optional, str): location of excel file to load
        createNew (optional, boolean): if true, a new workbook will be opened

    returns:
        Microsoft.Office.Interop.Excel.Workbook
    """
    LOGGER.debug("OpenWorkbook({},{})".format(excelPath, createNew))
    try:
        excel = Marshal.GetActiveObject("Excel.Application")
    except:
        excel = Excel.ApplicationClass()

    excel.Visible = True
    excel.DisplayAlerts = False

    workbook = None
    excelPath = excelPath or ""
    for wb in excel.Workbooks:
        fullName = wb.Fullname or ""
        if fullName.lower() == excelPath.lower():
            workbook = wb
            LOGGER.debug("Found opened workbook")
            break
    if workbook is None:
        if excelPath and path.exists(excelPath):
            workbook = excel.Workbooks.Open(excelPath)
            LOGGER.debug("Opened workbook at {}".format(excelPath))
        elif createNew:
            workbook = excel.Workbooks.Add()
            LOGGER.debug("Created new workbook")
        else:
            workbook = False
            LOGGER.debug("No workbook found")
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

    LOGGER.debug(
        "GetWorksheetData({}, group={}, skip={})".format(worksheet.Name, group, skip)
    )
    if hasattr(worksheet, "UsedRange"):
        rowIndexes = [row.Row for row in worksheet.UsedRange.Rows]
    else:
        rowIndexes = [row.Row for row in worksheet.Rows]
    lastRow = rowIndexes[-1]
    groupSort = 0
    groupName = None
    sheetValues = {
        "rowCount": worksheet.UsedRange.Rows.Count,
        "columnCount": worksheet.UsedRange.Columns.Count,
    }
    if hasattr(worksheet, "UsedRange"):
        columnIndexes = [column.Column for column in worksheet.UsedRange.Columns]
    else:
        columnIndexes = [column.Column for column in worksheet.Columns]
    for i in range(skip + 1, lastRow + 1):
        if i % 100 == 0:
            LOGGER.debug("Processing row {}".format(i))
        metaValues = {
            "Sort Name": groupName,
            "Sort Number": groupSort,
            "Row Number": i,
        }
        rowValues = []
        sortRow = False
        cellValue = None
        for j in columnIndexes:
            cellValue = worksheet.Cells[i, j].Text
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


def SetWorksheetData(workseheet, data):
    pass
