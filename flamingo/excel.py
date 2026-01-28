from pyrevit import clr, script
from pyrevit.coreutils.logger import get_logger

clr.AddReferenceByName(
    "Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
)
# Marshal2 is a helper class to get running COM objects provided by Curtis Wensley at McNeel
# https://discourse.mcneel.com/t/connect-to-excel-using-vb-component-in-rhino-8/173471/10
# get a path relative to this file
from Microsoft.Office.Interop import Excel
from System.Reflection import BindingFlags
from System import Activator, Type, Array, Object
from os import path

clr.AddReferenceToFileAndPath(
    "{}\excel_utils.dll".format(path.dirname(path.abspath(__file__)))
)
from Utils import Marshal2


LOGGER = get_logger("flamingo.excel")


def _com_setattr(obj, name, value):
    obj.GetType().InvokeMember(
        name, BindingFlags.SetProperty, None, obj, Array[Object]([value])
    )


def _com_getattr(obj, name, arg1=None, arg2=None):
    if arg1 is not None:
        if arg2 is not None:
            # For two arguments like Item[row, col]
            arg_array = Array[Object]([arg1, arg2])
        else:
            # For one argument like Item["Members"]
            arg_array = Array[Object]([arg1])
        return obj.GetType().InvokeMember(
            name, BindingFlags.GetProperty, None, obj, arg_array
        )
    else:
        # For simple properties
        return obj.GetType().InvokeMember(
            name, BindingFlags.GetProperty, None, obj, None
        )


def _com_invoke(obj, method, arg1=None, arg2=None, arg3=None):
    args = [arg for arg in [arg1, arg2, arg3] if arg is not None]
    arg_array = Array[Object](args) if args else None
    return obj.GetType().InvokeMember(
        method, BindingFlags.InvokeMethod, None, obj, arg_array
    )


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
        LOGGER.debug("Getting active excel instance with Marshal")
        excel = Marshal2.GetActiveObject("Excel.Application")
    except Exception as e:
        LOGGER.debug("No running Excel instance: {}".format(e))
        excel_type = Type.GetTypeFromProgID("Excel.Application")
        excel = Activator.CreateInstance(excel_type)

    _com_setattr(excel, "Visible", True)
    _com_setattr(excel, "DisplayAlerts", False)

    workbook = None
    excelPath = excelPath or ""
    workbooks = _com_getattr(excel, "Workbooks")
    for wb in workbooks:
        fullName = _com_getattr(wb, "FullName") or ""
        LOGGER.debug("Checking open workbook: {}".format(fullName))
        if fullName.lower() == excelPath.lower():
            workbook = wb
            LOGGER.debug("Found opened workbook")
            break
    if workbook is None:
        if excelPath and path.exists(excelPath):
            workbook = _com_invoke(workbooks, "Open", excelPath)
            LOGGER.debug("Opened workbook at {}".format(excelPath))
        elif createNew:
            workbook = _com_invoke(workbooks, "Add")
            LOGGER.debug("Created new workbook")
        else:
            workbook = False
            LOGGER.debug("No workbook found")
    return workbook


def GetWorksheetByName(workbook, sheetName):
    """
    Get a worksheet from the provided workbook by name.

    args:
        workbook(Microsoft.Office.Interop.Excel.Workbook)
        sheetName(str)

    returns:
        Microsoft.Office.Interop.Excel.Workbook.Worksheet
    """
    LOGGER.debug(
        "GetWorksheetByName({}, {})".format(
            _com_getattr(workbook, "FullName"), sheetName
        )
    )
    sheets = _com_getattr(workbook, "Worksheets")
    for sheet in sheets:
        name = _com_getattr(sheet, "Name")
        if name == sheetName:
            return sheet
    return None


def GetCellValue(worksheet, row, column):
    """
    Get the value of a specific cell in a worksheet.

    args:
        worksheet(Microsoft.Office.Interop.Excel.Workbook.Worksheet)
        row(int)
        column(int)

    returns:
        str
    """
    LOGGER.debug(
        "GetCellValue({}, {}, {})".format(
            _com_getattr(worksheet, "Name"), row, column
        )
    )
    cells = _com_getattr(worksheet, "Cells")
    cellValue = _com_getattr(cells, "Text", row, column)
    return cellValue

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
        "GetWorksheetData({}, group={}, skip={})".format(
            _com_getattr(worksheet, "Name"), group, skip
        )
    )
    usedRange = _com_getattr(worksheet, "UsedRange")
    if usedRange:
        rows = _com_getattr(usedRange, "Rows")
        rowIndexes = [_com_getattr(row, "Row") for row in rows]
    else:
        rows = _com_getattr(worksheet, "Rows")
        rowIndexes = [_com_getattr(row, "Row") for row in rows]
    lastRow = rowIndexes[-1]
    groupSort = 0
    groupName = None
    
    if usedRange:
        columns = _com_getattr(usedRange, "Columns")
        columnIndexes = [_com_getattr(column, "Column") for column in columns]
    else:
        columns = _com_getattr(worksheet, "Columns")
        columnIndexes = [_com_getattr(column, "Column") for column in columns]
    sheetValues = {
        "rowCount": _com_getattr(rows, "Count"),
        "columnCount": _com_getattr(columns, "Count"),
    }
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
        cells = _com_getattr(worksheet, "Cells")
        for j in columnIndexes:
            cell = _com_getattr(cells, "Item", arg1=i, arg2=j)
            cellValue = _com_getattr(cell, "Text")
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


def CloseWorkbook(workbook):
    """
    Close the provided workbook.

    args:
        workbook(Microsoft.Office.Interop.Excel.Workbook)
    """
    LOGGER.debug(
        "CloseWorkbook({})".format(_com_getattr(workbook, "FullName"))
    )
    _com_invoke(workbook, "Close", False)