from pyrevit import DB, HOST_APP, forms
from os import path
import re

def GetParameterFromProjectInfo(doc, parameterName):
    """
    Returns a parameter value from the Project Information category by name.
    """
    projectInformation = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ProjectInformation) \
        .ToElements()
    parameterValue = projectInformation[0].LookupParameter(parameterName)\
        .AsString()
    return parameterValue

def SetNoteBlockProperties(scheduleView,
                           viewName,
                           familyInstance,
                           metaParameterNameList,
                           columnNameList,
                           headerNames=None,
                           columnWidthList=None,
                           columnFormatList=None):
    """
    Set the name, fields, and filter of a schedule view to make is show the
    data that was assigned to the generic annoation family

    Parameters
    ----------
    scheduleView : DB.ViewSchedule
        Note block schedule to edit
    viewName :str
        Name to assign schedule
    symbol : DB.FamilyInstance
        Instance of family assigned to note block
    metaParameterNameList : list(str)
        List of parameter names used to sort/group schedule
    columnNameList : list(str)
        List of parameter names that represent excel columns
    headerNames : list(str)
        List of names to assign column headers
    columnWidthList : list(float)
        List of column widths in feet
    columnFormatList: Reserved
    
    Returns
    -------
    DB.ViewSchedule
        Modified note block schedule view
    """
    parameterNameList = metaParameterNameList + columnNameList
    if headerNames:
        headerNameList = metaParameterNameList + headerNames
    if columnWidthList:
        columnWidthList = ([1/12.0] * len(metaParameterNameList))\
            + columnWidthList
    scheduleView.Name = viewName
    scheduleDefinition = scheduleView.Definition
    scheduleDefinition.ClearFields()
    fields = {}
    for i in range(len(parameterNameList)):
        parameter = familyInstance.LookupParameter(parameterNameList[i])
        if parameter:
            schedulableField = DB.SchedulableField(
                DB.ScheduleFieldType.Instance,
                parameter.Id)
            newField = scheduleDefinition.AddField(schedulableField)
            if headerNames:
                if headerNameList[i]:
                    newField.ColumnHeading = headerNameList[i]
            if parameterNameList[i] in metaParameterNameList:
                newField.IsHidden = True
            if columnWidthList:
                newField.GridColumnWidth = columnWidthList[i]
            fields[parameterNameList[i]] = newField
    return scheduleView

def GetScheduleFields(viewSchedule):
    """
    Creates a dictionary of fields assigned to a schedule by field name

    Args:
        viewSchedule (DB.ViewSchedule): The schedule view to pull fields from

    Returns:
        dict: Dictionary of fields in the schedule added by field name
    """
    fields = {}
    scheduleDefinition = viewSchedule.Definition
    for scheduleId in scheduleDefinition.GetFieldOrder():
        field = scheduleDefinition.GetField(scheduleId)
        fields[field.GetName()] = field
    return fields

def MarkModelTransmitted(filepath, isTransmitted=True):
    """
    Marks a workshared project as transmitted. The model will be forced to
    open detached the next time it is opened. The model must not be opened
    or an error will be thrown.

    Parameters
    ----
    filepath: str
        Full path to a valid workshared revit model
    isTransmitted: bool
        Sets if the model will be marked as transmitted (default True)
    """

    modelPath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath)
    transmissionData = DB.TransmissionData.ReadTransmissionData(modelPath)
    transmissionData.IsTransmitted = isTransmitted
    transmissionData.WriteTransmissionData(modelPath, transmissionData)
    return

def GetModelFilePath(doc, promptIfBlank=True):
    if doc.IsWorkshared:
        revitModelPath = doc.GetWorksharingCentralModelPath()
        revitModelPath = DB.ModelPathUtils\
            .ConvertModelPathToUserVisiblePath(revitModelPath)
    else:
        revitModelPath = doc.PathName
    if not revitModelPath and promptIfBlank:
        forms.alert("Please save the model and try again.",
            exitscript=True)
    uncReg = re.compile(re.escape("\\\\wha-server02\\projects"),
        re.IGNORECASE)
    return uncReg.sub("x:", revitModelPath)

def GetModelDirectory(doc):
    return path.dirname(GetModelFilePath(doc))

def GetPhase(phaseName, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    phases = doc.Phases
    for phase in phases:
        if phase.Name == phaseName:
            return phase
    return None

def OpenDetached(filePath, audit=False, preserveWorksets=True):
    modelPath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(
        filePath
    )
    openOptions = DB.OpenOptions()
    if preserveWorksets:
        detachOption = DB.DetachFromCentralOption.DetachAndDiscardWorksets
    else:
        detachOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
    openOptions.DetachFromCentralOption = detachOption
    openOptions.Audit = True
    doc = HOST_APP.app.OpenDocumentFile(modelPath, openOptions)
    return doc