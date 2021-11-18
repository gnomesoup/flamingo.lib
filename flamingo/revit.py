from pyrevit import DB, HOST_APP, forms, revit
from os import path, sep
from string import ascii_uppercase
from flamingo.geometry import GetMinMaxPoints, GetMidPoint
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
        detachOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
    else:
        detachOption = DB.DetachFromCentralOption.DetachAndDiscardWorksets
    openOptions.DetachFromCentralOption = detachOption
    openOptions.Audit = True
    doc = HOST_APP.app.OpenDocumentFile(modelPath, openOptions)
    return doc

def SaveAsCentral(doc, centralPath):
    saveAsOptions = DB.SaveAsOptions()
    saveAsOptions.OverwriteExistingFile = True
    workshareOptions = DB.WorksharingSaveAsOptions()
    workshareOptions.SaveAsCentral = True
    workshareOptions.OpenWorksetsDefault = DB.SimpleWorksetConfiguration.AllWorksets
    saveAsOptions.SetWorksharingOptions(workshareOptions)
    DB.Document.SaveAs(doc, centralPath, saveAsOptions)
    relinquishOptions = DB.RelinquishOptions(True)
    transactionOptions = DB.TransactWithCentralOptions()
    DB.WorksharingUtils.RelinquishOwnership(doc, relinquishOptions,
                                            transactionOptions)
    return doc

def DoorRenameByRoomNumber(
    doors,
    phase,
    prefix=None,
    seperator=None,
    suffixList=None,
    doc=None,
):
    """[summary]

    Args:
        doors (list[DB.Elements]): List of Revit doors to renumber.
        phase (DB.Phases): Phase to use when getting room property.
        prefix (str, optional): String to append to beinging of door number.
            Defaults to None.
        seperator (str, optional): String to place between door number and 
            suffix. Defaults to None.
        suffixList (list[str], optional): List of strings to append to the end
            of doors located in the same room. Defaults to and uppercase letters
            in alphabetical order.
        doc (DB.Document, optional): [description]. Defaults to None.
    
    Returns:
        list[DB.Elements]: list of renumbered doors
    """
    if doc is None:
        doc = HOST_APP.doc
    if prefix is None:
        prefix = ""
    if seperator is None:
        seperator = ""
    if suffixList is None:
        suffixList = ascii_uppercase

    # Make a dictionary of rooms with door properties
    roomDoors = {}
    for door in doors:
        if not door:
            continue

        toRoom = (door.ToRoom)[phase]
        if toRoom:
            toRoomId = toRoom
            if toRoom.Id in roomDoors:
                roomDoors[toRoom.Id]["doors"] = \
                    roomDoors[toRoom.Id]["doors"] + [door.Id]
            else:
                roomDoors[toRoom.Id] = {"doors": [door.Id]}
                roomDoors[toRoom.Id]["roomNumber"] = toRoom.Number
                # point1, point2 = GetMinMaxPoints(toRoom, activeView)
                # roomCenter = GetMidPoint(point1, point2)
                # roomDoors[toRoom.Id]["roomLocation"] = roomCenter

        fromRoom = (door.FromRoom)[phase]
        if fromRoom:
            if fromRoom.Id in roomDoors:
                roomDoors[fromRoom.Id]["doors"] =\
                    roomDoors[fromRoom.Id]["doors"] + [door.Id]
            else:
                roomDoors[fromRoom.Id] = {"doors": [door.Id]}
                roomDoors[fromRoom.Id]["roomNumber"] = fromRoom.Number
                # point1, point2 = GetMinMaxPoints(fromRoom, activeView)
                # roomCenter = GetMidPoint(point1, point2)
                # roomDoors[fromRoom.Id]["roomLocation"] = roomCenter

    maxLength = max([len(values["doors"]) for values in roomDoors.values()])
    for key, values in roomDoors.items():
        roomDoors[key]["doorCount"] = len(values["doors"])

    # Number doors
    doorCount=len(doors)
    numberedDoors = set()
    with revit.Transaction("Renumber Doors"):
        i = 1
        n = 0
        nMax = doorCount * 2

        # iterate through counts of doors connected to each room
        while i < maxLength:
            # avoid a loop by stoping at double the count of doors
            if n > nMax:
                break
            noRooms = True
            # Go through each room and look for door counts that match
            # current level
            for key, values in roomDoors.items():
                if values["doorCount"] == i:
                    noRooms = False
                    doorIds = values["doors"]
                    doorsToNumber = [
                        doorId for doorId in doorIds
                        if doorId not in numberedDoors
                    ]
                    # Go through all the doors connected to the room
                    # and give them a number
                    for j, doorId in enumerate(doorsToNumber):
                        numberedDoors.add(doorId)
                        door = doc.GetElement(doorId)
                        if len(doorsToNumber) > 1:
                            # point1, point2 = GetMinMaxPoints(door, activeView)
                            # roomCenter = values["roomLocation"]
                            # doorCenter = GetMidPoint(point1, point2)
                            # doorVector = doorCenter.Subtract(roomCenter)
                            # angle = (DB.XYZ.BasisY).AngleTo(
                            #     doorVector.Normalize()
                            # )
                            # dotProduct = (DB.XYZ.BasisY).DotProduct(
                            #     doorVector.Normalize()
                            # )
                            mark = "{}{}".format(
                                values["roomNumber"], ascii_uppercase[j]
                            )
                            # print("{}: {} ({})".format(mark, angle, dotProduct))
                        else:
                            mark = values["roomNumber"]
                        markParameter = door.get_Parameter(
                            DB.BuiltInParameter.ALL_MODEL_MARK
                        )
                        markParameter.Set(mark)
            for roomId, values in roomDoors.items():
                unNumberedDoors = filter(None, [
                    doorId for doorId in values["doors"]
                    if doorId not in numberedDoors
                ])
                roomDoors[roomId]["doors"] = unNumberedDoors
                roomDoors[roomId]["doorCount"] = len(unNumberedDoors)
            if noRooms:
                i += 1
            n += 1
    return doors

def FilterByCategory(elements, builtInCategory):
    """Filters a list of Revit elements by a BuiltInCategory

    Args:
        elements (list[DB.Elements]): List of revit elements to filter
        builtInCategory (DB.BuiltInCategory): Revit built in category enum
            to filter by
    Returns:
        list[DB.Elements]: List of revit elements that matched the filter
    """
    return [
        element for element in elements
        if element.Category.Id.IntegerValue == int(builtInCategory)
    ]