from Autodesk.Revit import DB
import codecs
from datetime import datetime
from flamingo.geometry import GetMidPoint, GetSolids, MakeSolid
from math import atan2, pi
from pyrevit import HOST_APP, forms, PyRevitException, revit, script
from pyrevit.coreutils.configparser import configparser
from os import path
import re
import shutil
from string import ascii_uppercase
from System import Guid
from System.Collections.Generic import List

LOGGER = script.get_logger()
OUTPUT = script.get_output()


class iFamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True

    def OnSharedFamilyFound(
        self, sharedFamily, familyInUse, source, overwriteParameterValues
    ):
        overwriteParameterValues = True
        source = DB.FamilySource.Family
        return True


def CreateProjectParameter(
    parameterName,
    sharedParameterGroupName,
    revitCategories,
    parameterGroup,
    sharedParametersFilename=None,
    isTypeParameter=True,
    doc=None,
):
    """Import a project parameter from the specified shared parameter file

    Args:
        parameterName (str|System.Guid): Name or Guid of the parameter to import
        sharedParameterGroupName (str): Group the shared parameter is organized within
        revitCategories (list(DB.Categories)): Category the parameter should be added to
        parameterGroup (DB.): Parameter group enum
        sharedParametersFilename (str, optional): Path to the shared parameter file that
            has the parameter. Defaults to None.
        isTypeParameter (bool, optional): True if type parameter. False if instance
            parameter. Defaults to True.
        doc (DB.Document, optional): Document to load parameter. Defaults to None.

    Returns:
        DB.ParameterBinding: Revit parameter binding for added parameter
    """
    doc = doc or HOST_APP.doc
    app = doc.Application

    # Add the parameter to the model
    # if the current shared parameters file is not the provide file, change it
    originalFilename = app.SharedParametersFilename
    sharedParametersFilename = sharedParametersFilename or originalFilename
    if originalFilename != sharedParametersFilename:
        if path.exists(sharedParametersFilename):
            app.SharedParametersFilename = sharedParametersFilename
        else:
            PyRevitException(
                "Could not located specified shared parameter file at {}".format(
                    sharedParametersFilename
                )
            )

    # Find the parameter in the parameters file
    definitionsFile = app.OpenSharedParameterFile()
    if not definitionsFile:
        raise PyRevitException("Could not read from the shared parameters file")
    definitionGroup = definitionsFile.Groups.get_Item(sharedParameterGroupName)
    if type(parameterName) == Guid:
        externalDefinition = None
        for definition in definitionGroup.Definitions:
            if definition.GUID == parameterName:
                externalDefinition = definition
                break
    else:
        externalDefinition = definitionGroup.Definitions.get_Item(parameterName)
    if not externalDefinition:
        LOGGER.debug("No externalDefinition found")
        raise PyRevitException(
            "Could not locate parameter in shared parameter file: {}".format(
                str(parameterName)
            )
        )

    # Format Revit categories
    categories = doc.Settings.Categories
    categorySet = app.Create.NewCategorySet()
    for revitCategory in revitCategories:
        category = categories.get_Item(revitCategory)
        categorySet.Insert(category)

    # If the document is a family we can end here
    if doc.IsFamilyDocument:
        return doc.FamilyManager.AddParameter(
            externalDefinition, parameterGroup, not isTypeParameter
        )

    # Create the parameter
    if isTypeParameter:
        newBinding = app.Create.NewTypeBinding(categorySet)
    else:
        newBinding = app.Create.NewInstanceBinding(categorySet)
    LOGGER.debug("externalDefinition = {}".format(externalDefinition))
    LOGGER.debug("newBinding = {}".format(newBinding))
    LOGGER.debug("parameterGroup = {}".format(parameterGroup))
    return doc.ParameterBindings.Insert(externalDefinition, newBinding, parameterGroup)


def CreateGlobalParameter(doc, name, param_type, value=None):
    """
    Create a new Global Parameter in the Revit document

    Args:
        doc: The Revit document
        name: Name of the global parameter
        param_type: Parameter type (e.g., DB.ParameterType.Length)
        value: Value to assign to the parameter (optional)

    Returns:
        The newly created GlobalParameter
    """
    # Check if parameter already exists
    existing_param_ids = DB.GlobalParametersManager.GetAllGlobalParameters(doc)
    for param_id in existing_param_ids:
        param = doc.GetElement(param_id)
        if param.Name == name:
            LOGGER.debug("Global parameter '{}' already exists".format(name))
            return param

    # Create a new global parameter
    param_id = DB.GlobalParametersManager.CreateGlobalParameter(doc, name, param_type)
    param = doc.GetElement(param_id)

    # Set the parameter value if provided
    if value is not None:
        if param_type == DB.ParameterType.Length:
            value_wrapper = DB.DoubleParameterValue(value)
        elif param_type == DB.ParameterType.Number:
            value_wrapper = DB.DoubleParameterValue(value)
        elif param_type == DB.ParameterType.Integer:
            value_wrapper = DB.IntegerParameterValue(value)
        elif param_type == DB.ParameterType.Text:
            value_wrapper = DB.StringParameterValue(value)
        elif param_type == DB.ParameterType.YesNo:
            value_wrapper = DB.IntegerParameterValue(1 if value else 0)
        else:
            LOGGER.debug("Unsupported parameter type: {}".format(param_type))

        param.SetValue(value_wrapper)


def GetParameterFromProjectInfo(doc, parameterName, defaultValue=None):
    """
    Returns a parameter value from the Project Information category by name.
    """
    LOGGER.debug("GetParameterFromProjectInfo: {}".format(parameterName))
    parameterValue = None
    try:
        parameterValue = GetParameterValueByName(doc.ProjectInformation, parameterName)
    except Exception as e:
        LOGGER.debug(e)
        return defaultValue
    LOGGER.debug("parameterValue: {}".format(parameterValue))
    return parameterValue or defaultValue


def SetParameterFromProjectInfo(doc, parameterName, parameterValue):
    """
    Set a parameter value from the Project Information category by name.
    """
    try:
        SetParameter(
            doc.ProjectInformation, parameterName, parameterValue
        )
        parameter = doc.ProjectInformation.LookupParameter(parameterName)
        parameter.Set(parameterValue)
    except:
        return None
    return parameter


def SetNoteBlockProperties(
    scheduleView,
    viewName,
    metaParameterNameList,
    columnNameList,
    headerNames=None,
    columnWidthList=None,
    columnFormatList=None,
):
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
        columnWidthList = ([1 / 12.0] * len(metaParameterNameList)) + columnWidthList
    scheduleView.Name = viewName
    scheduleDefinition = scheduleView.Definition
    scheduleDefinition.ClearFields()
    schedulableFields = GetSchedulableFields(scheduleView)
    fields = {}
    for i in range(len(parameterNameList)):
        parameterName = parameterNameList[i]
        if parameterName in schedulableFields:
            schedulableField = schedulableFields[parameterName]
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


def GetSchedulableFields(viewSchedule):
    """Create a dictionary of schedulable fields for the provided schedule view.

    Args:
        viewSchedule (DB.ViewSchedule): Revit schedule view

    Returns:
        dict: Dictionary of fields that can be scheduled, keyed by field name
    """
    fields = {}
    scheduleDefinition = viewSchedule.Definition
    for field in scheduleDefinition.GetSchedulableFields():
        parameterId = field.ParameterId
        if parameterId is not None:
            parameter = viewSchedule.Document.GetElement(parameterId)
            if parameter:
                fields[parameter.Name] = field
    return fields


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


def PurgeUnused(doc=None):
    doc = doc or HOST_APP.doc
    purgeGuid = "e8c63650-70b7-435a-9010-ec97660c1bda"
    purgableElementIds = []
    performanceAdviser = DB.PerformanceAdviser.GetPerformanceAdviser()
    guid = Guid(purgeGuid)
    ruleId = None
    allRuleIds = performanceAdviser.GetAllRuleIds()
    for rule in allRuleIds:
        # Finds the PerformanceAdviserRuleId for the purge command
        if str(rule.Guid) == purgeGuid:
            ruleId = rule
    ruleIds = List[DB.PerformanceAdviserRuleId]([ruleId])
    for i in range(4):
        # Executes the purge
        failureMessages = performanceAdviser.ExecuteRules(doc, ruleIds)
        if failureMessages.Count > 0:
            # Retreives the elements
            purgableElementIds = failureMessages[0].GetFailingElements()
    # Deletes the elements
    try:
        doc.Delete(purgableElementIds)
    except:
        for e in purgableElementIds:
            try:
                doc.Delete(e)
            except:
                pass


def GetElementMaterialIds(element):
    LOGGER.debug("GetElementMaterialIds: element={}".format(OUTPUT.linkify(element.Id)))
    try:
        elementMaterials = element.GetMaterialIds(False)
        paintMaterials = element.GetMaterialIds(True)
        for materialId in paintMaterials:
            elementMaterials.append(materialId)
        return elementMaterials
    except Exception as e:
        LOGGER.debug(e)
        return []


def GetMaterialDictionary(doc):
    doc = doc or HOST_APP.doc
    materials = DB.FilteredElementCollector(doc).OfClass(DB.Material).ToElements()
    return {material.Name: material.Id for material in materials}


def GetAllProjectMaterialIds(inUseOnly=None, doc=None):
    doc = doc or HOST_APP.doc
    inUseOnly = inUseOnly if inUseOnly is not None else True
    if not inUseOnly:
        return DB.FilteredElementCollector(doc).OfClass(DB.Material).ToElementIds()
    elements = (
        DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    )
    projectMaterialIds = set()
    for element in elements:
        for materialElementId in GetElementMaterialIds(element):
            projectMaterialIds.add(materialElementId)
    return list(projectMaterialIds)


def GetUnusedMaterials(doc=None, nameFilter=None):
    doc = doc or HOST_APP.doc
    allMaterialIds = GetAllProjectMaterialIds(inUseOnly=False, doc=doc)
    usedMaterialIds = GetAllProjectMaterialIds(inUseOnly=True, doc=doc)
    usedMaterials = [
        doc.GetElement(materialId)
        for materialId in allMaterialIds
        if materialId not in usedMaterialIds
    ]
    if nameFilter:
        nameFilterRegex = re.compile("|".join(nameFilter))
        return [
            usedMaterial
            for usedMaterial in usedMaterials
            if not nameFilterRegex.match(usedMaterial.Name)
        ]
    else:
        return usedMaterials


def GetUnusedAssets(doc=None):
    postPurgeMaterials = (
        DB.FilteredElementCollector(doc).OfClass(DB.Material).ToElements()
    )

    currentAssetIds = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.AppearanceAssetElement)
        .ToElementIds()
    )

    usedAssetIds = [material.AppearanceAssetId for material in postPurgeMaterials]
    return [
        doc.GetElement(elementId)
        for elementId in currentAssetIds
        if elementId not in usedAssetIds
        if elementId is not None
    ]


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
        revitModelPath = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(
            revitModelPath
        )
    else:
        revitModelPath = doc.PathName
    if not revitModelPath and promptIfBlank:
        forms.alert("Please save the model and try again.", exitscript=True)
    uncReg = re.compile(re.escape("\\\\wha-server02\\projects"), re.IGNORECASE)
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


def OpenDetached(filePath, audit=False, preserveWorksets=True, visible=False):
    modelPath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath)
    openOptions = DB.OpenOptions()
    if preserveWorksets:
        detachOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
    else:
        detachOption = DB.DetachFromCentralOption.DetachAndDiscardWorksets
    openOptions.DetachFromCentralOption = detachOption
    openOptions.Audit = audit
    if visible:
        doc = HOST_APP.uiapp.OpenAndActivateDocument(modelPath, openOptions, True)
    else:
        doc = HOST_APP.app.OpenDocumentFile(modelPath, openOptions)
    return doc


def OpenLocal(filePath, localDir, extension="", hostApp=None):
    hostApp = hostApp or HOST_APP
    LOGGER.debug('Flamingo "OpenLocal" called')
    localDir = path.expandvars(localDir)
    centralFileName = path.basename(filePath)
    centralName = re.sub(r"\.rvt$", "", centralFileName)
    localFileName = "{}_{}{}.rvt".format(centralName, HOST_APP.username, extension)
    localPath = path.join(localDir, localFileName)
    LOGGER.info("filePath={}".format(filePath))
    LOGGER.info("localPath={}".format(localPath))
    shutil.copyfile(filePath, localPath)
    return hostApp.uiapp.OpenAndActivateDocument(localPath)


def GetDefaultPathForUserFiles(app=None):
    app = HOST_APP.doc.Application
    currentUsersDataFolderPath = app.CurrentUsersDataFolderPath
    revitIniPath = "{}/Revit.ini".format(currentUsersDataFolderPath)
    if path.exists(revitIniPath):
        cp = configparser.ConfigParser()
        with codecs.open(revitIniPath, mode="r", encoding="UTF-16") as f:
            cp.readfp(f)
        return cp.get("Directories", "projectpath")
    return None


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
    DB.WorksharingUtils.RelinquishOwnership(doc, relinquishOptions, transactionOptions)
    return doc


def DoorRenameByRoomNumber(
    doors,
    phase,
    prefix=None,
    seperator=None,
    suffixList=None,
    doc=None,
):
    """Utility to automatically assign door numbers based on room numbers.

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

    allDoors = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Doors)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    selectedDoorIds = [door.Id for door in doors]

    # Make a dictionary of rooms with door properties
    doorsByRoom = {}
    for door in allDoors:
        if not door:
            continue
        roomsToAdd = []
        try:
            toRoom = (door.ToRoom)[phase]
        except Exception as e:
            continue
        if toRoom:
            if toRoom.Id in doorsByRoom:
                doorsByRoom[toRoom.Id]["doors"] = doorsByRoom[toRoom.Id]["doors"] + [
                    door.Id
                ]
            else:
                roomsToAdd.append(toRoom)
        fromRoom = (door.FromRoom)[phase]
        if fromRoom:
            if fromRoom.Id in doorsByRoom:
                doorsByRoom[fromRoom.Id]["doors"] = doorsByRoom[fromRoom.Id][
                    "doors"
                ] + [door.Id]
            else:
                roomsToAdd.append(fromRoom)
        for roomToAdd in roomsToAdd:
            roomCenter = roomToAdd.Location.Point
            doorsByRoom[roomToAdd.Id] = {
                "doors": [door.Id],
                "roomNumber": roomToAdd.Number,
                "roomArea": roomToAdd.Area,
                "roomLocation": roomCenter,
            }

    # Make a dictionary of door connection counts. This will be used to
    # prioritize doors with more connections when assigning a letter value
    doorsByRoomDoors = [value["doors"] for value in doorsByRoom.values()]
    doorConnectors = {}
    for door in allDoors:
        for i in doorsByRoomDoors:
            if door.Id in i:
                if door.Id in doorConnectors:
                    doorConnectors[door.Id] += len(i)
                else:
                    doorConnectors[door.Id] = len(i)

    maxLength = max([len(values["doors"]) for values in doorsByRoom.values()])
    for key, values in doorsByRoom.items():
        doorsByRoom[key]["doorCount"] = len(values["doors"])

    # Number doors
    doorCount = len(doors)
    numberedDoors = set()
    with revit.Transaction("Renumber Doors"):
        i = 1
        n = 0
        nMax = doorCount * 2

        # iterate through counts of doors connected to each room
        while i <= maxLength:
            # avoid a loop by stoping at double the count of doors
            if n > nMax:
                # if n > 4:
                break
            noRooms = True
            roomsThisRound = [
                roomId
                for roomId, value in doorsByRoom.items()
                if value["doorCount"] == i
            ]
            roomsThisRound = sorted(
                roomsThisRound, key=lambda x: doorsByRoom[x]["roomArea"]
            )
            # Go through each room and look for door counts that match
            # current level
            # for key, values in doorsByRoom.items():
            for roomId in roomsThisRound:
                values = doorsByRoom[roomId]
                if values["doorCount"] == i:
                    noRooms = False
                    doorIds = values["doors"]
                    doorsToNumber = [
                        doorId for doorId in doorIds if doorId not in numberedDoors
                    ]
                    # Go through all the doors connected to the room
                    # and give them a number
                    sortedDoorsToNumber = sorted(
                        doorsToNumber, key=lambda x: doorConnectors[x], reverse=True
                    )
                    for j, doorId in enumerate(sortedDoorsToNumber):
                        numberedDoors.add(doorId)
                        door = doc.GetElement(doorId)
                        if len(doorsToNumber) > 1:
                            boundingBox = door.get_BoundingBox(None)
                            roomCenter = values["roomLocation"]
                            doorCenter = GetMidPoint(boundingBox.Min, boundingBox.Max)
                            doorVector = doorCenter.Subtract(roomCenter)
                            angle = atan2(doorVector.Y, doorVector.X)
                            angleReversed = angle * -1.0
                            pi2 = pi * 2.0
                            angleNormalized = ((angleReversed + pi2 + 2) % pi2) / pi2
                            mark = "{}{}".format(
                                values["roomNumber"], ascii_uppercase[j]
                            )
                            # print("{}: {} {} rad | {} ratio".format(mark, doorVector, angle, angleNormalized))
                        else:
                            mark = values["roomNumber"]
                        if door.Id in selectedDoorIds:
                            markParameter = door.get_Parameter(
                                DB.BuiltInParameter.ALL_MODEL_MARK
                            )
                            markParameter.Set(mark)
            for roomId, values in doorsByRoom.items():
                unNumberedDoors = filter(
                    None,
                    [
                        doorId
                        for doorId in values["doors"]
                        if doorId not in numberedDoors
                    ],
                )
                doorsByRoom[roomId]["doors"] = unNumberedDoors
                doorsByRoom[roomId]["doorCount"] = len(unNumberedDoors)
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
        element
        for element in elements
        if element.Category.Id.IntegerValue == int(builtInCategory)
    ]


def HideUnplacedViewTags(view=None, doc=None):
    """Hides all unreferenced view tags in the specified view by going through
    all elevation views, elevation tags, sections, and callouts in the view and
    identifying views that are not placed on a sheet and hides the associated
    tag. Elevation tags who have all hosted elevation view hidden by element or
    by a view filter will also be hidden.

    Args:
        view (Autodesk.Revit.DB.View, optional): View who's view tags are to be
            hidden. Defaults to None.
        doc (Autodesk.Revit.DB.Document, optional): Revit document that houses
            the view. Defaults to None.
    """
    # TODO: If current view is a sheet, hide on all non-drafting and non-3D views
    # TODO: Provide list of sheet sets to run the script on

    if doc is None:
        doc = HOST_APP.doc
    if view is None:
        view = doc.ActiveView

    viewers = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_Viewers)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    viewerIds = [viewer.Id.IntegerValue for viewer in viewers]
    elevs = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_Elev)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    # Go through all filtered elements to figure out if they have a sheet number
    # If they don't hide them by element
    hideList = List[DB.ElementId]()
    hideCount = 0
    for element in viewers:
        sheetNumberParam = element.get_Parameter(
            DB.BuiltInParameter.VIEWER_SHEET_NUMBER
        )
        if sheetNumberParam and sheetNumberParam.AsString() == "---":
            hideList.Add(element.Id)
            hideCount = hideCount + 1
        elif sheetNumberParam and sheetNumberParam.AsString() == "-":
            hideList.Add(element.Id)

    elementFilter = DB.ElementCategoryFilter(DB.BuiltInCategory.OST_Viewers)
    for element in elevs:
        sheetNumberParam = element.get_Parameter(
            DB.BuiltInParameter.VIEWER_SHEET_NUMBER
        )
        if sheetNumberParam and sheetNumberParam.AsString() == "-":
            hideList.Add(element.Id)
            continue
        dependentElementIds = element.GetDependentElements(elementFilter)
        shownViews = len(dependentElementIds)
        for dependentElementId in dependentElementIds:
            if dependentElementId.IntegerValue not in viewerIds:
                shownViews = shownViews - 1
        if shownViews < 1:
            hideList.Add(element.Id)
    if len(hideList) > 0:
        try:
            with revit.Transaction("Hide unplaced views"):
                view.HideElements(hideList)
        except Exception as e:
            print(e)


def UnhideViewTags(view=None, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    if view is None:
        view = doc.ActiveView
    categoryList = (DB.BuiltInCategory.OST_Viewers, DB.BuiltInCategory.OST_Elev)
    categoriesTyped = List[DB.BuiltInCategory](categoryList)
    categoryFilter = DB.ElementMulticategoryFilter(categoriesTyped)
    viewElements = (
        DB.FilteredElementCollector(doc)
        .WhereElementIsNotElementType()
        .WherePasses(categoryFilter)
        .ToElements()
    )

    hiddenElements = List[DB.ElementId]()
    [hiddenElements.Add(e.Id) for e in viewElements if e.IsHidden(view)]
    with revit.Transaction("Unhide view tags"):
        view.UnhideElements(hiddenElements)


def GetViewPhase(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    viewPhaseParameter = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
    viewPhaseId = viewPhaseParameter.AsElementId()
    return doc.GetElement(viewPhaseId)


def GetViewPhaseFilter(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    viewPhaseFilterParameter = view.get_Parameter(DB.BuiltInParameter.VIEW_PHASE_FILTER)
    viewPhaseFilterId = viewPhaseFilterParameter.AsElementId()
    return doc.GetElement(viewPhaseFilterId)


def _TestRoomIntersect(room, solid):
    from Autodesk.Revit.Exceptions import InvalidOperationException

    LOGGER.debug("room.Number = {}".format(room.Number))
    roomSolids = GetSolids(room)
    try:
        interSolid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
            roomSolids[0], solid, DB.BooleanOperationsType.Intersect
        )
        LOGGER.debug("interSolid.Volume = {}".format(interSolid.Volume))
        if hasattr(interSolid, "Volume") and abs(interSolid.Volume > 0.000001):
            return room, interSolid.Volume
    except InvalidOperationException as e:
        LOGGER.debug("Error: {}".format(str(e)))
    return None, None


def _ProjectToLevel(elementOutline, element):
    LOGGER.debug("Projecting to level")
    doc = element.Document
    location = element.Location
    if type(location) == DB.LocationCurve:
        curve = location.Curve
        elementPoint = curve.GetEndPoint(-1)
    elif hasattr(location, "Point"):
        elementPoint = location.Point
    else:
        elementPoint = None
    LOGGER.debug("elementPoint = {}".format(elementPoint))
    if elementPoint:
        if hasattr(element, "LevelId"):
            elementLevel = doc.GetElement(element.LevelId)
            if elementLevel:
                z = elementLevel.Elevation
        else:
            z = elementPoint.Z
            elementOutline.AddPoint(DB.XYZ(elementPoint.X, elementPoint.Y, z))
    return elementOutline


def GetElementRoomsFromLink(
    element,
    linkDoc,
    transform,
    roomsFromLink=None,
    offset=1.0,
    projectToLevel=True,
):
    LOGGER.info(
        "GetElementRoomsFromLink: element={}".format(OUTPUT.linkify(element.Id))
    )
    elementBoundingBox = element.get_BoundingBox(None)
    if not elementBoundingBox.Enabled:
        LOGGER.debug("Bounding Box Not Enabled")
        return

    elementOutline = DB.Outline(
        elementBoundingBox.Min.Add(DB.XYZ(-1.0 * offset, -1.0 * offset, -1.0 * offset)),
        elementBoundingBox.Max.Add(DB.XYZ(offset, offset, offset)),
    )

    if projectToLevel and element.Location is not None:
        elementOutline = _ProjectToLevel(elementOutline, element)

    # Transform the outline to the link
    inverseTransform = transform.Inverse
    transformedMinPoint = inverseTransform.OfPoint(elementOutline.MinimumPoint)
    transformedMaxPoint = inverseTransform.OfPoint(elementOutline.MaximumPoint)

    transformedOutline = DB.Outline(transformedMinPoint, transformedMaxPoint)
    boundingBoxIntersectsFilter = DB.BoundingBoxIntersectsFilter(transformedOutline)

    if roomsFromLink:
        rooms = (
            DB.FilteredElementCollector(linkDoc)
            .WherePasses(
                DB.ElementIdSetFilter(
                    List[DB.ElementId](room.Id for room in roomsFromLink)
                )
            )
            .WherePasses(boundingBoxIntersectsFilter)
            .ToElements()
        )
    else:
        rooms = (
            DB.FilteredElementCollector(linkDoc)
            .OfCategory(DB.BuiltInCategory.OST_Rooms)
            .WherePasses(boundingBoxIntersectsFilter)
            .ToElements()
        )
    LOGGER.debug("Number of intersecting rooms: {}".format(len(rooms)))
    elementSolids = GetSolids(element)
    matchedRooms = _GetMatchingRooms(element, elementSolids, elementOutline, rooms)
    return matchedRooms


def GetElementRooms(
    element, phase, rooms=None, offset=1.0, projectToLevel=True, doc=None
):
    """Find the room in a Revit model that a element is placed in or near.
    The function should find the room if it is located above or within a
    specified offset

    Args:
        element (DB.FamilyInstance): Element who's room is to be located
        phase (DB.Phase): The phase of the room that is to be
            matched to the element.
        rooms (list(DB.SpatialElement)): List of rooms to check
        offset (float, optional): Offset to provide the the room geometry. This
            allows for fuzzy matching of rooms that are close to the element
            but that the element is not directly located in. Defaults to 1.0.
        projectToLevel (bool, optional): If set to true, the boundary of the
            element will be expanded down to the level the object is associated.
            This is helpful for elements that may sit above the 3D extent of a
            room like a diffuser.
        doc (DB.Document, optional): The document in which the element is
            located. Defaults to None.

    Returns:
        DB.Room: Room in the Revit model to which the element is matched.
    """
    from Autodesk.Revit.Exceptions import InvalidOperationException

    LOGGER.debug(
        "GetElementRoom: element={}, phase={}, len(rooms)={}, offset={}, "
        "projectToLevel={}".format(
            OUTPUT.linkify(element.Id), phase.Id, len(rooms), offset, projectToLevel
        )
    )
    if doc is None:
        doc = element.Document

    # Check if element self reports room
    if hasattr(element, "ToRoom"):
        toRoom = [(element.ToRoom)[phase]]
        fromRoom = (element.FromRoom)[phase]
        if fromRoom:
            toRoom.append(fromRoom)
        if all(toRoom):
            LOGGER.debug("Element self reports room with ToRoom property")
            return toRoom
    if hasattr(element, "Room"):
        elementRoom = (element.Room)[phase]
        if elementRoom is not None:
            LOGGER.debug("Element self reports room with room property")
            return [elementRoom]

    LOGGER.debug("Element does not self report room, searching for intersecting rooms")

    # Get the elements bounding box
    elementBoundingBox = element.get_BoundingBox(None)
    if not elementBoundingBox.Enabled:
        LOGGER.debug("Bounding Box Not Enabled")
        return

    elementOutline = DB.Outline(
        elementBoundingBox.Min.Add(DB.XYZ(-1.0 * offset, -1.0 * offset, -1.0 * offset)),
        elementBoundingBox.Max.Add(DB.XYZ(offset, offset, offset)),
    )

    if projectToLevel and element.Location is not None:
        LOGGER.debug("Projecting to level")
        elementOutline = _ProjectToLevel(elementOutline, element)

    boundingBoxIntersectsFilter = DB.BoundingBoxIntersectsFilter(elementOutline)

    LOGGER.debug("Searching for intersecting rooms")
    if rooms:
        # msg = ["roomId={}/n".format(room.Id) for room in rooms]
        msg = ["room type={}/n".format(type(room)) for room in rooms]
        print(msg)
        rooms = (
            DB.FilteredElementCollector(doc)
            .WherePasses(
                DB.ElementIdSetFilter(
                    List[DB.ElementId](room.Id for room in rooms if hasattr(room, "Id"))
                )
            )
            .WherePasses(boundingBoxIntersectsFilter)
            .ToElements()
        )
    else:
        rooms = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_Rooms)
            .WherePasses(boundingBoxIntersectsFilter)
            .ToElements()
        )
    LOGGER.debug("Number of intersecting rooms: {}".format(len(rooms)))
    elementSolids = GetSolids(element)
    matchedRooms = _GetMatchingRooms(element, elementSolids, elementOutline, rooms)
    return matchedRooms


def _GetMatchingRooms(element, elementSolids, elementOutline, rooms):
    LOGGER.debug(
        "_GetMatchingRooms(element={}, elementSolids={}, elementOutline={}, "
        "len(rooms)={})".format(element.Id, elementSolids, elementOutline, len(rooms))
    )
    LOGGER.info("Matching room with element solid method")
    matchedRooms = []
    for room in rooms:
        roomSolid = GetSolids(room)[0]
        for elementSolid in elementSolids:
            matchedRoom, volume = _TestRoomIntersect(room, elementSolid)
            if matchedRoom:
                matchedRooms.append(matchedRoom)

    if not matchedRooms:
        LOGGER.info("No Matches: Getting dependent elements and trying again")
        for room in rooms:
            roomSolid = GetSolids(room)[0]
            dependentElements = element.GetDependentElements(
                DB.ElementIntersectsSolidFilter(roomSolid)
            )
            LOGGER.debug("len(dependentElements) = {}".format(len(dependentElements)))
            if dependentElements:
                matchedRooms.append(room)

    if not matchedRooms:
        LOGGER.info("No Matches: Resorting to bounding box method")
        elementSolid = MakeSolid(
            elementOutline.MinimumPoint, elementOutline.MaximumPoint
        )
        roomMatchVolumes = {}
        for room in rooms:
            matchedRoom, volume = _TestRoomIntersect(room, elementSolid)
            if matchedRoom:
                roomMatchVolumes[volume] = room
        # get the room with the largest volume
        if roomMatchVolumes:
            matchedRooms = [roomMatchVolumes[max(roomMatchVolumes.keys())]]

    return matchedRooms


def GetScheduledParameterIds(scheduleView):
    """Get the ids of the parameters that have been added to the provided
    schedule

    Args:
        scheduleView (DB.ViewSchedule): Schedule who's parameter ids are to be
            collected

    Returns:
        list[DB.ElementId] or None: List of parameter ids found in the schedule
    """
    fieldOrder = scheduleView.Definition.GetFieldOrder()
    fields = [
        scheduleView.Definition.GetField(fieldId)
        for fieldId in fieldOrder
        if fieldId is not None
    ]
    return [field.ParameterId for field in fields]


def GetScheduledParameterByName(scheduleView, parameterName, doc=None):
    """Gets a parameter included in the provided schedule by name. This just
    grabs the first parameter with the provided name. It may cause issues if
    multiple parameters with the same name are scheduled.

    Args:
        scheduleView (DB.ViewSchedule): Schedule view to search for parameter
        parameterName (str): Name of the parameter
        doc (DB.Document, optional): Revit document that hosts the parameter.
            Defaults to None.

    Returns:
        DB.Parameter: Revit parameter with the provided name
    """
    if doc is None:
        doc = HOST_APP.doc
    parameterIds = GetScheduledParameterIds(scheduleView=scheduleView)
    parameters = [doc.GetElement(parameterId) for parameterId in parameterIds]
    parameterMap = {
        parameter.Name: parameter for parameter in parameters if parameter is not None
    }
    if parameterName in parameterMap:
        return parameterMap[parameterName]
    else:
        return None


def GetAllGroups(includeModelGroups=True, includeDetailGroups=True, doc=None):
    doc = doc or HOST_APP.doc
    builtInCategories = List[DB.BuiltInCategory]()
    if includeModelGroups:
        builtInCategories.Add(DB.BuiltInCategory.OST_IOSModelGroups)
    if includeDetailGroups:
        builtInCategories.Add(DB.BuiltInCategory.OST_IOSDetailGroups)
    if not builtInCategories:
        return
    multiCategoryFilter = DB.ElementMulticategoryFilter(builtInCategories)
    groups = (
        DB.FilteredElementCollector(doc)
        .WherePasses(multiCategoryFilter)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return groups


def UngroupAllGroups(includeDetailGroups=True, includeModelGroups=True, doc=None):
    doc = doc or HOST_APP.doc
    groups = GetAllGroups(
        includeDetailGroups=includeDetailGroups,
        includeModelGroups=includeModelGroups,
        doc=doc,
    )
    groupTypes = set(group.GroupType for group in groups)
    [group.UngroupMembers() for group in groups]
    return groupTypes


def GetAllElementsInModelGroups(doc=None):
    doc = doc or HOST_APP.doc
    modelGroups = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_IOSModelGroups)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return set(
        doc.GetElement(memberId)
        for modelGroup in modelGroups
        for memberId in modelGroup.GetMemberIds()
    )


def GetAllElementIdsInModelGroups(doc=None):
    doc = doc or HOST_APP.doc
    modelGroups = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_IOSModelGroups)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return set(
        memberId for modelGroup in modelGroups for memberId in modelGroup.GetMemberIds()
    )


def SetParameter(element, parameterName, value):
    LOGGER.debug("SetParameter(element={},{},{})".format(element, parameterName, value))
    if type(parameterName) == Guid:
        LOGGER.debug("ParameterName is a Guid")
        parameter = element.get_Parameter(parameterName)
    elif type(parameterName) == DB.BuiltInParameter:
        LOGGER.debug("ParameterName is a BuiltInParameter")
        parameter = element.get_Parameter(parameterName)
    else:
        parameter = element.LookupParameter(parameterName)
    assert parameter
    parameter.Set(value)
    return parameter


def GetParameterValue(parameter, asValueString=False):
    LOGGER.debug("GetParameterValue")
    if asValueString == True:
        LOGGER.debug("AsValueString")
        return parameter.AsValueString()
    storageType = parameter.StorageType
    if storageType == DB.StorageType.Integer:
        LOGGER.debug("AsInteger")
        return parameter.AsInteger()
    elif storageType == DB.StorageType.Double:
        LOGGER.debug("AsDouble")
        return parameter.AsDouble()
    elif storageType == DB.StorageType.String:
        LOGGER.debug("AsString")
        return parameter.AsString()
    elif storageType == DB.StorageType.ElementId:
        LOGGER.debug("AsElementId")
        return parameter.AsElementId()
    LOGGER.debug("No matching storage type")
    return


def GetParameterValueByName(element, parameterName, asValueString=False):
    LOGGER.debug("GetParameterValueByName")
    if type(parameterName) == Guid or type(parameterName) == DB.BuiltInParameter:
        parameter = element.get_Parameter(parameterName)
    else:
        parameter = element.LookupParameter(parameterName)
    if parameter:
        return GetParameterValue(parameter, asValueString=asValueString)
    return


def SetFamilyParameterFormula(parameterName, formula, familyDoc=None):
    familyDoc = familyDoc or HOST_APP.doc
    familyManager = familyDoc.FamilyManager
    if not familyDoc.IsFamilyDocument:
        LOGGER.warn("SetParameterFormula may only be used on a family document")
        return
    if type(parameterName) == Guid or type(parameterName) == DB.BuiltInParameter:
        parameter = familyManager.get_Parameter(parameterName)
    else:
        parameter = familyManager.LookupParameter(parameterName)
    if not parameter:
        LOGGER.debug("No parameter found: {}".format(familyDoc.Title))
        return
    familyManager.SetFormula(parameter, formula)
    return parameter


def CopyParameterValue(element, source, destination):
    """Copy the value from the `source` parameter to the `destination`. Will only be
    successful on parameters that have the same storage type.

    Args:
        element (Autodesk.Revit.DB.Element): Revit element that hosts the parameter
            value to be copied
        source (Autodesk.Revit.DB.Parameter): Parameter that will provide the value to
            be copied
        destination (Autodesk.Revit.DB.Parameter): Parameter to which the value is
            copied

    Returns:
        Autodesk.Revit.DB.Parameter: Updated destination parameter
    """
    LOGGER.debug("CopyParameterValue")
    if type(destination) is not Guid:
        destination = Guid(destination)
    if type(source) is str:
        sourceParameter = element.LookupParameter(source)
    else:
        sourceParameter = element.get_Parameter(source)
    destinationParameter = element.get_Parameter(destination)
    if sourceParameter is None:
        return
    if destinationParameter is None:
        return
    try:
        if sourceParameter.StorageType == DB.StorageType.Double:
            value = sourceParameter.AsDouble()
        elif sourceParameter.StorageType == DB.StorageType.Integer:
            value = sourceParameter.AsInteger()
        elif sourceParameter.StorageType == DB.StorageType.ElementId:
            value = sourceParameter.AsElementId()
        else:
            value = sourceParameter.AsString()
        LOGGER.debug("value = {}".format(value))
        destinationParameter.Set(value or "")
    except (AttributeError, TypeError) as e:
        LOGGER.warn("\tError:{}".format(e))
    return destinationParameter


def GetElementsVisibleInView(
    view=None, IncludeLinkModelElements=True, displayGeometry=False
):
    if view is None:
        doc = HOST_APP.doc
        view = doc.ActiveView
    else:
        doc = view.Document

    elements = (
        DB.FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    print("View Element Count: {}".format(len(elements)))
    if IncludeLinkModelElements:
        categories = doc.Settings.Categories
        viewPhase = GetViewPhase(view)
        print("len(elements) = {}".format(len(elements)))
        visibleModelCategories = []
        for category in categories:
            try:
                if category.IsVisibleInUI:
                    if category.CategoryType == DB.CategoryType.Model:
                        visibleModelCategories.append(category)
            except Exception as e:
                print(e)
        viewModelCategories = List[DB.BuiltInCategory](
            [
                category.Id.IntegerValue
                for category in visibleModelCategories
                if not view.GetCategoryHidden(category.Id)
            ]
        )
        print("len(viewModelCategories) = {}".format(len(viewModelCategories)))
        viewBoundingBox = view.GetSectionBox()
        rvtLinks = (
            DB.FilteredElementCollector(doc, view.Id)
            .OfCategory(DB.BuiltInCategory.OST_RvtLinks)
            .ToElements()
        )
        for rvtLink in rvtLinks:
            linkDoc = rvtLink.GetLinkDocument()
            linkOffset = rvtLink.GetTotalTransform().Origin
            print("linkOffset = {}".format(linkOffset))
            rvtLinkType = doc.GetElement(rvtLink.GetTypeId())
            phaseMap = rvtLinkType.GetPhaseMap()
            linkPhaseId = phaseMap.TryGetValue(viewPhase.Id)[1]
            print(type(linkPhaseId))
            print(linkPhaseId)
            elementOnPhaseStatusFilter = DB.ElementPhaseStatusFilter(
                linkPhaseId,
                List[DB.ElementOnPhaseStatus](
                    [
                        DB.ElementOnPhaseStatus.New,
                        DB.ElementOnPhaseStatus.Existing,
                    ]
                ),
            )
            linkElements = (
                DB.FilteredElementCollector(linkDoc)
                .WhereElementIsNotElementType()
                .WherePasses(elementOnPhaseStatusFilter)
                .WherePasses(DB.ElementMulticategoryFilter(viewModelCategories))
                .ToElements()
            )
            print("len(linkElements) = {}".format(len(linkElements)))
            elements = list(elements) + list(linkElements)

            if displayGeometry:
                print("Displaying geometry")
                with revit.Transaction("Add direct shapes"):
                    for element in linkElements:
                        try:
                            solids = GetSolids(element)
                            translatedSolids = List[DB.GeometryObject]()
                            if linkOffset.IsAlmostEqualTo(DB.XYZ.Zero):
                                translatedSolids = solids
                            else:
                                for solid in solids:
                                    if not solid:
                                        continue
                                    translatedSolid = DB.SolidUtils.CreateTransformed(
                                        solid,
                                        DB.Transform.CreateTranslation(linkOffset),
                                    )
                                    translatedSolids.Add(translatedSolid)
                            if translatedSolids:
                                directShape = DB.DirectShape.CreateElement(
                                    doc,
                                    DB.ElementId(DB.BuiltInCategory.OST_GenericModel),
                                )
                                directShape.SetShape(translatedSolids)
                        except Exception as e:
                            print("solid exception: {}".format(e))

        print("View + Links Element Count: {}".format(len(elements)))

    return elements


def ExportScheduleAsDictionary(viewSchedule):
    LOGGER.set_debug_mode()
    LOGGER.info("flamingo.revit.ExportScheduleAsDictionary")
    LOGGER.debug(
        "viewSchedule.Id.IntegerValue = {}".format(viewSchedule.Id.IntegerValue)
    )
    #  DB.SectionType.None
    #  DB.SectionType.Header
    #  DB.SectionType.Body
    #  DB.SectionType.Summary
    #  DB.SectionType.Footer
    dictionary = []
    if type(viewSchedule) == DB.ViewSchedule:
        tableData = viewSchedule.GetTableData()
        LOGGER.info("NumberOfSections = {}".format(tableData.NumberOfSections))
        for i in range(tableData.NumberOfSections):
            sectionData = tableData.GetSectionData(i)
            LOGGER.debug(
                "tableData.NumberOfColumns = {}".format(sectionData.NumberOfColumns)
            )
            LOGGER.debug(
                "sectionData.NumberOfRows = {}".format(sectionData.NumberOfRows)
            )
            for x in range(sectionData.NumberOfRows):
                rowData = []
                for y in range(sectionData.NumberOfColumns):
                    cellText = GetCellValueAsText(sectionData, x, y)
                    calculatedValue = sectionData.GetCellCalculatedValue(x, y)
                    cellType = sectionData.GetCellType(x, y)
                    LOGGER.debug(
                        "cell {},{} = {} ({})".format(x, y, cellText, cellType)
                    )
                    rowData.append(cellText)
                dictionary.append(rowData)
    return dictionary


def GetCellValueAsText(sectionData, x, y, doc=None):
    doc = doc or HOST_APP.doc
    LOGGER.info("flamingo.revit.GetCellValueAsText")
    cellType = sectionData.GetCellType(x, y)
    if cellType in (DB.CellType.Text, DB.CellType.ParameterText):
        LOGGER.debug("Getting text")
        return sectionData.GetCellText(x, y)
    elif cellType == DB.CellType.Parameter:
        LOGGER.debug("Getting parameter")
        parameterId = sectionData.GetCellParamId(x, y).IntegerValue
        LOGGER.debug("type(parameterId) = {}".format(type(parameterId)))
        LOGGER.debug(
            "doc.GetElement(parameterId) = {}".format(doc.GetElement(parameterId))
        )
        return GetParameterValue(doc.GetElement(parameterId))
    elif cellType == DB.CellType.CalculatedValue:
        LOGGER.debug("Getting calculated value")
        return sectionData.GetCellCalculatedValue(x, y)
    elif cellType == DB.CellType.Inherited:
        LOGGER.debug("Getting inherited")
        return "INHERITED"
    elif cellType == DB.CellType.CombinedParameter(x, y):
        LOGGER.debug("Getting Combined Parameters")
        combinedParameters = sectionData.GetCellCombinedParameters(x, y)
        cellValue = ""
        LOGGER.debug("combinedParameters = {}".format(combinedParameters))
        return "COMBINED PARAMETERS"
    return


def GetLinkLoadTimes(logFilePath=None, doc=None):
    logFilePath = logFilePath or HOST_APP.app.RecordingJournalFilename
    doc = doc or HOST_APP.doc
    LOGGER.debug("logFilePath = {}".format(logFilePath))
    LOGGER.debug("doc.Title = {}".format(doc.Title))
    out = {}
    with open(logFilePath, "r") as f:
        currentTimestamp = None
        currentJournalC = None
        activeDocumentPath = None
        detach = False
        processNextLine = False
        for lineNumber, line in enumerate(f):
            if processNextLine:
                processNextLine = False
                m = re.search(r"([^\\]+?\.rvt)", line)
                if detach:
                    detach = False
                    filePath = "{}_detached.rvt".format(m.group(1)[0:-4])
                else:
                    filePath = m.group(1)
                activeDocumentPath = filePath
                continue
            m = re.search(r"'C (.*);\s+(.*)", line.strip())
            if m:
                currentTimestamp = m.group(1)
                currentJournalC = m.group(2)
                continue
            if '"DetachCheckBox", "True"' in line:
                detach = True
                LOGGER.debug("detach = {}".format(detach))
                continue
            if 'Jrn.Data "File Name"' in line:
                processNextLine = True
                continue
            m = re.search(r'>Open:Local.*".+\\(.+)"', line)
            if m:
                activeDocumentPath = m.group(1)
                continue
            m = re.search(r'"\[(.*)\]"', line)
            if m:
                activeDocumentPath = m.group(1)
                continue
            m = re.search(r"openFromModelPath.+\[(.*)\]", line)
            if m:
                # LOGGER.debug(
                #     "{}: {} {}|{}".format(
                #         lineNumber, currentTimestamp, activeDocumentPath, m.group(1)
                #     )
                # )
                # documentLoad = (currentTimestamp, m.group(1))
                # 'C 24-Oct-2022 13:39:40.220;  ->desktop InitApplication
                timestamp = datetime.strptime(currentTimestamp, "%d-%b-%Y %H:%M:%S.%f")
                if activeDocumentPath in out:
                    out[activeDocumentPath][m.group(1)] = timestamp
                else:
                    out[activeDocumentPath] = {m.group(1): timestamp}
    return out


def OperateOnNestedFamilies(
    families,
    function,
    familyRenameFunction=None,
    includeShared=True,
    doc=None,
    nestLevel=None,
    maxNest=32,
    closeDocs=True,
    **kwargs
):
    """Takes a function and applies it to the provided families and all families
    nested within. Shared families will only be processed once and loaded into the
    the primary upstream host family. The function that is passed to the family must
    include a 'doc' parameter to pass the opened family document (Autodesk.DB.Document).
    If an transaction is required, it should be included in the passed function. Do not
    nest this function in another transaction.

    Args:
        families (list(DB.Family)): List of families to operate on
        function (func): Function applied to each family and nest
        includeShared (bool, optional): Process shared families. Defaults to True.
        doc (Autodesk.Revit.DB.Document, optional): Document of host. Can be a family
            or project document. Defaults to active Revit document.
        nestLevel (int, optional): Current level of family nesting. Defaults to 0.
        maxNest (int, optional): Maximum level of nested families to recurse. Defaults
            to 32.
        closeDocs (bool, optional): Close the documents after processing is complete.
            Defaults to True.

    Returns:
        List of family names that where processed (list[str])
    """

    processedFamilyNames = []
    swallowedErrors = []
    doc = doc or HOST_APP.doc
    nestLevel = nestLevel or 0
    prefix = "\t" * nestLevel
    LOGGER.debug("{}OperateOnNestedFamilies: Nest level {}".format(prefix, nestLevel))
    LOGGER.debug("{}includeShared = {}".format(prefix, includeShared))
    for family in families:
        if nestLevel > maxNest:
            raise PyRevitException(
                "Reached max nest level of {}. If more nesting is required"
                ' set the "maxNest" parameter'.format(maxNest)
            )
        LOGGER.debug("{}family.Name = {}".format(prefix, family.Name))
        if "{}.rfa".format(family.Name) == doc.Title:
            LOGGER.warn(
                'Nested family "{}" has the same name as the host causing a '
                "circular reference".format(family.Name)
            )
            continue
        if not family.IsEditable:
            LOGGER.debug("{}Family is not editable".format(prefix))
            continue
        isShared = family.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED).AsInteger()
        LOGGER.debug("{}isShared = {}".format(prefix, isShared))
        if not includeShared and isShared:
            LOGGER.debug("{}Skipping shared family".format(prefix))
            continue
        with revit.ErrorSwallower() as swallower:
            familyDoc = family.Document.EditFamily(family)
            swallowedErrors = swallower.get_swallowed_errors()
            if swallowedErrors:
                msg = "\n".join(
                    [error.GetDescriptionText() for error in swallowedErrors]
                )
                msg = "{}\n{}".format(OUTPUT.linkify(family.Id), msg)
                LOGGER.warn(msg)
                swallowedErrors.append(msg)
        LOGGER.debug("{}familyDoc.Title = {}".format(prefix, familyDoc.Title))
        function(doc=familyDoc, nestLevel=nestLevel, **kwargs)
        processedFamilyNames.append(family.Name)
        nestedFamilies = (
            DB.FilteredElementCollector(familyDoc).OfClass(DB.Family).ToElements()
        )
        if nestedFamilies:
            if familyRenameFunction:
                for nestedFamily in nestedFamilies:
                    familyRenameFunction(nestedFamily)
            processedFamilyNamesFromNests = OperateOnNestedFamilies(
                nestedFamilies,
                function,
                familyRenameFunction=familyRenameFunction,
                doc=familyDoc,
                includeShared=False,
                nestLevel=nestLevel + 1,
                closeDocs=True,
                **kwargs,
            )
            if processedFamilyNamesFromNests:
                processedFamilyNames = (
                    processedFamilyNames + processedFamilyNamesFromNests
                )

        LOGGER.debug("{}Loading family back in: {}".format(prefix, familyDoc.Title))
        familyDoc.LoadFamily(doc, iFamilyLoadOptions())
        if closeDocs:
            try:
                familyDoc.Close(False)
                LOGGER.debug("Document Closed")
            except Exception as e:
                LOGGER.warn("{}Unable to close family: {}".format(prefix, e))
    for error in swallowedErrors:
        print(error)
    return processedFamilyNames


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    """Handle copy and paste errors."""

    def OnDuplicateTypeNamesFound(self, args):
        """Use destination model types if duplicate."""
        return DB.DuplicateTypeAction.UseDestinationTypes


def CopyElementsFromOtherDocument(
    sourceDoc, targetDoc, elementIds, closeSourceDoc=False
):
    if type(sourceDoc) == str:
        from flamingo.ensure import EnsureLibraryDoc

        sourceDoc = EnsureLibraryDoc(sourceDoc)

    copyPasteOptions = DB.CopyPasteOptions()
    copyPasteOptions.SetDuplicateTypeNamesHandler(CopyUseDestination())
    LOGGER.debug("elementIds = {}".format(elementIds))
    copiedElementIds = DB.ElementTransformUtils.CopyElements(
        sourceDoc,
        List[DB.ElementId](elementIds),
        targetDoc,
        None,
        copyPasteOptions,
    )
    if closeSourceDoc:
        sourceDoc.Close(False)
    return copiedElementIds


def GetElementSymbolId(element):
    LOGGER.debug("GetElementSymbol")
    if hasattr(element, "WallType"):
        return element.WallType.Id
    if hasattr(element, "CurtainSystemType"):
        return element.CurtainSystemType.Id
    elif hasattr(element, "Symbol"):
        return element.Symbol.Id
    elif hasattr(element, "TypeId"):
        return element.TypeId
    LOGGER.warn("{} Unable to get symbol".format(OUTPUT.linkify(element.Id)))
    return


def GetSolidFillId(doc):
    LOGGER.debug("GetSolidFillId({})".format(doc))
    fillPatternElements = (
        DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement).ToElements()
    )
    for fillPatternElement in fillPatternElements:
        if fillPatternElement.GetFillPattern().IsSolidFill:
            return fillPatternElement.Id
    return
