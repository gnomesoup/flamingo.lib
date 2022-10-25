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
        PyRevitException("Could not read from the shared parameters file")
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
        PyRevitException(
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

    # Create the parameter
    if isTypeParameter:
        newBinding = app.Create.NewTypeBinding(categorySet)
    else:
        newBinding = app.Create.NewInstanceBinding(categorySet)
    return doc.ParameterBindings.Insert(externalDefinition, newBinding, parameterGroup)


def GetParameterFromProjectInfo(doc, parameterName):
    """
    Returns a parameter value from the Project Information category by name.
    """
    try:
        projectInformation = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_ProjectInformation)
            .ToElements()
        )
        parameterValue = projectInformation[0].LookupParameter(parameterName).AsString()
    except:
        return None
    return parameterValue


def SetParameterFromProjectInfo(doc, parameterName, parameterValue):
    """
    Set a parameter value from the Project Information category by name.
    """
    if True:
        # try:
        projectInformation = (
            DB.FilteredElementCollector(doc)
            .OfCategory(DB.BuiltInCategory.OST_ProjectInformation)
            .ToElements()
        )
        parameter = projectInformation[0].LookupParameter(parameterName)
        parameter.Set(parameterValue)
    # except:
    #     return None
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
    try:
        elementMaterials = element.GetMaterialIds(False)
        paintMaterials = element.GetMaterialIds(True)
        for materialId in paintMaterials:
            elementMaterials.append(materialId)
        return elementMaterials
    except Exception as e:
        print(e)
        return []


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


def OpenLocal(filePath, localDir, extension=""):
    LOGGER.debug("Flamingo \"OpenLocal\" called")
    localDir = path.expandvars(localDir)
    centralFileName = path.basename(filePath)
    centralName = re.sub(r"\.rvt$", "", centralFileName)
    localFileName = "{}_{}{}.rvt".format(centralName, HOST_APP.username, extension)
    localPath = path.join(localDir, localFileName)
    LOGGER.info("filePath={}".format(filePath))
    LOGGER.info("localPath={}".format(localPath))
    shutil.copyfile(filePath, localPath)
    return HOST_APP.uiapp.OpenAndActivateDocument(localPath)


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

    allDoors = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Doors)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    # Make a dictionary of rooms with door properties
    roomDoors = {}
    for door in allDoors:
        if not door:
            continue
        roomsToAdd = []
        toRoom = (door.ToRoom)[phase]
        if toRoom:
            if toRoom.Id in roomDoors:
                roomDoors[toRoom.Id]["doors"] = roomDoors[toRoom.Id]["doors"] + [
                    door.Id
                ]
            else:
                roomsToAdd.append(toRoom)
        fromRoom = (door.FromRoom)[phase]
        if fromRoom:
            if fromRoom.Id in roomDoors:
                roomDoors[fromRoom.Id]["doors"] = roomDoors[fromRoom.Id]["doors"] + [
                    door.Id
                ]
            else:
                roomsToAdd.append(fromRoom)
        for roomToAdd in roomsToAdd:
            roomDoors[roomToAdd.Id] = {"doors": [door.Id]}
            roomDoors[roomToAdd.Id]["roomNumber"] = roomToAdd.Number
            roomDoors[roomToAdd.Id]["roomArea"] = roomToAdd.Area
            roomCenter = roomToAdd.Location.Point
            roomDoors[roomToAdd.Id]["roomLocation"] = roomCenter

    # Make a dictionary of door connection counts. This will be used to
    # prioritize doors with more connections when assigning a letter value
    roomDoorsDoors = [value["doors"] for value in roomDoors.values()]
    doorConnectors = {}
    for door in allDoors:
        for i in roomDoorsDoors:
            if door.Id in i:
                if door.Id in doorConnectors:
                    doorConnectors[door.Id] += len(i)
                else:
                    doorConnectors[door.Id] = len(i)

    maxLength = max([len(values["doors"]) for values in roomDoors.values()])
    for key, values in roomDoors.items():
        roomDoors[key]["doorCount"] = len(values["doors"])

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
                roomId for roomId, value in roomDoors.items() if value["doorCount"] == i
            ]
            roomsThisRound = sorted(
                roomsThisRound, key=lambda x: roomDoors[x]["roomArea"]
            )
            # Go through each room and look for door counts that match
            # current level
            # for key, values in roomDoors.items():
            for roomId in roomsThisRound:
                values = roomDoors[roomId]
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
                        markParameter = door.get_Parameter(
                            DB.BuiltInParameter.ALL_MODEL_MARK
                        )
                        markParameter.Set(mark)
            for roomId, values in roomDoors.items():
                unNumberedDoors = filter(
                    None,
                    [
                        doorId
                        for doorId in values["doors"]
                        if doorId not in numberedDoors
                    ],
                )
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


def GetElementRoom(element, phase, offset=1.0, projectToLevel=True, doc=None):
    """Find the room in a Revit model that a element is placed in or near.
    The function should find the room if it is located above or within a
    specified offset

    Args:
        element (DB.FamilyInstance): Element who's room is to be located
        phase (DB.Phase): The phase of the room that is to be
            matched to the element.
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
    if doc is None:
        doc = HOST_APP.doc

    room = (element.Room)[phase]
    if room is not None:
        return room

    elementBoundingBox = element.get_BoundingBox(None)
    if not elementBoundingBox.Enabled:
        print("Bounding Box Not Enabled")
        return

    elementOutline = DB.Outline(
        elementBoundingBox.Min.Add(DB.XYZ(-1.0 * offset, -1.0 * offset, -1.0 * offset)),
        elementBoundingBox.Max.Add(DB.XYZ(offset, offset, offset)),
    )

    if projectToLevel and element.Location is not None:
        elementLevel = doc.GetElement(element.LevelId)
        levelElevation = elementLevel.Elevation
        elementPoint = element.Location.Point
        elementOutline.AddPoint(DB.XYZ(elementPoint.X, elementPoint.Y, levelElevation))

    boundingBoxIntersectsFilter = DB.BoundingBoxIntersectsFilter(elementOutline)

    rooms = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Rooms)
        .WherePasses(boundingBoxIntersectsFilter)
        .ToElements()
    )
    elementSolid = MakeSolid(elementOutline.MinimumPoint, elementOutline.MaximumPoint)
    matchedRoom = None
    maxVolume = 0
    for room in rooms:
        solids = [
            geometryElement
            for geometryElement in room.get_Geometry(DB.Options())
            if type(geometryElement) == DB.Solid
        ]
        interSolid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
            solids[0], elementSolid, DB.BooleanOperationsType.Intersect
        )
        if abs(interSolid.Volume > 0.000001):
            if interSolid.Volume > maxVolume:
                maxVolume = interSolid.Volume
                matchedRoom = room
    return matchedRoom


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


def UngroupAllGroups(includeDetailGroups=True, includeModelGroups=True, doc=None):
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
    groupTypes = set(group.GroupType for group in groups)
    [group.UngroupMembers() for group in groups]
    return groupTypes


def GetAllElementsInModelGroups(doc=None):
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


def SetParameter(element, parameterName, value):
    try:
        if type(parameterName) == Guid:
            parameter = element.get_Parameter(parameterName)
        elif type(parameterName) == DB.BuiltInParameter:
            parameter = element.get_Parameter(parameterName)
        else:
            parameter = element.LookupParameter(parameterName)
        if parameter:
            parameter.Set(value)
        else:
            LOGGER.info("Parameter {} missing project.".format(parameterName))
            parameter.Set(value)
    except AttributeError as e:
        LOGGER.debug("SetParameter Error: {}".format(e))
    return parameter


def GetParameterValue(parameter):
    unitType = parameter.StorageType
    if unitType == DB.StorageType.Integer:
        return parameter.AsInteger()
    elif unitType == DB.StorageType.Double:
        return parameter.AsDouble()
    elif unitType == DB.StorageType.String:
        return parameter.AsString()
    elif unitType == DB.StorageType.ElementId:
        return parameter.AsElementId()
    return


def GetParameterValueByName(element, parameterName):
    if type(parameterName) == Guid or type(parameterName) == DB.BuiltInParameter:
        parameter = element.get_Parameter(parameterName)
    else:
        parameter = element.LookupParameter(parameterName)
    if parameter:
        return GetParameterValue(parameter)
    return


def SetFamilyParameterFormula(parameterName, formula, familyDoc=None):
    familyDoc = familyDoc or HOST_APP.doc
    familyManager = familyDoc.FamilyManager
    if not familyDoc.IsFamilyDocument:
        LOGGER.debug("SetParameterFormula may only be used on a family document")
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
                                        DB.Transform.CreateTranslation(linkOffset)
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
    LOGGER.debug("viewSchedule.Id.IntegerValue = {}".format(viewSchedule.Id.IntegerValue))
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
            LOGGER.debug("tableData.NumberOfColumns = {}".format(sectionData.NumberOfColumns))
            LOGGER.debug("sectionData.NumberOfRows = {}".format(sectionData.NumberOfRows))
            for x in range(sectionData.NumberOfRows):
                rowData = []
                for y in range(sectionData.NumberOfColumns):
                    cellText = GetCellValueAsText(sectionData, x, y)
                    calculatedValue = sectionData.GetCellCalculatedValue(x,y)
                    cellType = sectionData.GetCellType(x,y)
                    LOGGER.debug("cell {},{} = {} ({})".format(x,y, cellText, cellType))
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
        LOGGER.debug("doc.GetElement(parameterId) = {}".format(doc.GetElement(parameterId)))
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
                    detach=False
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
                detach=True
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
