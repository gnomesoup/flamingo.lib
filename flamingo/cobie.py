from Autodesk.Revit import DB
from flamingo.revit import GetViewPhase, GetElementRoom
from flamingo.revit import GetScheduledParameterIds, GetScheduledParameterByName
from flamingo.revit import GetParameterFromProjectInfo
from pyrevit import forms, HOST_APP, revit, script

def SetCOBieParameter(element, parameterName, value, blankOnly=False):
    """Set a value for a COBie parameter. The parameter must be a string
    parameter.

    Args:
        element (DB.Element): Element who's value is to be set
        parameterName (str): Name of parameter to set
        value (str): Value to set to parameter
        blankOnly (bool, optional): If set to true, the value will not be set
            if the parameter already has a non-blank value. Defaults to False.

    Returns:
        DB.Element: Returns the element regardless of whether the value was
            updated.
    """
    cobieParameter = element.LookupParameter(parameterName)
    # currentvalue = cobieParameter.AsString()
    # if currentvalue != value:
    #     output = script.get_output()
    #     print(
    #         "{} {} != {}".format(
    #             output.linkify(element.Id), currentvalue, value
    #         )
    #     )
    if blankOnly:
        currentValue = cobieParameter.AsString()
        if currentValue is not None and currentValue != "":
            return element
    cobieParameter.Set(value)
    return element

def COBieParameterIsBlank(element, parameterName):
    """Check if a COBie parameter for an element is blank. Only checks string
        parameter types.

    Args:
        element (DB.Element): Revit element who's parameter is to be checked
        parameterName (str): Name of the parmater who's value is to be checked

    Returns:
        bool: True if the parameter value is blank. False otherwise.
    """
    parameter = element.LookupParameter(parameterName)
    parameterValue = parameter.AsString()
    if parameterValue is None or parameterValue == "":
        return True
    else:
        return False

def SetCOBieComponentSpace(view, blankOnly, doc=None):
    """Finds matches a room to every element listed in the COBie.Components
    schedule that has the "COBie" parameter checked. It then assigns the value
    of matched room number to the COBie.Component.Space parameter.

    Args:
        view (DB.ScheduleView): COBie Component Schedule
        blankOnly (bool): If true, only update values that are currently blank
        doc (DB.Document, optional): [description]. Defaults to None.
    """
    if doc is None:
        doc = HOST_APP.doc
    phase = GetViewPhase(view, doc)
    elements = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType()\
        .ToElements()
    if blankOnly:
        elements = [
            element for element in elements
            if COBieParameterIsBlank(
                element,
                "COBie.Component.Space",
            )
        ]

    with revit.Transaction("Set COBie.Component.Space"):
        for element in elements:
            parameter = element.LookupParameter("COBie")
            try:
                if parameter.AsInteger() < 1:
                    continue
                room = GetElementRoom(
                    element=element,
                    phase=phase,
                    doc=doc
                )
                if room is None:
                    continue
                SetCOBieParameter(
                    element,
                    "COBie.Component.Space",
                    room.Number
                )
            except Exception as e:
                print("Error: {}".format(e))
    return

def COBieComponentSetDescription(view, blankOnly=True, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    elements = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType().ToElements()
    
    outElements = []
    with revit.Transaction("Set COBie.Component.Description"):
        for element in elements:
            isCobie = element.LookupParameter("COBie").AsInteger()
            if isCobie == 1:
                description = element.get_Parameter(
                    DB.BuiltInParameter.ELEM_TYPE_PARAM
                ).AsValueString()
                elementDescription = element.LookupParameter(
                    "COBie.Component.Description"
                    )
                if (
                    elementDescription.AsString() == "" or
                    elementDescription.AsString() == None or
                    not blankOnly
                ):
                    elementDescription.Set(description.replace("_", " "))
                    outElements.append(element)
    return outElements

def COBieComponentAssignMarks(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    elements = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    with revit.Transaction("Set unassigned instance marks"):
        marks = {}

        for element in elements:
            isCobie = element.LookupParameter("COBie").AsInteger()
            if isCobie != 1:
                continue
            symbolId = element.Symbol.Id.IntegerValue
            if symbolId not in marks:
                familyInstanceFilter = DB.FamilyInstanceFilter(
                    doc, element.Symbol.Id
                )
                instances = DB.FilteredElementCollector(doc, view.Id)\
                    .WherePasses(familyInstanceFilter) \
                    .ToElements()
                marks[symbolId] = {
                    "count": len(instances),
                    "marks": []
                }
            markParameter = element.get_Parameter(
                DB.BuiltInParameter.DOOR_NUMBER
            )
            mark = markParameter.AsString()
            if not mark or mark == "":
                for i in range(marks[symbolId]['count']):
                    mark = "{:03d}".format(i+1)
                    if mark not in marks[symbolId]['marks']:
                        marks[symbolId]['marks'].append(mark)
                        markParameter.Set(mark)
                        break
            
def COBieParameterBlankOut(scheduleList, doc=None):
    """Clears values for selected COBie paramaters. This is helpful for values
    that are not required at particular time points, like Manufacture at the CD
    project phase.

    Args:
        scheduleList (list[DB.ViewSchedule]): List of schedules to operate on
        doc (DB.Document, optional): Revit document that hosts the parameters to
            be blanked out. Defaults to None.
    """
    # TODO Only attempt to change a family symbol once
    # TODO Only operate on elements that have the "COBie" or "COBie.Type"
    #      parameters checked
    if doc is None:
        doc = HOST_APP.doc
    scheduledParameters = {}
    groupedParameterNames = {}
    for schedule in scheduleList:
        parameterIds = GetScheduledParameterIds(schedule)
        parameterNames = []
        for parameterId in parameterIds:
            if parameterId is not None and parameterId.IntegerValue > 0:
                parameter = doc.GetElement(parameterId)
                scheduledParameters[parameter.Name] = parameterId
                parameterNames.append(parameter.Name)
        groupedParameterNames[schedule.Name] = parameterNames
    parameterNamesOut = forms.SelectFromList.show(
        groupedParameterNames,
        multiselect=True,
        group_selector_title='COBie Schedule',
        button_name='Select Parameters'
    )

    if not parameterNamesOut:
        script.exit()

    bindings = doc.ParameterBindings
    with revit.Transaction("COBie Parameter Blank Out"):
        for parameterName in parameterNamesOut:
            parameterId = scheduledParameters[parameterName]
            parameter = doc.GetElement(parameterId)
            definition = parameter.GetDefinition()
            parameterBinding = bindings.get_Item(definition)
            parameterValueProvider = DB.ParameterValueProvider(
                parameterId
            )
            filterStringRule = DB.FilterStringRule(
                parameterValueProvider,
                DB.FilterStringGreater(),
                "",
                True
            )
            elementParameterFilter = DB.ElementParameterFilter(
                filterStringRule
            )
            if type(parameterBinding) is DB.TypeBinding:
                collector = DB.FilteredElementCollector(doc)\
                    .WhereElementIsElementType()
            else:
                collector = DB.FilteredElementCollector(doc)\
                    .WhereElementIsNotElementType()
            elements = collector.WherePasses(elementParameterFilter).ToElements()

            for element in elements:

                elementParameter = element.get_Parameter(parameter.GuidValue)
                elementParameter.Set("")

def COBieUncheckAll(typeViewSchedule, componentViewSchedule, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    print("Collecting COBie Type elements")
    cobieTypeParameter = GetScheduledParameterByName(
        typeViewSchedule,
        "COBie.Type",
        doc
    )
    cobieTypes = DB.FilteredElementCollector(doc)\
        .WhereElementIsElementType()\
        .WherePasses(
            DB.ElementParameterFilter(
                DB.HasNoValueFilterRule(cobieTypeParameter.Id)
            )
        ).ToElements()
    
    print("len(cobieTypes) = {}".format(len(cobieTypes)))
    print("")
    print("Collecting COBie Instance elements")
    cobieParameter = GetScheduledParameterByName(
        componentViewSchedule,
        "COBie",
        doc
    )
    hasNoValueFilterRule = DB.HasNoValueFilterRule(cobieParameter.Id)
    elementParameterFilter = DB.ElementParameterFilter(
        hasNoValueFilterRule
    )
    cobieInstances = DB.FilteredElementCollector(doc)\
        .WhereElementIsNotElementType()\
        .WherePasses(
            DB.ElementParameterFilter(
                DB.HasNoValueFilterRule(cobieParameter.Id)
            )
        ).ToElements()
    print("len(cobieInstances) = {}".format(len(cobieInstances)))
    print("")
    print("Setting COBie Type elements")

    progressMax = len(cobieTypes) + len(cobieInstances)
    progress = 0
    unsetElements = []
    with revit.TransactionGroup(
        "Uncheck COBie & COBie.Type Checkboxes",
    ):
        with forms.ProgressBar() as pb:
            with revit.Transaction(
                "Type Parmaeter Set",
                swallow_errors=True,
                show_error_dialog=False,
                # nested=True
            ) as t:
                for element in cobieTypes:
                    progress += 1
                    pb.update_progress(progress, progressMax)
                    try:
                        parameter = element.get_Parameter(cobieTypeParameter.GuidValue)
                        # parameter.HasValue = True
                        parameter.Set(0)
                    except Exception as e:
                        print("{}: {}".format(e, element.Id))
                        unsetElements.append(element)
            with revit.Transaction(
                "Type Parmaeter Set",
                swallow_errors=True,
                show_error_dialog=False,
                # nested=True
            ) as t:
                for element in cobieInstances:
                    progress += 1
                    pb.update_progress(progress, progressMax)
                    try:
                        parameter = element.get_Parameter(cobieParameter.GuidValue)
                        parameter.Set(0)
                    except Exception as e:
                        print("{}: {}".format(e, element.Id))
                        unsetElements.append(element)
    return unsetElements

def COBieIsEnabled(element):
    parameter = element.LookupParameter("COBie")
    try:
        assert parameter.AsInteger() > 0
        return True
    except (AssertionError, AttributeError):
        return False

def COBieTypeIsEnabled(element):
    parameter = element.LookupParameter("COBie.Type")
    try:
        assert parameter.AsInteger() > 0
        return True
    except (AssertionError, AttributeError):
        return False

def COBieEnable(element):
    parameter = element.LookupParameter("COBie")
    parameter.Set(True)

def GetElementSymbol(element):
    if type(element) == DB.Wall:
        return element.WallType
    else:
        return element.Symbol

def COBieComponentAutoSelect(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    elements = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType()\
        .ToElements()
    familySymbolIds = set()
    for element in elements:
        try:
            symbol = GetElementSymbol(element)
            if COBieTypeIsEnabled(symbol):
                familySymbolIds.add(symbol.Id)
        except Exception as e:
            print("{}: {}".format(element.Id.IntegerValue, e)) 
        
    componentView = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .WhereElementIsNotElementType()\
        .Where(lambda x: x.Name == "COBie.Component").First()
    componentElements = DB.FilteredElementCollector(doc, componentView.Id)\
        .ToElements()
    for element in componentElements:
        parameter = element.LookupParameter("COBie")
        try:
            if parameter.AsInteger() > 0:
                symbol = GetElementSymbol(element)
                if symbol.Id not in familySymbolIds:
                    parameter.Set(0)
        except AttributeError:
            continue

    with revit.Transaction("Enable COBie.Component Components"):
        
        for symbolId in familySymbolIds:
            symbolInstances = DB.FilteredElementCollector(doc)\
                .WhereElementIsNotElementType()\
                .WherePasses(
                    DB.FamilyInstanceFilter(doc, symbolId)
                )\
                .ToElements()
            for instance in symbolInstances:
                print("instance.Group = {}".format(instance.Group))
                parameter = instance.LookupParameter("COBie")
                parameter.Set(1)
    return

def COBieSpaceSetDescription(view, blankOnly=True, doc=None):
    doc = doc or HOST_APP.doc
    rooms = DB.FilteredElementCollector(doc, view.Id)\
        .WhereElementIsNotElementType().ToElements()

    roomsOut = []
    for room in rooms:
        if not COBieIsEnabled(room):
            continue
        roomName = room.get_Parameter(
                DB.BuiltInParameter.ROOM_NAME
            ).AsString()
        spaceDescription = room.LookupParameter("COBie.Space.Description")
        spaceDescriptionValue = spaceDescription.AsString()
        if (
            spaceDescriptionValue == "" or
            spaceDescriptionValue == None or
            not blankOnly
        ):
            spaceDescription.Set(roomName)
            roomsOut.append(room)
    return roomsOut

def COBieParametersToCDX(doc=None):
    doc = doc or HOST_APP.doc

    # Parameter Dictionary
    cdxParameterMap = [
        {
            "type": "Facility",
            "cobie": "COBie.Facility.Name",
            "cdx": "GSA.01.Facility.GSABuildingCode",
            "guid": "7704d27b-55e1-4e0d-a323-be403aee4cdd"
        },
        {
            "type": "Facility",
            "cobie": "COBie.Facility.Category",
            "cdx": "GSA.01.Facility.OmniClass.FacilityCategory",
            "guid": "a0a94036-dce0-45dc-b25b-b19632ceea23"
        },
        {
            "type": "Facility",
            "cobie": "COBie.Facility.ProjectName",
            "cdx": "GSA.01.Facility.ProjectName",
            "guid": "d18b5ea9-0c02-45a4-acd6-d8887a5bbad4"
        },
        {
            "type": "Facility",
            "cobie": "COBie.Facility.SiteName",
            "cdx": "GSA.01.Facility.SiteName",
            "guid": "52e4ef7a-08fb-44c5-8935-66ba1c2ee94b"
        },
        {
            "type": "Floor",
            "cobie": "COBie.Floor.Name",
            "cdx": "GSA.01.Floor.Name",
            "guid": "1bb202d3-d9e2-428a-8375-ee6d01a98ffa"
        },
        {
            "type": "Floor",
            "cobie": "COBie.Floor.Category",
            "cdx": "GSA.01.Floor.FloorCategory",
            "guid": "8836acff-ea52-41fe-bc49-8e81d1cd8605"
        },
        {
            "type": "Space",
            "cobie": "COBie.Space.Name",
            "cdx": "GSA.02.Space.RoomNumber",
            "guid": "57E67856-3996-4DC3-AD75-4C33A76704A0"
        },
        {
            "type": "Space",
            "cobie": "COBie.Space.Category",
            "cdx": "GSA.03.Space.OmniClass.SpaceCategory",
            "guid": "00DCCF44-54E8-4290-B386-2C6787C06376"
        },
        {
            "type": "Space",
            "cobie": "COBie.Space.FloorName",
            "cdx": "GSA.03.Space.FloorName",
            "guid": "FC13D9B2-31FE-4B58-A19D-F679C3F641CD"
        },
        {
            "type": "Space",
            "cobie": "COBie.Space.Description",
            "cdx": "GSA.03.Space.Description",
            "guid": "DF8F1C12-F287-4360-AFBC-CC92B0B56D2A"
        },
        {
            "type": "Space",
            "cobie": "COBie.Space.RoomTag",
            "cdx": "GSA.05.Space.Signage",
            "guid": "28ECE7CD-8818-436D-AF6C-733D3AE0E69B"
        },
        {
            "type": "Space",
            "cobie": "COBie.Zone.Name",
            "cdx": "GSA.03.Space.ZoneName",
            "guid": "B5DB0243-5AFE-4153-B037-775E21CC57F1"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.Name",
            "cdx": "GSA.05.Asset.TypeName",
            "guid": "0798BDCD-761D-4942-8FF2-0E7007F40E6A"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.Category",
            "cdx": "GSA.05.Asset.AssetType.Abbreviation",
            "guid": "37791872-234C-4766-9F31-7A3E72050582"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.Description",
            "cdx": "GSA.05.Asset.AssetType.Description",
            "guid": "ACD08F9A-6EAB-4981-8388-14A806DB7DFF"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.AssetType",
            "cdx": "GSA.05.Asset.Mobility",
            "guid": "F5A7AEB6-2216-45BA-B85F-C9E75201F9B2"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.Manufacturer",
            "cdx": "GSA.07.Asset.Manufacturer.Email",
            "guid": "639F8842-58BF-48C9-9B00-8C4D2790BA68"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.ModelNumber",
            "cdx": "GSA.07.Asset.Model.Number",
            "guid": "C20C3645-948F-4846-B80E-93049348FA0C"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.WarrantyDurationParts",
            "cdx": "GSA.07.Asset.WarrantyDuration.Parts",
            "guid": "4469578B-15B2-46DF-BA68-45BFA7D53AF7"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.WarrantyDurationLabor",
            "cdx": "GSA.07.Asset.WarrantyDuration.Labor",
            "guid": "C9021772-A5C2-471D-B116-3BD710D0FE31"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.WarrantyDurationUnit",
            "cdx": "GSA.07.Asset.WarrantyDuration.Unit",
            "guid": "4C4306D4-927C-4CD5-8501-6D58096AC4A3"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.ExpectedLife",
            "cdx": "GSA.07.Asset.ExpectedLife",
            "guid": "73F20796-C594-4929-A364-F34CBEB0CCD7"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Type.DurationUnit",
            "cdx": "GSA.07.Asset.ExpectedLife.Unit",
            "guid": "E70826DF-0EFA-4F2B-96FA-682DB5E6CDEA"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Component.Name",
            "cdx": "GSA.06.Asset.InstanceName",
            "guid": "E21B3065-6702-4973-9E1F-E3F21EF27581"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Component.Space",
            "cdx": "GSA.06.Asset.SpaceCode",
            "guid": "E2D4398B-F79D-486A-9B23-81614ED9AA84"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Component.Description",
            "cdx": "GSA.05.Asset.Description",
            "guid": "434BE294-6B1A-4B85-A30E-A2986B2AF718"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Component.SerialNumber",
            "cdx": "GSA.09.Asset.SerialNumber",
            "guid": "E700395C-92AA-4AE8-86A5-5F3EDC67C944"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Component.InstallationDate",
            "cdx": "GSA.09.Asset.InstallationDate",
            "guid": "AB317582-D3F4-4437-B9BC-2671B0965AD4"
        },
        {
            "type": "Asset",
            "cobie": "COBie.Component.WarrantyStartDate" ,
            "cdx": "GSA.09.Asset.WarrantyStartDate",
            "guid": "B46565AA-C5A7-4E6E-AC31-5A2D291846BC"
        },
        {
            "type": "Asset",
            "cobie": "COBie.System.Name",
            "cdx": "GSA.06.Asset.SystemName",
            "guid": "C9711AA6-0A1C-4299-9A46-429F9D6E3517"
        },
        {
            "type": "Asset",
            "cobie": "COBie.System.Category",
            "cdx": "GSA.06.Asset.OmniClass.SystemCategory",
            "guid": "0CECE43C-75DB-4EFE-85BE-90E4056AB09" 
        },
    ]

    facilityParameters = [
        item for item in cdxParameterMap if item['type'] == "Facility"
    ]

    projectInformations = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_ProjectInformation) \
        .ToElements()
    projectInformation = projectInformations[0]

    rooms = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_Rooms) \
        .WhereElementIsNotElementType() \
        .ToElements
    spaceParameters = [
        item for item in cdxParameterMap if item['type'] == "Space"
    ]

    with revit.Transaction():
        for item in facilityParameters:
            cobieParameter = projectInformation.LookupParameter(item['cobie'])
            cdxParameter = projectInformation.get_Parameter(item['guid'])
            cdxParameter.Set(cobieParameter.AsString())
        
        for room in rooms:
            for item in spaceParameters:
                if COBieIsEnabled(room):
                    cobieParameter = room.LookupParameter(item['cobie'])
                    cdxParameter = room.get_Parameter(item['guid'])
                    cdxParameter.Set(cobieParameter.AsString())
