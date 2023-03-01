# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from flamingo.revit import (
    GetViewPhase,
    GetElementRooms,
    SetParameter,
    GetScheduledParameterIds,
    GetScheduledParameterByName,
    GetSchedulableFields,
    GetParameterValueByName,
)
import json
from pyrevit import forms, HOST_APP, revit, script
import re
from System import Guid
from System.Collections.Generic import List
from xml.etree import ElementTree as ET

LOGGER = script.get_logger()
OUTPUT = script.get_output()

CDX_PARAMETER_MAP = [
    {
        "type": "Facility",
        "cobie": "COBie.Facility.Name",
        "cdx": "GSA.01.Facility.GSABuildingCode",
        "guid": "7704D27B-55E1-4E0D-A323-BE403AEE4CDD",
    },
    {
        "type": "Facility",
        "cobie": "COBie.Facility.Category",
        "cdx": "GSA.01.Facility.OmniClass.FacilityCategory",
        "guid": "A0A94036-DCE0-45DC-B25B-B19632CEEA23",
    },
    {
        "type": "Facility",
        "cobie": "COBie.Facility.ProjectName",
        "cdx": "GSA.01.Facility.ProjectName",
        "guid": "D18B5EA9-0C02-45A4-ACD6-D8887A5BBAD4",
    },
    {
        "type": "Facility",
        "cobie": "COBie.Facility.SiteName",
        "cdx": "GSA.01.Facility.SiteName",
        "guid": "52E4EF7A-08FB-44C5-8935-66BA1C2EE94B",
    },
    {
        "type": "Floor",
        "cobie": "COBie.Floor.Name",
        "cdx": "GSA.01.Floor.Name",
        "guid": "1BB202D3-D9E2-428A-8375-EE6D01A98FFA",
    },
    {
        "type": "Floor",
        "cobie": "COBie.Floor.Category",
        "cdx": "GSA.01.Floor.FloorCategory",
        "guid": "8836ACFF-EA52-41FE-BC49-8E81D1CD8605",
    },
    {
        "type": "Space",
        "cobie": "COBie.Space.Name",
        "cdx": "GSA.02.Space.RoomNumber",
        "guid": "57E67856-3996-4DC3-AD75-4C33A76704A0",
    },
    {
        "type": "Space",
        "cobie": "COBie.Space.Category",
        "cdx": "GSA.03.Space.OmniClass.SpaceCategory",
        "guid": "00DCCF44-54E8-4290-B386-2C6787C06376",
    },
    {
        "type": "Space",
        "cobie": DB.BuiltInParameter.LEVEL_NAME,
        "cdx": "GSA.03.Space.FloorName",
        "guid": "FC13D9B2-31FE-4B58-A19D-F679C3F641CD",
    },
    {
        "type": "Space",
        "cobie": "COBie.Space.Description",
        "cdx": "GSA.03.Space.Description",
        "guid": "DF8F1C12-F287-4360-AFBC-CC92B0B56D2A",
    },
    {
        "type": "Space",
        "cobie": "COBie.Space.RoomTag",
        "cdx": "GSA.05.Space.Signage",
        "guid": "28ECE7CD-8818-436D-AF6C-733D3AE0E69B",
    },
    {
        "type": "Space",
        "cobie": "COBie.Zone.Name",
        "cdx": "GSA.03.Space.ZoneName",
        "guid": "B5DB0243-5AFE-4153-B037-775E21CC57F1",
    },
    {
        "type": "Space",
        "cobie": DB.BuiltInParameter.ROOM_AREA,
        "cdx": "GSA.02.Space.ANSIBOMA.UsableSFCalculation",
        "guid": "C4BF9EEB-6A71-4BFD-B4E9-AF88D4302B25",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.Name",
        "cdx": "GSA.05.Asset.TypeName",
        "guid": "0798BDCD-761D-4942-8FF2-0E7007F40E6A",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.Category",
        "cdx": "GSA.05.Asset.AssetType.Description",
        "guid": "ACD08F9A-6EAB-4981-8388-14A806DB7DFF",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.Description",
        "cdx": "GSA.05.Asset.AssetType.Description",
        "guid": "ACD08F9A-6EAB-4981-8388-14A806DB7DFF",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.AssetType",
        "cdx": "GSA.05.Asset.AssetType.Abbreviation",
        "guid": "37791872-234C-4766-9F31-7A3E72050582",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.Manufacturer",
        "cdx": "GSA.07.Asset.Manufacturer.Email",
        "guid": "639F8842-58BF-48C9-9B00-8C4D2790BA68",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.ModelNumber",
        "cdx": "GSA.07.Asset.Model.Number",
        "guid": "C20C3645-948F-4846-B80E-93049348FA0C",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.WarrantyDurationParts",
        "cdx": "GSA.07.Asset.WarrantyDuration.Parts",
        "guid": "4469578B-15B2-46DF-BA68-45BFA7D53AF7",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.WarrantyDurationLabor",
        "cdx": "GSA.07.Asset.WarrantyDuration.Labor",
        "guid": "C9021772-A5C2-471D-B116-3BD710D0FE31",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.WarrantyDurationUnit",
        "cdx": "GSA.07.Asset.WarrantyDuration.Unit",
        "guid": "4C4306D4-927C-4CD5-8501-6D58096AC4A3",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.ExpectedLife",
        "cdx": "GSA.07.Asset.ExpectedLife",
        "guid": "73F20796-C594-4929-A364-F34CBEB0CCD7",
    },
    {
        "type": "Type",
        "cobie": "COBie.Type.DurationUnit",
        "cdx": "GSA.07.Asset.ExpectedLife.Unit",
        "guid": "E70826DF-0EFA-4F2B-96FA-682DB5E6CDEA",
    },
    {
        "type": "Component",
        "cobie": "COBie.Component.Name",
        "cdx": "GSA.06.Asset.InstanceName",
        "guid": "E21B3065-6702-4973-9E1F-E3F21EF27581",
    },
    {
        "type": "Component",
        "cobie": "COBie.Component.Space",
        "cdx": "GSA.06.Asset.SpaceCode",
        "guid": "E2D4398B-F79D-486A-9B23-81614ED9AA84",
    },
    {
        "type": "Component",
        "cobie": "COBie.Component.Description",
        "cdx": "GSA.05.Asset.Description",
        "guid": "434BE294-6B1A-4B85-A30E-A2986B2AF718",
    },
    {
        "type": "Component",
        "cobie": "COBie.Component.SerialNumber",
        "cdx": "GSA.09.Asset.SerialNumber",
        "guid": "E700395C-92AA-4AE8-86A5-5F3EDC67C944",
    },
    {
        "type": "Component",
        "cobie": "COBie.Component.InstallationDate",
        "cdx": "GSA.09.Asset.InstallationDate",
        "guid": "AB317582-D3F4-4437-B9BC-2671B0965AD4",
    },
    {
        "type": "Component",
        "cobie": "COBie.Component.WarrantyStartDate",
        "cdx": "GSA.09.Asset.WarrantyStartDate",
        "guid": "B46565AA-C5A7-4E6E-AC31-5A2D291846BC",
    },
    {
        "type": "Asset",
        "cobie": "COBie.System.Name",
        "cdx": "GSA.06.Asset.SystemName",
        "guid": "C9711AA6-0A1C-4299-9A46-429F9D6E3517",
    },
    {
        "type": "Asset",
        "cobie": "COBie.System.Category",
        "cdx": "GSA.06.Asset.OmniClass.SystemCategory",
        "guid": "0CECE43C-75DB-4EFE-85BE-90E4056AB09",
    },
]

CDX_ANSI_BOMA_CODE_MAP = {
    "Office": "01",
    "Building Common": "02",
    "Floor Common": "03",
    "Vertical Penetration": "04",
    "PBS Specific": "05",
}


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
    LOGGER.debug(
        "COBieParameterIsBlank: element={}, parameter={}".format(
            element.Id, parameterName
        )
    )
    if type(parameterName) is str:
        parameter = element.LookupParameter(parameterName)
    else:
        parameter = element.get_Parameter(parameterName)
    try:
        parameterValue = parameter.AsString()
        if parameterValue is None or parameterValue == "":
            return True
    except Exception as e:
        LOGGER.debug(str(e))
    return


def SetCOBieComponentSpace(element, phase, blankOnly, doc=None):
    """Finds matches a room to every element listed in the COBie.Components
    schedule that has the "COBie" parameter checked. It then assigns the value
    of matched room number to the COBie.Component.Space parameter.

    Args:
        view (DB.ScheduleView): COBie Component Schedule
        blankOnly (bool): If true, only update values that are currently blank
        doc (DB.Document, optional): [description]. Defaults to None.
    """
    LOGGER.debug("SetCOBieComponentSpace")
    if doc is None:
        doc = HOST_APP.doc

    if blankOnly:
        if not COBieParameterIsBlank(element, "COBie.Component.Space"):
            return
    try:
        rooms = GetElementRooms(
            # element=element, phase=doc.Phases[phaseId], offset=1, doc=doc
            element=element,
            phase=phase,
            offset=1,
            doc=doc,
        )
        if not rooms:
            LOGGER.debug("No room found")
            return
        roomNumbers = ", ".join([room.Number for room in rooms])
        if SetCOBieParameter(element, "COBie.Component.Space", roomNumbers):
            LOGGER.info("{} -> {}".format(OUTPUT.linkify(element.Id), roomNumbers))
    except Exception as e:
        LOGGER.warn("Error: {}".format(e))
        return
    LOGGER.debug("SetCOBieComponentSpace: Complete")
    return element


def COBieComponentSetDescription(elements, blankOnly=True, skipGrouped=True, doc=None):
    LOGGER.debug("COBieComponentSetDescription")

    from flamingo.revit import GetAllElementIdsInModelGroups

    doc = doc or HOST_APP.doc

    if blankOnly:
        LOGGER.info("Get elements that have a blank value")
        elements = [
            element
            for element in elements
            if COBieParameterIsBlank(element, "COBie.Component.Description")
        ]

    groupedElementIds = GetAllElementIdsInModelGroups(doc) if skipGrouped else []
    LOGGER.debug("len(groupedElementIds) = {}".format(len(groupedElementIds)))

    outElements = []
    grouped = []
    for element in elements:
        if element.Id in groupedElementIds:
            LOGGER.warn(
                "{} Skipping grouped element".format(OUTPUT.linkify(element.Id))
            )
            grouped.append(element.Id)
            continue
        description = GetParameterValueByName(
            element,
            DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
            asValueString=True,
        )
        LOGGER.debug("description = {}".format(description))
        elementDescription = element.LookupParameter("COBie.Component.Description")
        elementDescription.Set(description.replace("_", " "))
        LOGGER.debug(
            "{} Set Component Description: {}".format(
                OUTPUT.linkify(element.Id), description
            )
        )
        outElements.append(element)
    return outElements, grouped


def COBieComponentAssignMarks(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    elements = (
        DB.FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    with revit.Transaction("Set unassigned instance marks"):
        marks = {}

        for element in elements:
            isCobie = element.LookupParameter("COBie").AsInteger()
            if isCobie != 1:
                continue
            if hasattr(element, "WallType"):
                LOGGER.warn("WallType")
                symbolId = element.WallType.Id
                if symbolId not in marks:
                    instances = element.WallType.GetDependentElements(
                        DB.ElementClassFilter(DB.Wall)
                    )
            elif hasattr(element, "Symbol"):
                LOGGER.warn("Symbol")
                symbolId = element.Symbol.Id
                if symbolId not in marks:
                    familyInstanceFilter = DB.FamilyInstanceFilter(doc, symbolId)
                    instances = (
                        DB.FilteredElementCollector(doc, view.Id)
                        .WherePasses(familyInstanceFilter)
                        .ToElements()
                    )
            else:
                continue
            LOGGER.warn("len(instances) = {}".format(len(instances)))
            marks[symbolId] = {"count": len(instances), "marks": []}
            markParameter = element.get_Parameter(DB.BuiltInParameter.DOOR_NUMBER)
            if not markParameter:
                markParameter = element.get_Parameter(
                    DB.BuiltInParameter.ALL_MODEL_MARK
                )
            mark = markParameter.AsString()
            LOGGER.warn("mark = {}".format(mark))
            if not mark or mark == "":
                for i in range(marks[symbolId]["count"]):
                    mark = "{:03d}".format(i + 1)
                    if mark not in marks[symbolId]["marks"]:
                        marks[symbolId]["marks"].append(mark)
                        markParameter.Set(mark)
                        break


def COBieParameterBlankOut(scheduleList, elements=None, doc=None):
    """Clears values for selected COBie paramaters. This is helpful for values
    that are not required at particular time points, like Manufacture at the CD
    project phase.

    Args:
        scheduleList (list[DB.ViewSchedule]): List of schedules to operate on
        elements (list[DB.Elements]): List of Revit elements to operate on. If not
            included, all elements that have the parameter will be blanked out.
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
        sorted(scheduledParameters.keys()),
        multiselect=True,
        group_selector_title="COBie Schedule",
        button_name="Select Parameters",
    )

    if not parameterNamesOut:
        script.exit()

    if elements:
        elementIds = [element.Id for element in elements]
        elementTypes = [GetElementSymbol(element) for element in elements]
        elementTypeIds = set(element.Id for element in elementTypes)
    else:
        elementIds = None
        elementTypeIds = None
    bindings = doc.ParameterBindings
    with revit.Transaction("COBie Parameter Blank Out"):
        for parameterName in parameterNamesOut:
            parameterId = scheduledParameters[parameterName]
            parameter = doc.GetElement(parameterId)
            definition = parameter.GetDefinition()
            parameterBinding = bindings.get_Item(definition)
            parameterValueProvider = DB.ParameterValueProvider(parameterId)
            filterStringRule = DB.FilterStringRule(
                parameterValueProvider, DB.FilterStringGreater(), "", True
            )
            elementParameterFilter = DB.ElementParameterFilter(filterStringRule)
            if type(parameterBinding) is DB.TypeBinding:
                collector = DB.FilteredElementCollector(doc).WhereElementIsElementType()
                if elementTypeIds:
                    collector.WherePasses(
                        DB.ElementIdSetFilter(List[DB.ElementId](elementTypeIds))
                    )
            else:
                collector = DB.FilteredElementCollector(
                    doc
                ).WhereElementIsNotElementType()
                if elementIds:
                    collector.WherePasses(
                        DB.ElementIdSetFilter(List[DB.ElementId](elementIds))
                    )
            elements = collector.WherePasses(elementParameterFilter).ToElements()

            for element in elements:
                elementParameter = element.get_Parameter(parameter.GuidValue)
                elementParameter.Set("")


def COBieUncheckAll(typeViewSchedule, componentViewSchedule, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    print("Collecting COBie Type elements")
    cobieTypeParameter = GetScheduledParameterByName(
        typeViewSchedule, "COBie.Type", doc
    )
    cobieTypes = (
        DB.FilteredElementCollector(doc)
        .WhereElementIsElementType()
        .WherePasses(
            DB.ElementParameterFilter(DB.HasNoValueFilterRule(cobieTypeParameter.Id))
        )
        .ToElements()
    )

    print("len(cobieTypes) = {}".format(len(cobieTypes)))
    print("")
    print("Collecting COBie Instance elements")
    cobieParameter = GetScheduledParameterByName(componentViewSchedule, "COBie", doc)
    hasNoValueFilterRule = DB.HasNoValueFilterRule(cobieParameter.Id)
    elementParameterFilter = DB.ElementParameterFilter(hasNoValueFilterRule)
    cobieInstances = (
        DB.FilteredElementCollector(doc)
        .WhereElementIsNotElementType()
        .WherePasses(
            DB.ElementParameterFilter(DB.HasNoValueFilterRule(cobieParameter.Id))
        )
        .ToElements()
    )
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


def COBieParameterVaryByGroup(typeViewSchedule, doc=None):
    cobieTypeParameter = GetScheduledParameterByName(
        typeViewSchedule, "COBie.Type", doc
    )
    cobieTypes = (
        DB.FilteredElementCollector(doc)
        .WhereElementIsElementType()
        .WherePasses(
            DB.ElementParameterFilter(DB.HasValueFilterRule(cobieTypeParameter.Id))
        )
        .ToElements()
    )
    print(len(cobieTypes))
    doc = doc or HOST_APP.doc
    return


def COBieIsEnabled(element):
    cobieGuid = Guid("a4a71d65-98ff-466f-9c70-d8d281aae297")
    parameter = element.get_Parameter(cobieGuid)
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
    cobieGuid = Guid("a4a71d65-98ff-466f-9c70-d8d281aae297")
    parameter = element.get_Parameter(cobieGuid)
    parameter.Set(True)


def COBieTypeEnable(element):
    cobieTypeGuid = Guid("1691beae-d724-4a70-b6f6-7abd114e9dda")
    parameter = element.get_Parameter(cobieTypeGuid)
    parameter.Set(True)


def GetElementSymbol(element):
    LOGGER.debug("GetElementSymbol")
    if hasattr(element, "WallType"):
        return element.WallType
    if hasattr(element, "CurtainSystemType"):
        return element.CurtainSystemType
    elif hasattr(element, "Symbol"):
        return element.Symbol
    elif hasattr(element, "TypeId"):
        return element.Document.GetElement(element.TypeId)
    LOGGER.warn("{} Unable to get symbol".format(OUTPUT.linkify(element.Id)))
    return


def GetCOBieSpaceElements(doc=None, phase=None, cobieParameterId=None):
    LOGGER.debug("GetCOBieTypeElements")
    doc = doc or HOST_APP.doc
    if cobieParameterId is None:
        sharedParameters = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.SharedParameterElement)
            .ToElements()
        )
        cobieGuid = Guid("a4a71d65-98ff-466f-9c70-d8d281aae297")
        cobieParameter = [
            sharedParameter
            for sharedParameter in sharedParameters
            if sharedParameter.GuidValue == cobieGuid
        ][0]
        cobieParameterId = cobieParameter.Id
    LOGGER.debug("cobieTypeParameterId = {}".format(cobieParameterId))
    parameterValueProvider = DB.ParameterValueProvider(cobieParameterId)
    filterIntegerRule = DB.FilterIntegerRule(
        parameterValueProvider, DB.FilterNumericEquals(), 1
    )
    elementParameterFilter = DB.ElementParameterFilter(filterIntegerRule)
    multiCategoryFilter = DB.ElementMulticategoryFilter(
        List[DB.BuiltInCategory](
            [DB.BuiltInCategory.OST_Rooms, DB.BuiltInCategory.OST_MEPSpaces]
        )
    )
    elements = (
        DB.FilteredElementCollector(doc)
        .WherePasses(multiCategoryFilter)
        .WherePasses(elementParameterFilter)
        .ToElements()
    )

    if phase:
        LOGGER.debug("Filtering by new on phaseId: {}".format(phase.Id.IntegerValue))
        elements = [
            element
            for element in elements
            if GetParameterValueByName(element, DB.BuiltInParameter.ROOM_PHASE_ID)
            == phase.Id
        ]

    LOGGER.debug("COBie Space Element Count: {}".format(len(elements)))
    return elements


def GetCOBieTypeElements(doc=None, cobieTypeParameterId=None):
    LOGGER.debug("GetCOBieTypeElements")
    doc = doc or HOST_APP.doc
    if cobieTypeParameterId is None:
        sharedParameters = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.SharedParameterElement)
            .ToElements()
        )
        cobieTypeGuid = Guid("1691beae-d724-4a70-b6f6-7abd114e9dda")
        cobieTypeParameter = [
            sharedParameter
            for sharedParameter in sharedParameters
            if sharedParameter.GuidValue == cobieTypeGuid
        ][0]
        cobieParameterId = cobieTypeParameter.Id
    LOGGER.debug("cobieTypeParameterId = {}".format(cobieParameterId))
    parameterValueProvider = DB.ParameterValueProvider(cobieParameterId)
    filterIntegerRule = DB.FilterIntegerRule(
        parameterValueProvider, DB.FilterNumericEquals(), 1
    )
    elementParameterFilter = DB.ElementParameterFilter(filterIntegerRule)
    elements = (
        DB.FilteredElementCollector(doc)
        .WhereElementIsElementType()
        .WherePasses(elementParameterFilter)
        .ToElements()
    )
    LOGGER.debug("COBie Type Element Count: {}".format(len(elements)))
    return elements


def GetCOBieComponentElements(doc=None, phase=None, cobieParameterId=None):
    LOGGER.debug("GetCOBieComponentElements")
    doc = doc or HOST_APP.doc
    if cobieParameterId is None:
        sharedParameters = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.SharedParameterElement)
            .ToElements()
        )
        cobieGuid = Guid("a4a71d65-98ff-466f-9c70-d8d281aae297")
        cobieParameter = [
            sharedParameter
            for sharedParameter in sharedParameters
            if sharedParameter.GuidValue == cobieGuid
        ][0]
        cobieParameterId = cobieParameter.Id
    LOGGER.debug("cobieParameterId = {}".format(cobieParameterId))
    parameterValueProvider = DB.ParameterValueProvider(cobieParameterId)
    filterIntegerRule = DB.FilterIntegerRule(
        parameterValueProvider, DB.FilterNumericEquals(), 1
    )
    elementParameterFilter = DB.ElementParameterFilter(filterIntegerRule)
    elementMulticategoryFilter = DB.ElementMulticategoryFilter(
        List[DB.BuiltInCategory](
            [DB.BuiltInCategory.OST_Levels, DB.BuiltInCategory.OST_Rooms]
        ),
        True,
    )
    collection = (
        DB.FilteredElementCollector(doc)
        .WhereElementIsNotElementType()
        .WherePasses(elementMulticategoryFilter)
        .WherePasses(elementParameterFilter)
    )
    if phase:
        LOGGER.debug("Filtering by new on phaseId: {}".format(phase.Id.IntegerValue))
        elementPhaseStatusFilter = DB.ElementPhaseStatusFilter(
            phase.Id, DB.ElementOnPhaseStatus.New
        )
        collection = collection.WherePasses(elementPhaseStatusFilter)

    elements = collection.ToElements()

    LOGGER.debug("COBie Component Element Count: {}".format(len(elements)))
    return elements


def COBieComponentAutoSelect(view, doc=None):
    if doc is None:
        doc = HOST_APP.doc

    elements = (
        DB.FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    familySymbolIds = set()
    for element in elements:
        try:
            symbol = GetElementSymbol(element)
            if COBieTypeIsEnabled(symbol):
                familySymbolIds.add(symbol.Id)
        except Exception as e:
            print("{}: {}".format(element.Id.IntegerValue, e))

    componentView = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSchedule)
        .WhereElementIsNotElementType()
        .Where(lambda x: x.Name == "COBie.Component")
        .First()
    )
    componentElements = DB.FilteredElementCollector(doc, componentView.Id).ToElements()
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
            symbolInstances = (
                DB.FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WherePasses(DB.FamilyInstanceFilter(doc, symbolId))
                .ToElements()
            )
            for instance in symbolInstances:
                print("instance.Group = {}".format(instance.Group))
                parameter = instance.LookupParameter("COBie")
                parameter.Set(1)
    return


def COBieSpaceSetDescription(view, blankOnly=True, doc=None):
    doc = doc or HOST_APP.doc
    rooms = (
        DB.FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    roomsOut = []
    for room in rooms:
        if not COBieIsEnabled(room):
            continue
        roomName = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString()
        spaceDescription = room.LookupParameter("COBie.Space.Description")
        spaceDescriptionValue = spaceDescription.AsString()
        if (
            spaceDescriptionValue == ""
            or spaceDescriptionValue == None
            or not blankOnly
        ):
            spaceDescription.Set(roomName)
            roomsOut.append(room)
    return roomsOut


def SetCDXParameterFromCOBie(element, cdxGuid, COBieParameterName):
    if type(cdxGuid) is not Guid:
        cdxGuid = Guid(cdxGuid)
    if type(COBieParameterName) is str:
        cobieParameter = element.LookupParameter(COBieParameterName)
    else:
        cobieParameter = element.get_Parameter(COBieParameterName)
    cdxParameter = element.get_Parameter(cdxGuid)
    if cobieParameter is None:
        return
    if cdxParameter is None:
        return
    try:
        if cobieParameter.StorageType == DB.StorageType.Double:
            value = cobieParameter.AsDouble()
        elif cobieParameter.StorageType == DB.StorageType.Integer:
            value = cobieParameter.AsInteger()
        else:
            value = cobieParameter.AsString()
        cdxParameter.Set(value or "")
    except (AttributeError, TypeError) as e:
        print("\tError:{}".format(e))
    return cdxParameter


def COBieParametersToCDX(doc=None):
    doc = doc or HOST_APP.doc

    # Parameter Dictionary
    facilityParameters = [
        item for item in CDX_PARAMETER_MAP if item["type"] == "Facility"
    ]

    projectInformation = doc.ProjectInformation

    rooms = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    spaceParameters = [item for item in CDX_PARAMETER_MAP if item["type"] == "Space"]

    zoneList = GetCOBieZones(doc)

    levels = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Levels)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    floorParameters = [item for item in CDX_PARAMETER_MAP if item["type"] == "Floor"]

    cobieTypes = revit.query.get_elements_by_parameter("COBie.Type", 1, doc=doc)
    typeParameters = [item for item in CDX_PARAMETER_MAP if item["type"] == "Type"]
    componentParameters = [
        item for item in CDX_PARAMETER_MAP if item["type"] == "Component"
    ]

    with revit.Transaction("Set CDX from COBie"):
        SetParameter(
            projectInformation,
            "07c80000-aef8-4fb1-8e88-994372265bc2",
            HOST_APP.version_name,
        )

        unitsParameter = projectInformation.get_Parameter(
            Guid("99b55570-50da-4f79-9155-b4e41adc0283")
        )
        unitsParameter.Set(
            "Imperial" if doc.DisplayUnitSystem == DB.DisplayUnit.IMPERIAL else "Metric"
        )

        for item in facilityParameters:
            SetCDXParameterFromCOBie(
                element=projectInformation,
                cdxGuid=item["guid"],
                COBieParameterName=item["cobie"],
            )

        for zone in zoneList:
            zoneName = zone.attrib["Name"]
            for space in zone:
                if True:
                    room = doc.GetElement(space.attrib["ID"])
                    if room:
                        zoneParameter = room.get_Parameter(
                            Guid("B5DB0243-5AFE-4153-B037-775E21CC57F1")
                        )
                        zoneParameter.Set(zoneName)

        for room in rooms:
            for item in spaceParameters:
                if COBieIsEnabled(room):
                    cobieName = item["cobie"]
                    SetCDXParameterFromCOBie(
                        element=room,
                        cdxGuid=item["guid"],
                        COBieParameterName=cobieName,
                    )

        for level in levels:
            levelIdMatch = re.findall(r"(\d{1,3})", level.Name)
            if levelIdMatch:
                levelId = levelIdMatch[0]
            else:
                continue
            levelIdParameter = level.get_Parameter(
                Guid("12379615-bbe6-46de-9f58-572d56b75142")
            )
            levelIdParameter.Set(levelId)
            for item in floorParameters:
                if COBieIsEnabled(level):
                    cobieName = item["cobie"]
                    SetCDXParameterFromCOBie(
                        element=level,
                        cdxGuid=item["guid"],
                        COBieParameterName=cobieName,
                    )

        for cobieType in cobieTypes:
            for item in typeParameters:
                cobieName = item["cobie"]
                SetCDXParameterFromCOBie(
                    element=cobieType,
                    cdxGuid=item["guid"],
                    COBieParameterName=cobieName,
                )
            instanceIds = cobieType.GetDependentElements(
                DB.ElementClassFilter(DB.FamilyInstance)
            )
            for instanceId in instanceIds:
                element = doc.GetElement(instanceId)
                if COBieIsEnabled(element):
                    for item in componentParameters:
                        cobieName = item["cobie"]
                        SetCDXParameterFromCOBie(
                            element=element,
                            cdxGuid=item["guid"],
                            COBieParameterName=cobieName,
                        )

        return


def GetCDXCrosswalkData():
    import sys

    libraryPath = [path for path in sys.path if path.endswith("\\flamingo.lib")]
    assert libraryPath, "Library Path not found in sys.path"
    filePath = "{}\\flamingo\\cobie_cdx_crosswalk.json".format(libraryPath[0])
    with open(filePath, "r") as f:
        cdxCrosswalk = json.load(f)
    return cdxCrosswalk


def CDXParametersToCOBie(cdxCrosswalkData=None, doc=None):
    from flamingo.extensible_storage import SetSchemaData
    from datetime import datetime

    doc = doc or HOST_APP.doc

    # Parameter Dictionary
    cdxCrosswalkData = cdxCrosswalkData or GetCDXCrosswalkData()

    sharedParameters = (
        DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement).ToElements()
    )
    sharedParametersByGuid = {
        str(sharedParameter.GuidValue).lower(): sharedParameter.Id
        for sharedParameter in sharedParameters
    }

    spacesByZoneName = {}
    for parameterName, data in cdxCrosswalkData.items():
        message = ""
        cdxGuid = data["CDXGuid"].lower()
        cobieName = data["COBieName"]
        OUTPUT.print_md("## {}".format(parameterName))
        OUTPUT.print_md("Matches {}".format(cobieName))
        if cdxGuid not in sharedParametersByGuid:
            message = "{}\n- Parameter not in model".format(message)
            continue
        cobieGuid = data["COBieGuid"]
        if not cobieGuid and parameterName != "GSA.03.Space.ZoneName":
            message = "{}\n- COBie parameter not found ({})".format(message, cobieGuid)
            continue
        parameterId = sharedParametersByGuid[cdxGuid]
        elements = (
            DB.FilteredElementCollector(doc)
            .WherePasses(DB.ElementParameterFilter(DB.HasValueFilterRule(parameterId)))
            .ToElements()
        )
        for element in elements:
            elementIdBold = "**{}**".format(element.Id)
            value = GetParameterValueByName(element, Guid(cdxGuid))
            if parameterName == "GSA.03.Space.ZoneName":
                message = "{}\n- {}: Collecting zone info".format(
                    message, elementIdBold
                )
                if not COBieIsEnabled(element):
                    continue
                if value in spacesByZoneName:
                    spaceList = spacesByZoneName[value]
                    spacesByZoneName[value].append(element.UniqueId)
                else:
                    spacesByZoneName[value] = [element.UniqueId]
                continue
            oldValue = GetParameterValueByName(element, Guid(cobieGuid))
            if not value and not oldValue:
                message = "{}\n- {}: Values are blank".format(message, elementIdBold)
            elif value == oldValue:
                message = '{}\n- {}: Values are equal ("{}")'.format(
                    message, elementIdBold, value
                )
            else:
                message = '{}\n- {}: "{}" -> "{}"'.format(
                    message, elementIdBold, value, oldValue
                )
                try:
                    SetParameter(element, Guid(cobieGuid), value)
                except Exception as e:
                    LOGGER.warn(
                        "Error setting value to {} for {}: {}".format(
                            cobieName, element.Id, e
                        )
                    )
        OUTPUT.print_md(message)

    if spacesByZoneName:
        OUTPUT.print_md("## Building COBie zone data")
        currentZoneXML = GetCOBieZones(doc)
        dateTime = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        LOGGER.warn("dateTime = {}".format(dateTime))
        currentZones = {}
        for zone in currentZoneXML:
            currentZones[zone.attrib["Name"]] = {
                "ID": zone.attrib["ID"],
                "ExternalID": zone.attrib["ExternalID"],
                "Description": zone.attrib["Description"],
                "CreatedBy": zone.attrib["CreatedBy"],
                "CreatedOn": zone.attrib["CreatedOn"],
                "Category": zone.attrib["Category"],
            }

        zoneXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        zoneXML = "{}<ZoneList>".format(zoneXML)
        for zoneName, spaceGuids in spacesByZoneName.items():
            if zoneName in currentZones:
                currentZoneData = currentZones[zoneName]
                zoneXML = (
                    '{zoneXML}<Zone ID="{id}" ExternalID="{externalId}" Name="{name}" '
                    'Description="{description}" CreatedBy="{createdBy}" '
                    'CreatedOn="{createdOn}" Category="{category}">'.format(
                        zoneXML=zoneXML,
                        id=currentZoneData["ID"],
                        externalId=currentZoneData["ExternalID"],
                        name=zoneName,
                        description=currentZoneData['Description'],
                        createdBy=currentZoneData["CreatedBy"],
                        createdOn=currentZoneData["CreatedOn"],
                        category=currentZoneData["Category"],
                    )
                )
            else:
                newGuid = Guid.NewGuid()
                zoneXML = (
                    '{}<Zone ID="{}" ExternalID="{}" Name="{}" '
                    'Description="{}" CreatedBy="{}" CreatedOn="{}" '
                    'Category="{}">'.format(
                        zoneXML,
                        newGuid,
                        newGuid,
                        zoneName,
                        zoneName,
                        "mpfammatter@ks.partners",
                        dateTime,
                        "Circulation Zone",
                    )
                )
            for spaceGuid in spaceGuids:
                zoneXML = '{}<Space ID="{}" />'.format(zoneXML, spaceGuid)
            zoneXML = "{}</Zone>".format(zoneXML)
        zoneXML = "{}</ZoneList>".format(zoneXML)
        print(zoneXML)
        zoneSchemaGUID = "e0fc673a-2f54-4f88-b168-186716faaff4"
        zoneSchema = DB.ExtensibleStorage.Schema.Lookup(Guid(zoneSchemaGUID))
        projectInformation = doc.ProjectInformation
        try:
            zoneSchema = DB.ExtensibleStorage.Schema.Lookup(Guid(zoneSchemaGUID))
            SetSchemaData(zoneSchema, "Zones", projectInformation, zoneXML)
        except Exception as e:
            LOGGER.warn("Unable to save COBie zone data: {}".format(e))

    return


def GetCOBieZones(doc=None):
    projectInformation = doc.ProjectInformation
    zoneSchemaGUID = "e0fc673a-2f54-4f88-b168-186716faaff4"
    try:
        zoneSchema = DB.ExtensibleStorage.Schema.Lookup(Guid(zoneSchemaGUID))
        zoneXML = revit.query.get_schema_field_values(projectInformation, zoneSchema)[
            "Zones"
        ]
        return ET.fromstring(zoneXML)
    except Exception as e:
        LOGGER.warn("Unable to acquire COBie zone data: {}".format(e))
    return


def GetCDXSpaceBOMAValues(doc=None):
    doc = doc or HOST_APP.doc
    bomaData = []

    rooms = (
        DB.FilteredElementCollector(doc)
        .OfCategory(DB.BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for room in rooms:
        if COBieIsEnabled(room):
            bomaNameParameter = room.get_Parameter(
                Guid("B0C2E838-BE22-44C4-8B0D-F9FAC2D9DAFD")
            )
            bomaCodeParameter = room.get_Parameter(
                Guid("EE6CECDF-20AA-4A8A-BC49-69C2E230BE52")
            )
            if bomaNameParameter and bomaCodeParameter:
                bomaData.append(
                    {
                        "roomElement": room,
                        "roomNumber": room.Number,
                        "roomName": revit.query.get_name(room),
                        "bomaName": bomaNameParameter.AsString(),
                        "bomaCode": bomaCodeParameter.AsString(),
                    }
                )
    return bomaData


def SetCDXSpaceBOMAValue(room, ansi_boma_name, ansi_boma_code_map=None):
    ansi_boma_code_map = ansi_boma_code_map or CDX_ANSI_BOMA_CODE_MAP
    bomaNameParameter = room.get_Parameter(Guid("B0C2E838-BE22-44C4-8B0D-F9FAC2D9DAFD"))
    bomaCodeParameter = room.get_Parameter(Guid("EE6CECDF-20AA-4A8A-BC49-69C2E230BE52"))
    if bomaNameParameter and bomaCodeParameter:
        bomaNameParameter.Set(ansi_boma_name)
        bomaCodeParameter.Set(ansi_boma_code_map[ansi_boma_name])
    return room


def CreateScheduleView(viewName, category, parameterGUIDList=None, doc=None):
    doc = doc or HOST_APP.doc
    view = DB.ViewSchedule.CreateSchedule(doc, category)
    view.Name = viewName
    scheduleDefinition = view.Definition
    schedulableFields = GetSchedulableFields(view)
    for name, schedulableField in schedulableFields.items():
        if name.startswith("GSA"):
            scheduleDefinition.AddField(schedulableField)
    return view
