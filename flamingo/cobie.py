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
