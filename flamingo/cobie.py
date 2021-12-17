from Autodesk.Revit import DB
from flamingo.revit import GetViewPhase, GetElementRoom
from flamingo.revit import GetScheduledParameterIds
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
    spaceParameter = element.LookupParameter(parameterName)
    if blankOnly:
        currentValue = spaceParameter.AsString()
        if currentValue is not None and currentValue != "":
            return element
        else:
            print("Blank Value!")
    spaceParameter.Set(value)
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
            if parameter.AsInteger < 1:
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
                GetElementRoom(
                    element, phase, doc=doc
                ).Number,
            )
    return

def COBieComponentSetDescription(view, doc=None):
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
                    selection == "All Values"
                ):
                    elementDescription.Set(description)
                    outElements.append(element)

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
                    .WherePasses(familyInstanceFilter)
                marks[symbolId] = {
                    "count": instances.Count(),
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
    if doc is None:
        doc = HOST_APP.doc
    scheduledParameters = {}
    groupedParameterNames = {}
    for schedule in scheduleList:
        parameterIds = GetScheduledParameterIds(scheduleView, doc=doc)
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