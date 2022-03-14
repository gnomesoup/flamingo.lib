# -*- coding: utf-8 -*-
from pyrevit import clr, DB, HOST_APP, revit, forms
from pyrevit.revit import ensure
import sys
import System
clr.ImportExtensions(System.Linq)

def set_element_phase_created(
    element,
    phaseId,
    doc=None,
):
    """
    Sets the phase the element was created 

    Args:
        element (Autodesk.DB.Element): element to assign phase created
        phaseId (Autodesk.DB.Phase.Id): Id of the phase to assign
        doc (Autodesk.DB.Document): Document of the element
    
    Returns:
        Autodesk.DB.Element: element if successful, or None if error
    """

    if doc is None:
        doc = HOST_APP.doc
    try:
        if element.IsPhaseCreatedValid(phaseId):
            if element.CreatedPhaseId != phaseId:
                element.CreatedPhaseId = phaseId
                return element
    except Exception:
        return None
    return None

def set_element_phase_demolished(
    element,
    phaseId,
    doc=None,
):
    """
    Sets the phase the element was demolished 

    Args:
        element (Autodesk.DB.Element): element to assign phase demolished
        phaseId (Autodesk.DB.Phase.Id): Id of the phase to assign
        doc (Autodesk.DB.Document): Document of the element
    
    Returns:
        Autodesk.DB.Element: element if successful, or None if error
    """

    if doc is None:
        doc = HOST_APP.doc
    try:
        if element.IsPhaseDemolishedValid(phaseId):
            if element.DemolishedPhaseId:
                    element.DemolishedPhaseId = phaseId
                    return element
    except Exception as e:
        return None
    return None

def set_element_workset(
    element,
    worksetId,
    doc=None,
):
    """
    Sets the workset of the element

    Args:
        element (Autodesk.DB.Element): element to assign workset
        worksetId (Autodesk.DB.Phase.Id): Id of the workset to assign
        doc (Autodesk.DB.Document): Document of the element
    
    Returns:
        Autodesk.DB.Element: element if successful, or None if error
    """
    if doc is None:
        doc = HOST_APP.doc
    try:
        elementWorkset = element.get_Parameter(
            DB.BuiltInParameter.ELEM_PARTITION_PARAM)
        elementWorksetId = element.WorksetId.IntegerValue
        if (
            elementWorksetId != worksetId.IntegerValue and
            not elementWorkset.IsReadOnly
        ):
            elementWorkset.Set(worksetId.IntegerValue)
            return element
    except Exception as e:
        print("Set workset error: " + str(e))
    return None

def EnsureNoteBlockView(viewname, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    viewname = "«" + viewname + "» KS Placeholder"
    try:
        viewDrafting = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewDrafting)\
            .WhereElementIsNotElementType()\
            .ToElements()\
            .Where(lambda x: x.Name == viewname)\
            .First()
    except:
        draftingView = DB.FilteredElementCollector(doc)\
            .OfClass(DB.ViewFamilyType)\
            .ToElements()\
            .Where(lambda x: x.ViewFamily == DB.ViewFamily.Drafting)\
            .First()
        viewDrafting = DB.ViewDrafting.Create(doc, draftingView.Id)
        viewDrafting.Name = viewname
        viewDrafting.Scale = 1
        #TODO Add browser sort data
    return viewDrafting

def EnsureLibraryFamily(familyName, revitVersion = None, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    if revitVersion is None:
        revitVersion = HOST_APP.version
    try:
        libraryPath = [
            path for path in sys.path
            if path.endswith("\\KSP.lib".format(revitVersion))
        ]
        family = ensure.ensure_family(
            familyName,
            family_file="{}\\{}\\{}.rfa".format(
                libraryPath[0],
                revitVersion,
                familyName,
            ),
            doc=doc
        )
        return family if type(family) == DB.AnnotationSymbolType else family[0]
    except:
        forms.alert(
            msg="Could not load required note placeholder family",
            exitscript=True
        )
        return None