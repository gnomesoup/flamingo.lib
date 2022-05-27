# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
import clr
from flamingo.revit import OpenDetached
from os import path
from pyrevit import HOST_APP, revit, forms, PyRevitException
from pyrevit.coreutils.logger import get_logger
from pyrevit.revit import ensure
import sys
import System

clr.ImportExtensions(System.Linq)


LOGGER = get_logger(__name__)


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
        # if element.IsPhaseCreatedValid(phaseId):
        if element.CreatedPhaseId != phaseId:
            element.CreatedPhaseId = phaseId
            return element
    except Exception as e:
        LOGGER.debug("Error setting phase created: {}".format(e))
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
        # if element.IsPhaseDemolishedValid(phaseId):
        if element.DemolishedPhaseId:
            element.DemolishedPhaseId = phaseId
            return element
    except Exception as e:
        LOGGER.debug("Error setting phase demo'd: {}".format(e))
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
        elementWorkset = element.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM)
        elementWorksetId = element.WorksetId.IntegerValue
        if elementWorksetId != worksetId.IntegerValue and not elementWorkset.IsReadOnly:
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
        viewDrafting = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.ViewDrafting)
            .WhereElementIsNotElementType()
            .ToElements()
            .Where(lambda x: x.Name == viewname)
            .First()
        )
    except:
        draftingView = (
            DB.FilteredElementCollector(doc)
            .OfClass(DB.ViewFamilyType)
            .ToElements()
            .Where(lambda x: x.ViewFamily == DB.ViewFamily.Drafting)
            .First()
        )
        viewDrafting = DB.ViewDrafting.Create(doc, draftingView.Id)
        viewDrafting.Name = viewname
        viewDrafting.Scale = 1
        # TODO Add browser sort data
    return viewDrafting


def EnsureLibraryDoc(documentName, revitVersion=None):
    revitVersion = revitVersion or HOST_APP.version
    docs = HOST_APP.docs
    libraryPath = [path for path in sys.path if path.endswith("\\flamingo.lib")]
    LOGGER.debug("libraryPath = {}".format(libraryPath))
    documentPath = "{}\\{}\\{}.rvt".format(libraryPath[0], revitVersion, documentName)
    libraryDoc = None
    for doc in docs:
        if doc.PathName == documentPath:
            libraryDoc = doc
            LOGGER.debug("Found opened library document")
            break
    if libraryDoc is None:
        libraryDoc = OpenDetached(documentPath)
        LOGGER.debug("Opening library document detached.")
    return libraryDoc


def EnsureLibraryFamily(familyName, revitVersion=None, doc=None):
    if doc is None:
        doc = HOST_APP.doc
    if revitVersion is None:
        revitVersion = HOST_APP.version
    try:
        libraryPath = [path for path in sys.path if path.endswith("\\flamingo.lib")]
        family = ensure.ensure_family(
            familyName,
            family_file="{}\\{}\\{}.rfa".format(
                libraryPath[0],
                revitVersion,
                familyName,
            ),
            doc=doc,
        )
        return family if type(family) == DB.AnnotationSymbolType else family[0]
    except:
        forms.alert(
            msg="Could not load required note placeholder family", exitscript=True
        )
        return None
