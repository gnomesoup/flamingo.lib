import sys

from pyrevit.framework import List
from pyrevit import forms
from pyrevit import revit, DB
from pyrevit.revit import query
from pyrevit import script
from Autodesk.Revit.DB import Element as DBElement


LOGGER = script.get_logger()
OUTPUT = script.get_output()

VIEW_TOS_PARAM = DB.BuiltInParameter.VIEW_DESCRIPTION

class Option(forms.TemplateListItem):
    def __init__(self, op_name, default_state=False):
        super(Option, self).__init__(op_name)
        self.state = default_state


class OptionSet:
    def __init__(self):
        self.op_copy_vports = Option("Copy Viewports", True)
        self.op_copy_schedules = Option("Copy Schedules", True)
        self.op_copy_titleblock = Option("Copy Sheet Titleblock", True)
        self.op_copy_revisions = Option("Copy and Set Sheet Revisions", False)
        self.op_copy_placeholders_as_sheets = Option(
            "Copy Placeholders as Sheets", True
        )
        self.op_copy_guides = Option("Copy Guide Grids", True)
        self.op_update_exist_view_contents = Option("Update Existing View Contents")
        # self.op_update_exist_vport_locations = \
        #    Option('Update Existing Viewport Locations')


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    """Handle copy and paste errors."""

    def OnDuplicateTypeNamesFound(self, args):
        """Use destination model types if duplicate."""
        return DB.DuplicateTypeAction.UseDestinationTypes


def get_user_options():
    op_set = OptionSet()
    return_options = forms.SelectFromList.show(
        [getattr(op_set, x) for x in dir(op_set) if x.startswith("op_")],
        title="Select Copy Options",
        button_name="Copy Now",
        multiselect=True,
    )

    if not return_options:
        sys.exit(0)

    return op_set

def get_default_type(source_doc, type_group):
    return source_doc.GetDefaultElementTypeId(type_group)


def find_matching_view(dest_doc, source_view):
    for v in DB.FilteredElementCollector(dest_doc).OfClass(DB.View):
        if v.ViewType == source_view.ViewType and query.get_name(v) == query.get_name(
            source_view
        ):
            if source_view.ViewType == DB.ViewType.DrawingSheet:
                if v.SheetNumber == source_view.SheetNumber:
                    return v
            else:
                return v


def find_guide(guide_name, source_doc):
    # collect guides in dest_doc
    guide_elements = (
        DB.FilteredElementCollector(source_doc)
        .OfCategory(DB.BuiltInCategory.OST_GuideGrid)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    # find guide with same name
    for guide in guide_elements:
        if str(guide.Name).lower() == guide_name.lower():
            return guide


def get_view_contents(dest_doc, source_view):
    view_elements = (
        DB.FilteredElementCollector(dest_doc, source_view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    elements_ids = []
    for element in view_elements:
        if (
            element.Category and element.Category.Name == "Title Blocks"
        ) and not OPTION_SET.op_copy_titleblock:
            continue
        elif (
            isinstance(element, DB.ScheduleSheetInstance)
            and not OPTION_SET.op_copy_schedules
        ):
            continue
        elif isinstance(element, DB.Viewport) or "ExtentElem" in query.get_name(
            element
        ):
            continue
        elif (
            isinstance(element, DB.Element)
            and element.Category
            and "guide" in str(element.Category.Name).lower()
        ):
            continue
        elif (
            isinstance(element, DB.Element)
            and element.Category
            and "views" == str(element.Category.Name).lower()
        ):
            continue
        else:
            elements_ids.append(element.Id)
    return elements_ids


def ensure_dest_revision(src_rev, all_dest_revs, dest_doc):
    # check to see if revision exists
    for rev in all_dest_revs:
        if query.compare_revisions(rev, src_rev):
            return rev

    # if no matching revisions found, create a new revision and return
    LOGGER.warning(
        "Revision could not be found in destination model.\n"
        "Revision Date: {}\n"
        "Revision Description: {}\n"
        "Creating a new revision. Please review revisions "
        "after copying process is finished.".format(
            src_rev.RevisionDate, src_rev.Description
        )
    )
    return revit.create.create_revision(
        description=src_rev.Description,
        by=src_rev.IssuedBy,
        to=src_rev.IssuedTo,
        date=src_rev.RevisionDate,
        doc=dest_doc,
    )


def clear_view_contents(dest_doc, dest_view):
    LOGGER.debug("Removing view contents: {}".format(dest_view.Name))
    elements_ids = get_view_contents(dest_doc, dest_view)

    with revit.Transaction("Delete View Contents", doc=dest_doc):
        for el_id in elements_ids:
            try:
                dest_doc.Delete(el_id)
            except Exception as err:
                continue

    return True


def copy_view_contents(
    activedoc, source_view, dest_doc, dest_view, clear_contents=False
):
    LOGGER.debug(
        "Copying view contents: {} : {}".format(source_view.Name, source_view.ViewType)
    )

    elements_ids = get_view_contents(activedoc, source_view)

    if clear_contents:
        if not clear_view_contents(dest_doc, dest_view):
            return False

    cp_options = DB.CopyPasteOptions()
    cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

    if elements_ids:
        with revit.Transaction("Copy View Contents", doc=dest_doc, swallow_errors=True):
            DB.ElementTransformUtils.CopyElements(
                source_view,
                List[DB.ElementId](elements_ids),
                dest_view,
                None,
                cp_options,
            )

    return True


def copy_view_props(source_view, dest_view):
    dest_view.Scale = source_view.Scale
    dest_view.Parameter[VIEW_TOS_PARAM].Set(
        source_view.Parameter[VIEW_TOS_PARAM].AsString()
    )


def copy_view(activedoc, source_view, dest_doc):
    matching_view = find_matching_view(dest_doc, source_view)
    if matching_view:
        print("\t\t\tView/Sheet already exists in document.")
        if OPTION_SET.op_update_exist_view_contents:
            if not copy_view_contents(
                activedoc, source_view, dest_doc, matching_view, clear_contents=True
            ):
                LOGGER.error(
                    "Could not copy view contents: {}".format(source_view.Name)
                )

        return matching_view

    LOGGER.debug("Copying view: {}".format(source_view.Name))
    new_view = None

    if source_view.ViewType == DB.ViewType.DrawingSheet:
        try:
            LOGGER.debug("Source view is a sheet. " "Creating destination sheet.")

            with revit.Transaction("Create Sheet", doc=dest_doc):
                if not source_view.IsPlaceholder or (
                    source_view.IsPlaceholder
                    and OPTION_SET.op_copy_placeholders_as_sheets
                ):
                    new_view = DB.ViewSheet.Create(
                        dest_doc, DB.ElementId.InvalidElementId
                    )
                else:
                    new_view = DB.ViewSheet.CreatePlaceholder(dest_doc)

                revit.update.set_name(new_view, revit.query.get_name(source_view))
                new_view.SheetNumber = source_view.SheetNumber
        except Exception as sheet_err:
            LOGGER.error("Error creating sheet. | {}".format(sheet_err))
    elif source_view.ViewType == DB.ViewType.DraftingView:
        try:
            LOGGER.debug(
                "Source view is a drafting. " "Creating destination drafting view."
            )

            with revit.Transaction("Create Drafting View", doc=dest_doc):
                new_view = DB.ViewDrafting.Create(
                    dest_doc,
                    get_default_type(dest_doc, DB.ElementTypeGroup.ViewTypeDrafting),
                )
                revit.update.set_name(new_view, revit.query.get_name(source_view))
                copy_view_props(source_view, new_view)
        except Exception as sheet_err:
            LOGGER.error("Error creating drafting view. | {}".format(sheet_err))
    elif source_view.ViewType == DB.ViewType.Legend:
        try:
            LOGGER.debug(
                "Source view is a legend. " "Creating destination legend view."
            )

            first_legend = query.find_first_legend(dest_doc)
            if first_legend:
                with revit.Transaction("Create Legend View", doc=dest_doc):
                    new_view = dest_doc.GetElement(
                        first_legend.Duplicate(DB.ViewDuplicateOption.Duplicate)
                    )
                    revit.update.set_name(new_view, revit.query.get_name(source_view))
                    copy_view_props(source_view, new_view)
            else:
                LOGGER.error(
                    "Destination document must have at least one "
                    "Legend view. Skipping legend."
                )
        except Exception as sheet_err:
            LOGGER.error("Error creating drafting view. | {}".format(sheet_err))

    if new_view:
        copy_view_contents(activedoc, source_view, dest_doc, new_view)

    return new_view


def copy_viewport_types(activedoc, vport_type, vport_typename, dest_doc, newvport):
    dest_vport_typenames = [
        DBElement.Name.GetValue(dest_doc.GetElement(x))
        for x in newvport.GetValidTypes()
    ]

    cp_options = DB.CopyPasteOptions()
    cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

    if vport_typename not in dest_vport_typenames:
        with revit.Transaction(
            "Copy Viewport Types", doc=dest_doc, swallow_errors=True
        ):
            DB.ElementTransformUtils.CopyElements(
                activedoc,
                List[DB.ElementId]([vport_type.Id]),
                dest_doc,
                None,
                cp_options,
            )


def apply_viewport_type(activedoc, vport_id, dest_doc, newvport_id):
    with revit.Transaction("Apply Viewport Type", doc=dest_doc):
        vport = activedoc.GetElement(vport_id)
        vport_type = activedoc.GetElement(vport.GetTypeId())
        vport_typename = DBElement.Name.GetValue(vport_type)

        newvport = dest_doc.GetElement(newvport_id)

        copy_viewport_types(activedoc, vport_type, vport_typename, dest_doc, newvport)

        for vtype_id in newvport.GetValidTypes():
            vtype = dest_doc.GetElement(vtype_id)
            if DBElement.Name.GetValue(vtype) == vport_typename:
                newvport.ChangeTypeId(vtype_id)


def copy_sheet_viewports(activedoc, source_sheet, dest_doc, dest_sheet):
    existing_views = [
        dest_doc.GetElement(x).ViewId for x in dest_sheet.GetAllViewports()
    ]

    for vport_id in source_sheet.GetAllViewports():
        vport = activedoc.GetElement(vport_id)
        vport_view = activedoc.GetElement(vport.ViewId)

        print("\t\tCopying/updating view: {}".format(revit.query.get_name(vport_view)))
        new_view = copy_view(activedoc, vport_view, dest_doc)

        if new_view:
            ref_info = revit.query.get_view_sheetrefinfo(new_view)
            if ref_info and ref_info.sheet_num != dest_sheet.SheetNumber:
                LOGGER.error(
                    'View is already placed on sheet "%s - %s"',
                    ref_info.sheet_num,
                    ref_info.sheet_name,
                )
                continue

            if new_view.Id not in existing_views:
                print("\t\t\tPlacing copied view on sheet.")
                with revit.Transaction("Place View on Sheet", doc=dest_doc):
                    nvport = DB.Viewport.Create(
                        dest_doc, dest_sheet.Id, new_view.Id, vport.GetBoxCenter()
                    )
                if nvport:
                    apply_viewport_type(activedoc, vport_id, dest_doc, nvport.Id)
            else:
                print("\t\t\tView already exists on the sheet.")


def copy_sheet_revisions(activedoc, source_sheet, dest_doc, dest_sheet):
    all_src_revs = query.get_revisions(doc=activedoc)
    all_dest_revs = query.get_revisions(doc=dest_doc)
    revisions_to_set = []

    with revit.Transaction("Copy and Set Revisions", doc=dest_doc):
        for src_revid in source_sheet.GetAdditionalRevisionIds():
            set_rev = ensure_dest_revision(
                activedoc.GetElement(src_revid), all_dest_revs, dest_doc
            )
            revisions_to_set.append(set_rev)

        if revisions_to_set:
            revit.update.update_sheet_revisions(
                revisions_to_set, [dest_sheet], state=True, doc=dest_doc
            )


def copy_sheet_guides(activedoc, source_sheet, dest_doc, dest_sheet):
    # sheet guide
    source_sheet_guide_param = source_sheet.Parameter[
        DB.BuiltInParameter.SHEET_GUIDE_GRID
    ]
    source_sheet_guide_element = activedoc.GetElement(
        source_sheet_guide_param.AsElementId()
    )

    if source_sheet_guide_element:
        if not find_guide(source_sheet_guide_element.Name, dest_doc):
            # copy guides to dest_doc
            cp_options = DB.CopyPasteOptions()
            cp_options.SetDuplicateTypeNamesHandler(CopyUseDestination())

            with revit.Transaction("Copy Sheet Guide", doc=dest_doc):
                DB.ElementTransformUtils.CopyElements(
                    activedoc,
                    List[DB.ElementId]([source_sheet_guide_element.Id]),
                    dest_doc,
                    None,
                    cp_options,
                )

        dest_guide = find_guide(source_sheet_guide_element.Name, dest_doc)
        if dest_guide:
            # set the guide
            with revit.Transaction("Set Sheet Guide", doc=dest_doc):
                dest_sheet_guide_param = dest_sheet.Parameter[
                    DB.BuiltInParameter.SHEET_GUIDE_GRID
                ]
                dest_sheet_guide_param.Set(dest_guide.Id)
        else:
            LOGGER.error(
                "Error copying and setting sheet guide for sheet {}".format(
                    source_sheet.Name
                )
            )


def copy_sheet(activedoc, source_sheet, dest_doc, optionSet):
    LOGGER.debug(
        "Copying sheet {} to document {}".format(source_sheet.Name, dest_doc.Title)
    )
    print("\tCopying/updating Sheet: {}".format(source_sheet.Name))
    with revit.TransactionGroup("Import Sheet", doc=dest_doc):
        LOGGER.debug("Creating destination sheet...")
        new_sheet = copy_view(activedoc, source_sheet, dest_doc)

        if new_sheet:
            if not new_sheet.IsPlaceholder:
                if optionSet.op_copy_vports:
                    LOGGER.debug("Copying sheet viewports...")
                    copy_sheet_viewports(activedoc, source_sheet, dest_doc, new_sheet)
                else:
                    print("Skipping viewports...")

                if optionSet.op_copy_guides:
                    LOGGER.debug("Copying sheet guide grids...")
                    copy_sheet_guides(activedoc, source_sheet, dest_doc, new_sheet)
                else:
                    print("Skipping sheet guides...")

            if optionSet.op_copy_revisions:
                LOGGER.debug("Copying sheet revisions...")
                copy_sheet_revisions(activedoc, source_sheet, dest_doc, new_sheet)
            else:
                print("Skipping revisions...")

        else:
            LOGGER.error("Failed copying sheet: {}".format(source_sheet.Name))
