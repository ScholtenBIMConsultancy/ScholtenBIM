# -*- coding: utf-8 -*-

__title__ = "Unused Revision Sequences"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 28.03.2025
__________________________________________________________________
Description:

Met deze tool kan je alle niet gebruikte Revision Sequences zien binnen je project.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [28.03.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

-
__________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, RevisionCloud, ViewSheet
from RevitServices.Persistence import DocumentManager
from pyrevit import forms
from pyrevit import script

# Collect current document
doc = __revit__.ActiveUIDocument.Document

# Collect all revision clouds in the document
revision_clouds = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RevisionClouds).WhereElementIsNotElementType().ToElements()

# Collect all revisions in the document
revisions = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType().ToElements()

# Collect all sheets in the document
sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()

# Create a set of used revision IDs from revision clouds
used_revision_ids_clouds = set(cloud.RevisionId for cloud in revision_clouds)

# Create a set of used revision IDs from sheets
used_revision_ids_sheets = set()
for sheet in sheets:
    for revision_id in sheet.GetAllRevisionIds():
        used_revision_ids_sheets.add(revision_id)

# Combine both sets of used revision IDs
used_revision_ids = used_revision_ids_clouds.union(used_revision_ids_sheets)

# Find unused revisions
unused_revisions = [revision for revision in revisions if revision.Id not in used_revision_ids]

# Sort unused revisions by sequence number
unused_revisions_sorted = sorted(unused_revisions, key=lambda rev: rev.SequenceNumber)

# Create a list of unused revisions for the form
revision_items = ["{} - {}".format(rev.SequenceNumber, rev.Description) for rev in unused_revisions_sorted]

# Show the form
selected_revisions = forms.SelectFromList.show(
    revision_items,
    title='Unused Revisions | Scholten BIM Consultancy',
    button_name='Delete Selected',
    multiselect=True
)

# Handle form results
if selected_revisions:
    selected_revision_ids = [unused_revisions_sorted[revision_items.index(rev)].Id for rev in selected_revisions]
    TransactionManager.Instance.EnsureInTransaction(doc)
    for rev_id in selected_revision_ids:
        doc.Delete(rev_id)
    TransactionManager.Instance.TransactionTaskDone()