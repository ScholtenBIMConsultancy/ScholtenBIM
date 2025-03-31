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
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, RevisionCloud, ViewSheet, Transaction
from RevitServices.Persistence import DocumentManager
from pyrevit import forms
from pyrevit import script
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

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
    
    # Check if there is more than one revision
    if len(revisions) > 1:
        # Ensure at least one revision remains
        if len(selected_revision_ids) >= len(revisions):
            # Remove the first selected revision from the deletion list
            first_revision_id = selected_revision_ids.pop(0)
            first_revision = next(rev for rev in revisions if rev.Id == first_revision_id)
            
            t = Transaction(doc, 'Delete Selected Revisions')
            try:
                t.Start()
                for rev_id in selected_revision_ids:
                    doc.Delete(rev_id)
                t.Commit()
                # Show MessageBox with message
                MessageBox.Show("Revision '{0} - {1}' has been retained.".format(first_revision.SequenceNumber, first_revision.Description), "Unused Revision Sequences | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            except Exception as e:
                t.RollBack()
                print("Error: {}".format(e))
        else:
            t = Transaction(doc, 'Delete Selected Revisions')
            try:
                t.Start()
                for rev_id in selected_revision_ids:
                    doc.Delete(rev_id)
                t.Commit()
            except Exception as e:
                t.RollBack()
                print("Error: {}".format(e))
    else:
        # Show MessageBox with message
        MessageBox.Show("Cannot delete the only remaining revision.", "Unused Revision Sequences | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)