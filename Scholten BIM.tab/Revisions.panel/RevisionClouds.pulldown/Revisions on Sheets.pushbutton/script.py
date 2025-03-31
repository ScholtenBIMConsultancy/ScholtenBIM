# -*- coding: utf-8 -*-

__title__ = "Revisions on Sheet"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 28.03.2025
__________________________________________________________________
Description:

Met deze tool kan je zien op welke sheets de geselecteerde revision te zien is.
__________________________________________________________________
How-to:

-> Run het script.
-> Kies uit je pulldown menu de uit te lezen sequence.
-> Overzicht van sheets.
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
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Revision, ViewSheet
from RevitServices.Persistence import DocumentManager
from pyrevit import forms, output

doc = __revit__.ActiveUIDocument.Document

# Collect all revisions
revisions = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType().ToElements()

# Create a dictionary to store revision sequences and their descriptions
revision_dict = {rev.SequenceNumber: "{} - {}".format(rev.SequenceNumber, rev.Description) for rev in revisions}

# Create a list of revision sequences for the dropdown menu
revision_sequences = list(revision_dict.values())

# Show a dropdown menu to select a revision sequence
selected_seq = forms.SelectFromList.show(revision_sequences, title="Select Revision Sequence", width=300, button_name="Select")

if selected_seq:
    sheets_found = False
    selected_seq_number = int(selected_seq.split(" - ")[0])  # Extract the sequence number from the selected item
    selected_description = selected_seq.split(" - ")[1]  # Extract the description from the selected item
    
    # Collect sheets with the selected revision sequence
    sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    sheet_list = []
    for sheet in sheets:
        revs_on_sheet = sheet.GetAllRevisionIds()
        for rev_id in revs_on_sheet:
            rev = doc.GetElement(rev_id)
            if rev.SequenceNumber == selected_seq_number:
                sheets_found = True
                sheet_list.append(sheet)
    
    output_window = output.get_output()
    
    # Print the selected sequence and description at the top
    output_window.print_md("##**==== Selected Revision Sequence:** {} - {} ====##".format(selected_seq_number, selected_description))
    
    if sheets_found:
        # Sort sheets by sheet number
        sorted_sheets = sorted(sheet_list, key=lambda s: s.SheetNumber)
        # Show sorted sheets without extra space
        for sheet in sorted_sheets:
            sheet_info = "Sheet: {} - {}".format(sheet.SheetNumber, sheet.Name)
            output_window.print_md(sheet_info)
    else:
        output_window.print_md("No sheets found for the selected revision sequence.")