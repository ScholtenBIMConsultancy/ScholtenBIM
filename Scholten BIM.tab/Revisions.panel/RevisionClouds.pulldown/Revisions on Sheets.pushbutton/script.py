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
    
    if sheets_found:
        # Sort sheets by sheet number
        sorted_sheets = sorted(sheet_list, key=lambda s: s.SheetNumber)
        # Show sorted sheets with an "Open" button
        output_window = output.get_output()
        for sheet in sorted_sheets:
            sheet_info = "Sheet: {} - {}".format(sheet.SheetNumber, sheet.Name)
            output_window.print_md("{} Open".format(sheet_info, sheet.Id))
    else:
        output_window.print_md("No sheets found for the selected revision sequence.")