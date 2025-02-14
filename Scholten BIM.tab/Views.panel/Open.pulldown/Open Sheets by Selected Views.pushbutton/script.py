__title__ = "Open Sheets by Selected Views"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.1
Datum    = 20.12.2024
__________________________________________________________________
Description:

Met deze tool kan je de sheets op basis van de geselecteerde views uit de project browser in 1x openen ipv per stuk.
__________________________________________________________________
How-to:
-> Zorg dat je een of meerdere views in de 
     Project Browser selecteerd.
-> Run het script.
__________________________________________________________________
Last update:

- [11.02.2025] - 1.1 PyRevit Forms omgezet naar Windows Forms.
- [20.12.2024] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Tips/Tricks?.
__________________________________________________________________
"""
# Importeer de benodigde Revit API-modules en PyRevit tools
import clr
import sys
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

from pyrevit import revit, DB, HOST_APP

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Sla de geselecteerde elementen op
selected_ids = uidoc.Selection.GetElementIds()

# Controleer of er iets geselecteerd is
if not selected_ids:
    MessageBox.Show("Geen views geselecteerd.", "Open Selected Views | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    sys.exit()  # Stop het script als er geen views zijn geselecteerd
else:
    sheets_to_open = set()

    # Itereer door de selectie
    for id in selected_ids:
        element = doc.GetElement(id)
        
        # Controleer of het element een View is
        if isinstance(element, View):
            # Zoek de sheet waar deze view op staat
            collector = FilteredElementCollector(doc).OfClass(ViewSheet)
            for sheet in collector:
                view_ids = [view_id for view_id in sheet.GetAllPlacedViews()]
                if element.Id in view_ids:
                    sheets_to_open.add(sheet)

    # Open de gevonden sheets
    if sheets_to_open:
        for sheet in sheets_to_open:
            uidoc.RequestViewChange(sheet)
        MessageBox.Show("Er zijn {} sheets geopend.".format(len(sheets_to_open)), "Open Selected Views | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    else:
        MessageBox.Show("Geen bijbehorende sheets gevonden.", "Open Selected Views | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)