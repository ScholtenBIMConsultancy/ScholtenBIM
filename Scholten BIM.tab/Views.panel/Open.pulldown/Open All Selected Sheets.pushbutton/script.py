__title__ = "Open Selected Sheets"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.1
Datum    = 20.12.2024
__________________________________________________________________
Description:

Met deze tool kan je geselecteerde sheets uit de project browser in 1x openen ipv per stuk.
__________________________________________________________________
How-to:
-> Zorg dat je een of meerdere sheets in de 
     Project Browser selecteert.
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
from pyrevit import revit, DB, HOST_APP

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Haal geselecteerde sheets op uit de projectbrowser
selected_ids = revit.uidoc.Selection.GetElementIds()

# Controleer of er sheets geselecteerd zijn
if not selected_ids:
    MessageBox.Show("Geen sheets geselecteerd. Selecteer sheets in de Projectbrowser en probeer opnieuw.", "Open Selected Sheets | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    sys.exit()  # Stop het script als er geen sheets zijn geselecteerd

# Itereer door de geselecteerde elementen en open de sheets
sheets_opened = False  # Vlag om bij te houden of er sheets zijn geopend
for sheet_id in selected_ids:
    element = doc.GetElement(sheet_id)
    uidoc.RequestViewChange(element)  # Open de sheet in het hoofdvenster
    sheets_opened = True  # Stel de vlag in op True als er een sheet is geopend

# Toon de messagebox alleen als er sheets zijn geopend
if sheets_opened:
    MessageBox.Show("Geselecteerde sheets zijn geopend.", "Open Selected Sheets | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)