__title__ = "Open Selected Views"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.1
Datum    = 20.12.2024
__________________________________________________________________
Description:

Met deze tool kan je geselecteerde views uit de project browser in 1x openen ipv per stuk.
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
from pyrevit import revit, DB, HOST_APP

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Haal geselecteerde views op uit de projectbrowser
selected_ids = revit.uidoc.Selection.GetElementIds()

# Controleer of er views geselecteerd zijn
if not selected_ids:
    MessageBox.Show("Geen views geselecteerd. Selecteer views in de Projectbrowser en probeer opnieuw.", "Open Selected Views | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
else:
    views_opened = False  # Vlag om bij te houden of er views zijn geopend

    # Itereer door de geselecteerde elementen en open de views
    for view_id in selected_ids:
        element = doc.GetElement(view_id)
        if isinstance(element, DB.View) and not element.IsTemplate:  # Controleer of het een view is en geen template
            revit.uidoc.RequestViewChange(element)  # Open de view in het hoofdvenster
            views_opened = True  # Stel de vlag in op True als er een view is geopend

    # Toon de messagebox alleen als er views zijn geopend
    if views_opened:
        MessageBox.Show("Geselecteerde views zijn geopend.", "Open Selected Views | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)