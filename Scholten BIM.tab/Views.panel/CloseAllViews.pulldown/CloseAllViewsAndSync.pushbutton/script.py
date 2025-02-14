# -*- coding: utf-8 -*-

__title__ = "Close All Windows And Sync"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.1
Datum    = 31-01-2025
__________________________________________________________________
Description:

Met deze tool wordt je "starting view" geopend en worden de daarna inactieve views gesloten.
Hierna krijg je de mogelijkheid om je model te synchroniseren mits WorkShared.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [11.02.2025] - 1.1 Messageboxes toegepast ipv pyrevit forms.
- [31.01.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
import sys  # Voeg deze regel toe om de sys module te importeren
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import RevitCommandId, PostableCommand
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Functie om de "Starting View" te vinden en actief te maken
def activate_starting_view():
    starting_view_settings = StartingViewSettings.GetStartingViewSettings(doc)
    starting_view_id = starting_view_settings.ViewId

    if starting_view_id != ElementId.InvalidElementId:
        starting_view = doc.GetElement(starting_view_id)
        uidoc.ActiveView = starting_view
        return starting_view
    else:
        MessageBox.Show("Starting view niet ingesteld.", "Close All Windows & Sync | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()
        return None

# Haal de actieve view op
active_view = activate_starting_view()

if active_view:
    # Sluit alle andere vensters behalve de actieve view
    uiviews = uidoc.GetOpenUIViews()
    for uiview in uiviews:
        if uiview.ViewId != active_view.Id:
            try:
                uiview.Close()
            except Exception as e:
                MessageBox.Show("Fout bij het sluiten van view {0}: {1}".format(uiview.ViewId, e), "Close All Windows & Sync | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
else:
    MessageBox.Show("Kan niet verder gaan zonder actieve starting view.", "Close All Windows & Sync | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    sys.exit()

# Functie om te controleren of het een worksharing model is
def is_worksharing_model():
    return doc.IsWorkshared

# Functie om het "Synchronize with Central" dialoogvenster te openen
def open_synchronize_with_central():
    command_id = RevitCommandId.LookupCommandId("ID_FILE_SAVE_TO_CENTRAL")
    uidoc.Application.PostCommand(command_id)

# Controleer of het een worksharing model is en open het "Synchronize with Central" dialoogvenster
if is_worksharing_model():
    open_synchronize_with_central()
else:
    MessageBox.Show("Dit is geen workshared model.", "Close All Windows & Sync | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    sys.exit()