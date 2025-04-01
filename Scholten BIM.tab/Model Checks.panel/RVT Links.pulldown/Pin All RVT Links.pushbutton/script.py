# -*- coding: utf-8 -*-

__title__ = "Pin All Revit Links"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.4
Datum    = 20.12.2025
__________________________________________________________________
Description:

Met deze tool kan je controleren of alle Revit Links gepind zijn en bij niet gepinde links kan je deze met 1 klik pinnen. Niet gepinde links worden weergegeven.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [01.04.2025] - 1.4 Als er geen Revit Links zijn wordt dan aangegeven i.p.v. dat de Links al zijn gepind.
- [25.03.2025] - 1.3 Revit 2025 compatible gemaakt
- [11.02.2025] - 1.2 Icons toegevoegd aan de meldingen en messageboxes gebruikt ipv taskdialogs
- [06.02.2025] - 1.1 Toevoegen van opsommingstekens en aangepaste tekst voor niet gepinde links
- [20.12.2024] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Verkrijg de huidige document en UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Functie om te controleren of een link gepind is
def is_link_pinned(link):
    return link.Pinned

# Functie om alle Revit links te pinnen
def pin_all_links(links):
    t = Transaction(doc, "Pin all Revit links")
    t.Start()
    for link in links:
        link.Pinned = True
    t.Commit()

# Verzamelen van alle Revit links in het document
collector = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
links = [link for link in collector]

if len(links) == 0:
    MessageBox.Show("Er zijn geen Revit links in het document.", "Pin All Revit Links | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
else:
    # Controleren of alle links gepind zijn
    unpinned_links = [link for link in links if not is_link_pinned(link)]

    if len(unpinned_links) > 0:
        # Log de niet gepinde links
        unpinned_link_types = ["• " + doc.GetElement(link.GetTypeId()).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for link in unpinned_links]
        unpinned_links_str = "\n".join(unpinned_link_types)

        # Popup venster met de vraag of de links gepind moeten worden
        if len(unpinned_links) == 1:
            message = "Er is 1 niet gepinde Revit link gevonden.\n\nNiet gepinde link:\n\n{}\n\nWilt u deze link pinnen?".format(unpinned_links_str)
        else:
            message = "Er zijn {} niet gepinde Revit links gevonden.\n\nNiet gepinde links:\n\n{}\n\nWilt u deze links pinnen?".format(len(unpinned_links), unpinned_links_str)
        
        result = MessageBox.Show(message, "Pin All Revit Links | Scholten BIM Consultancy", MessageBoxButtons.YesNo, MessageBoxIcon.Error)

        if result == DialogResult.Yes:
            pin_all_links(unpinned_links)
            MessageBox.Show("Alle Revit links zijn nu gepind.", "Pin All Revit Links | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
        else:
            MessageBox.Show("Er zijn {} niet gepinde links. Deze zijn niet gepind.".format(len(unpinned_links)), "Pin All Revit Links | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    else:
        MessageBox.Show("Alle Revit links zijn al gepind.", "Pin All Revit Links | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
