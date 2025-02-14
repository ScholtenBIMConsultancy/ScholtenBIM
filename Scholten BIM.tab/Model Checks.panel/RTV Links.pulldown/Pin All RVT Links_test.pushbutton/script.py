__title__ = "Pin All Revit Links"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version = 1.0
Datum    = 20.12.2024
__________________________________________________________________
Omschrijving:

Met deze tool kan je op controleren of alle Revit Links gepind zijn en bij niet gepinde links kan je deze met 1 klik pinnen.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

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
clr.AddReference('IronPython.Wpf')

from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.UI import *
#from RevitServices.UI import UIApplication
from pyrevit import forms



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

# Controleren of alle links gepind zijn
unpinned_links = [link for link in links if not is_link_pinned(link)]

if len(unpinned_links) == 1:
    # Popup venster met de vraag of de links gepind moeten worden
    dialog = TaskDialog("Revit Links Pinnen")
    dialog.MainInstruction = "Er is 1 niet gepinde Revit link gevonden."
    dialog.MainContent = "Wilt u alle Revit links pinnen?"
    dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    dialog.DefaultButton = TaskDialogResult.Yes

    result = dialog.Show()

    if result == TaskDialogResult.Yes:
        pin_all_links(unpinned_links)
        TaskDialog.Show("Revit Links Pinnen", "Alle Revit links zijn nu gepind.")
    else:
        TaskDialog.Show("Revit Links Pinnen", "Er is 1 niet gepinde link. Deze is niet gepind.")

else:
    # Popup venster met de vraag of de links gepind moeten worden
    dialog = TaskDialog("Revit Links Pinnen")
    dialog.MainInstruction = "Er zijn {} niet gepinde Revit links gevonden.".format(len(unpinned_links))
    dialog.MainContent = "Wilt u alle Revit links pinnen?"
    dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    dialog.DefaultButton = TaskDialogResult.Yes

    result = dialog.Show()

    if result == TaskDialogResult.Yes:
        pin_all_links(unpinned_links)
        TaskDialog.Show("Revit Links Pinnen", "Alle Revit links zijn nu gepind.")
    else:
        TaskDialog.Show("Revit Links Pinnen", "Er zijn {} niet gepinde links. Deze zijn niet gepind.".format(len(unpinned_links)))