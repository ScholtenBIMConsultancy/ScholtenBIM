# -*- coding: utf-8 -*-

__title__ = "CropView on new view"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 26.06.2025
__________________________________________________________________
Description:

Met deze tool kan je op basis van de geselecteerde lijnen een crop view instellen op een nieuw aangemaakte view.
__________________________________________________________________
How-to:

-> Zorg dat je een rechthoek hebt gemodelleerd 
   met detaillines of modellines.
-> Selecteer deze lijnen.
-> Run het script.
__________________________________________________________________
Last update:

- [26.06.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Nog uit te testen met minder/meer lijnen geselecteerd.
__________________________________________________________________
"""

import clr
import sys
from pyrevit import revit, DB

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

# Active document en view
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Validatie actieve view
if not isinstance(active_view, ViewPlan) or active_view.ViewType not in (ViewType.FloorPlan, ViewType.CeilingPlan):
    MessageBox.Show(
        "De actieve view is geen Floor Plan of Ceiling Plan.",
        "CropView on copy | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Error
    )
    sys.exit()

if not hasattr(active_view, "CropBoxActive"):
    MessageBox.Show(
        "De actieve view ondersteunt geen crop regions.",
        "CropView on copy | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Error
    )
    sys.exit()

# Verzamelen en filteren van geselecteerde lijnen
selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
lines = [el for el in selection if isinstance(el, DB.CurveElement)]
if not lines:
    MessageBox.Show(
        "Selecteer eerst lijnen om een crop view te maken.",
        "CropView on copy | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Error
    )
    sys.exit()

# Berekening bounding box
min_pt = max_pt = None
for line in lines:
    g = line.GeometryCurve
    for i in (0, 1):
        p = g.GetEndPoint(i)
        if min_pt is None:
            min_pt = XYZ(p.X, p.Y, p.Z)
            max_pt = XYZ(p.X, p.Y, p.Z)
        else:
            min_pt = XYZ(min(min_pt.X, p.X), min(min_pt.Y, p.Y), min(min_pt.Z, p.Z))
            max_pt = XYZ(max(max_pt.X, p.X), max(max_pt.Y, p.Y), max(max_pt.Z, p.Z))

if not min_pt or not max_pt:
    MessageBox.Show(
        "Kan geen geldige bounding box berekenen.",
        "CropView on copy | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Error
    )
    sys.exit()

bounding_box = BoundingBoxXYZ()
bounding_box.Min = min_pt
bounding_box.Max = max_pt

# Één transactie voor álle model-wijzigingen
t = Transaction(doc, "CropView: duplicate & crop in één transactie")
try:
    t.Start()

    # 1. Dupliceer view mét detailing
    new_view_id = active_view.Duplicate(ViewDuplicateOption.WithDetailing)
    new_view = doc.GetElement(new_view_id)
    if not new_view or not isinstance(new_view, View):
        raise Exception("Duplicatie mislukt: nieuwe view niet gevonden.")

    # 2. Unieke naam "<origineel> - Copy", "<origineel> - Copy 1", etc.
    base_name = "{} - Copy".format(active_view.Name)
    new_name = base_name
    existing = {v.Name for v in FilteredElementCollector(doc).OfClass(View)}
    idx = 1
    while new_name in existing:
        new_name = "{} {}".format(base_name, idx)
        idx += 1
    new_view.Name = new_name

    # 3. Scope box uitzetten
    new_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)

    # 4. Crop-box instellen
    new_view.CropBox = bounding_box
    new_view.CropBoxActive = True
    new_view.CropBoxVisible = True

    # 5. Verwijder gebruikte lijnen één voor één
    for l in lines:
        doc.Delete(l.Id)

    # 6. Commit
    t.Commit()

    # 7. Zet nieuwe view actief
    uidoc.ActiveView = new_view

    MessageBox.Show(
        "Crop view '{}' is succesvol gemaakt en lijnen verwijderd!".format(new_name),
        "CropView on copy | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Information
    )

except Exception as e:
    # Rollback bij fout
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    MessageBox.Show(str(e), "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
    sys.exit()
