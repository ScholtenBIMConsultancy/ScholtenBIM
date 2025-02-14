__title__ = "CropView on dependent"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 20.12.2024
__________________________________________________________________
Description:

Met deze tool kan je op basis van de geselecteerde lijnen een crop view instellen op een nieuw aangemaakte dependent view.
__________________________________________________________________
How-to:

-> Zorg dat je een rechthoek hebt gemodelleerd 
   met detaillines of modellines.
-> Selecteer deze lijnen.
-> Run het script.
__________________________________________________________________
Last update:

- [13.02.2025] - 1.2 Foutaanpassingen verwerkt
- [11.02.2025] - 1.1 PyRevit Forms omgezet naar Windows Forms.
- [20.12.2024] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Nog uit te testen met minder/meer lijnen geselecteerd.
__________________________________________________________________
"""

# Importeer de benodigde Revit API-modules en PyRevit tools
import clr
import sys
from pyrevit import revit, DB, HOST_APP

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
#from Autodesk.Revit.DB import XYZ, BoundingBoxXYZ, Transaction, ViewDuplicateOption, FilteredElementCollector

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Controleer of de actieve view crop regions ondersteunt
if not hasattr(active_view, "CropBoxActive"):
    MessageBox.Show("De actieve view ondersteunt geen crop regions.", "CropView on dependent | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    sys.exit()  # Stop het script als de actieve view geen crop regions ondersteunt

# Verkrijg de geselecteerde elementen
selection = [doc.GetElement(elid) for elid in uidoc.Selection.GetElementIds()]

# Filter geselecteerde lijnen
lines = [el for el in selection if isinstance(el, DB.CurveElement)]
if not lines:
    MessageBox.Show("Selecteer eerst lijnen om een crop view te maken.", "CropView on dependent | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    sys.exit()  # Stop het script als er geen lijnen zijn geselecteerd

# Bereken de bounding box van de lijnen
min_point = None
max_point = None

for line in lines:
    geom_line = line.GeometryCurve
    if not min_point:
        min_point = geom_line.GetEndPoint(0)
        max_point = geom_line.GetEndPoint(0)
    
    # Update bounding box
    for i in range(2):
        pt = geom_line.GetEndPoint(i)
        min_point = XYZ(
            min(min_point.X, pt.X),
            min(min_point.Y, pt.Y),
            min(min_point.Z, pt.Z)
        )
        max_point = XYZ(
            max(max_point.X, pt.X),
            max(max_point.Y, pt.Y),
            max(max_point.Z, pt.Z)
        )
  
# Controleer of een geldige bounding box is gevonden
if not min_point or not max_point:
    MessageBox.Show("Kan geen geldige bounding box berekenen.", "CropView on dependent | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    sys.exit()  # Stop het script als er geen geldige bounding box is gevonden

# Bounding box instellen als crop view
bounding_box = BoundingBoxXYZ()
bounding_box.Min = min_point
bounding_box.Max = max_point

t = Transaction(doc, "CropView on dependent")
try:
    t.Start()

    # Maak een dependent view aan
    dependent_view_id = active_view.Duplicate(ViewDuplicateOption.AsDependent)  # Maak een dependent view
    dependent_view = doc.GetElement(dependent_view_id)  # Verkrijg de nieuwe dependent view

    # Genereer een unieke naam voor de nieuwe view
    base_name = active_view.Name + " - Dependent"
    new_name = base_name
    index = 1
    view_names = [view.Name for view in FilteredElementCollector(doc).OfClass(DB.View)]

    while new_name in view_names:
        new_name = "{} {}".format(base_name, index)
        index += 1

    dependent_view.Name = new_name  # Geef de nieuwe view een unieke naam

    # Schakel de ScopeBox uit
    dependent_view.get_Parameter(DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(ElementId.InvalidElementId)

    # Pas de crop box van de dependent view aan
    dependent_view.CropBox = bounding_box
    dependent_view.CropBoxActive = True
    dependent_view.CropBoxVisible = True
    doc.Delete(uidoc.Selection.GetElementIds())
    
    t.Commit()

    # Stel de nieuwe view in als actieve view buiten de transactie
    uidoc.ActiveView = dependent_view
    
    MessageBox.Show("Crop View is succesvol ingesteld en de lijnen zijn verwijderd!", "CropView | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)

except Exception as e:
    MessageBox.Show(str(e), "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error)
    t.RollBack()