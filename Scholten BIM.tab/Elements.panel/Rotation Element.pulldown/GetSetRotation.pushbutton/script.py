# -*- coding: utf-8 -*-

__title__ = "Get & Set Rotation Element"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 31.01.2025
__________________________________________________________________
Description:

Met deze tool kan je de rotatie van een element opvragen en toepassen op een ander object.
__________________________________________________________________
How-to:

-> Selecteer het object en run het script.
__________________________________________________________________
Last update:

- [10.03.2025] - 1.2 Het is nu mogelijk om de rotatie op meerdere elementen te gelijk toe te passen.
- [11.02.2025] - 1.1 Als script tussentijds wordt gestopt volgt er geen foutmelding meer. Icons toegevoegd aan de meldingen.
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
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
from pyrevit import forms

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

try:
    # Vraag de gebruiker om een object te selecteren voor het uitlezen van de rotatie
    with forms.WarningBar(title="Pick reference element"):
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick reference element.")
    element = doc.GetElement(selected_ref.ElementId)

    # Rotatie ophalen
    location = element.Location
    if isinstance(location, LocationPoint):
        rotation = location.Rotation
        rotation_degrees = rotation * 180 / 3.14159  # Omzetten naar graden
        if rotation_degrees == 360:
            rotation_degrees = 0
        MessageBox.Show("De rotatie van het geselecteerde object is {:.2f} graden.".format(rotation_degrees), "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    else:
        MessageBox.Show("Het geselecteerde object heeft geen rotatie.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()

    # Vraag de gebruiker om meerdere nieuwe objecten te selecteren voor het toepassen van de rotatie
    with forms.WarningBar(title="Pick target elements"):
        new_selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Pick target elements.")
    new_elements = [doc.GetElement(ref.ElementId) for ref in new_selected_refs]

    # Rotatie toepassen op elk geselecteerd doelobject
    t = Transaction(doc, "Get & Set Rotation Element")
    t.Start()
    try:
        for new_element in new_elements:
            new_location = new_element.Location
            if isinstance(new_location, LocationPoint):
                # Reset de rotatie van het doel-element naar 0
                new_location.Rotate(Line.CreateBound(new_location.Point, new_location.Point + XYZ(0, 0, 1)), -new_location.Rotation)
                # Pas de nieuwe rotatie toe
                new_location.Rotate(Line.CreateBound(new_location.Point, new_location.Point + XYZ(0, 0, 1)), rotation)
        t.Commit()
        MessageBox.Show("De rotatie is toegepast op de geselecteerde objecten.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    except Exception as e:
        MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        if t.HasStarted():
            t.RollBack()
except OperationCanceledException:
    MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)