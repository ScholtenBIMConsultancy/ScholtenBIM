# -*- coding: utf-8 -*-

__title__ = "Get & Set Rotation Element"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.1
Datum    = 31.01.2025
__________________________________________________________________
Description:

Met deze tool kan je de rotatie van een element opvragen en toepassen op een ander object.
__________________________________________________________________
How-to:

-> Selecteer het object en run het script.
__________________________________________________________________
Last update:

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

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

try:
    # Vraag de gebruiker om een object te selecteren voor het uitlezen van de rotatie
    MessageBox.Show("Selecteer één object om de rotatie uit te lezen.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
    selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer één object om de rotatie uit te lezen.")
    element = doc.GetElement(selected_ref.ElementId)

    # Rotatie ophalen
    location = element.Location
    if isinstance(location, LocationPoint):
        rotation = location.Rotation
        rotation_degrees = rotation * 180 / 3.14159  # Omzetten naar graden
        MessageBox.Show("De rotatie van het geselecteerde object is {:.2f} graden.".format(rotation_degrees), "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    else:
        MessageBox.Show("Het geselecteerde object heeft geen rotatie.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()

    # Vraag de gebruiker om een nieuw object te selecteren voor het toepassen van de rotatie
    MessageBox.Show("Selecteer één object om de rotatie toe te passen.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
    new_selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer één object object om de rotatie toe te passen.")
    new_element = doc.GetElement(new_selected_ref.ElementId)

    # Rotatie toepassen
    new_location = new_element.Location
    if isinstance(new_location, LocationPoint):
        try:
            t = Transaction(doc, "Get & Set Rotation Element")
            t.Start()
            new_location.Rotate(Line.CreateBound(new_location.Point, new_location.Point + XYZ(0, 0, 1)), rotation)
            t.Commit()
            MessageBox.Show("De rotatie is toegepast op het nieuwe object.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as e:
            MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            if t.HasStarted():
                t.RollBack()
    else:
        MessageBox.Show("Het nieuwe geselecteerde object kan geen rotatie hebben.", "Get & Set Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)

except OperationCanceledException:
    MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)