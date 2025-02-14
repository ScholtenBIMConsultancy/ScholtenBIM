# -*- coding: utf-8 -*-

__title__ = "Get Rotation Element"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.1
Datum    = 31.01.2025
__________________________________________________________________
Description:

Met deze tool kan je de rotatie van een element opvragen.
__________________________________________________________________
How-to:

-> Selecteer het object en run het script.
__________________________________________________________________
Last update:

- [11.02.2025] - 1.1 Als script tussentijds wordt gestopt volgt er geen foutmelding meer. Icons toegevoegd aan de meldingen.
- [30.01.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

try:
    # Vraag de gebruiker om een object te selecteren voor het uitlezen van de rotatie
    MessageBox.Show("Selecteer een object om de rotatie uit te lezen.", "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
    selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een object om de rotatie uit te lezen.")
    selection = doc.GetElement(selected_ref.ElementId)

    # Rotatie ophalen
    location = selection.Location
    if isinstance(location, LocationPoint):
        rotation = location.Rotation
        rotation_degrees = rotation * 180 / 3.14159  # Omzetten naar graden
        message = "De rotatie van het geselecteerde object is {:.2f} graden.".format(rotation_degrees)
    else:
        message = "Het geselecteerde object heeft geen rotatie."

    # Weergeven in een MessageBox
    MessageBox.Show(message, "Resultaat", "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)

except OperationCanceledException:
    MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Get Rotation Element | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)