# -*- coding: utf-8 -*-

__title__ = "Copy Parameter to Parameter From/To (Single)"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.3
Datum    = 22.10.2025
__________________________________________________________________
Description:
SHIFT-CLICK to display options.

Met deze tool kan je parameters van een object uit je huidige model kopiëren naar één ander object.
Houd de Shift-knop ingedrukt bij het uitvoeren van deze actie en je kan de parameters aanpassen.
__________________________________________________________________
"""

import clr
import json
import os
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon
from pyrevit import forms

# Voeg deze regel toe
from Autodesk.Revit.DB import StorageType

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Configuratiebestand pad
config_path = os.path.join(os.path.dirname(__file__), 'config.json')

# Initialiseer de errors variabele
errors = {}

# Functie om parameterwaarde en opslagtype op te halen
def get_parameter_info(element, param_name):
    param = element.LookupParameter(param_name)
    if param:
        if param.StorageType == StorageType.String:
            return param.AsString(), param.StorageType
        elif param.StorageType == StorageType.ElementId:
            return param.AsElementId().IntegerValue, param.StorageType
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger(), param.StorageType
        elif param.StorageType == StorageType.Double:
            return param.AsDouble(), param.StorageType
    return None, None

# Functie om parameterwaarde in te stellen
def set_parameter_value(element, param_name, value, storage_type):
    param = element.LookupParameter(param_name)
    if param and value is not None:
        if storage_type == StorageType.String:
            param.Set(value)
        elif storage_type == StorageType.ElementId:
            param.Set(ElementId(value))
        elif storage_type == StorageType.Integer:
            param.Set(int(value))
        elif storage_type == StorageType.Double:
            param.Set(float(value))

# Functie om parameters te laden uit config.json
def load_parameters_from_config():
    if not os.path.exists(config_path) or os.stat(config_path).st_size == 0:
        MessageBox.Show("Config-bestand niet gevonden of leeg. De gebruiker heeft de actie gestopt.",
                        "Copy Parameters to Parameters From/To | Scholten BIM Consultancy",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()

    with open(config_path, 'r') as config_file:
        try:
            params = json.load(config_file)
            if not params:
                raise ValueError
            return params
        except Exception:
            MessageBox.Show("Config-bestand bevat ongeldige of lege JSON. De gebruiker heeft de actie gestopt.",
                            "Copy Parameters to Parameters From/To | Scholten BIM Consultancy",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            sys.exit()

# Aangepaste filter om Revit-links uit te sluiten
class ExcludeRevitLinks(ISelectionFilter):
    def AllowElement(self, element):
        return not isinstance(element, RevitLinkInstance)
    def AllowReference(self, reference, position):
        return True

try:
    selected_params = load_parameters_from_config()

    with forms.WarningBar(title="Pick reference element"):
        source_reference = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer het bronobject")
    source_element = doc.GetElement(source_reference.ElementId)

    # Waarden ophalen van het bronobject
    source_values = {p: get_parameter_info(source_element, p) for p in selected_params}

    with forms.WarningBar(title="Pick target element"):
        target_reference = uidoc.Selection.PickObject(ObjectType.Element, ExcludeRevitLinks(), "Selecteer het doelobject")
    target_element = doc.GetElement(target_reference.ElementId)

    t = Transaction(doc, "Copy Parameter to Parameter From/To (Single)")
    t.Start()

    category_name = target_element.Category.Name
    for param_name, (value, storage_type) in source_values.items():
        try:
            if target_element.LookupParameter(param_name):
                set_parameter_value(target_element, param_name, value, storage_type)
            else:
                errors.setdefault(param_name, {}).setdefault(category_name, 0)
                errors[param_name][category_name] += 1
        except Exception:
            errors.setdefault(param_name, {}).setdefault(category_name, 0)
            errors[param_name][category_name] += 1

    if errors:
        sorted_errors = sorted(errors.items())
        error_messages = [
            "Parameter '{}': ontbreekt in categorieën: {}.".format(
                param,
                ", ".join(["{} ({} object(en))".format(cat, count) for cat, count in cats.items()])
            ) for param, cats in sorted_errors
        ]
        MessageBox.Show("\n".join(error_messages),
                        "Copy Parameters to Parameters From/To | Scholten BIM Consultancy",
                        MessageBoxButtons.OK, MessageBoxIcon.Error)

    t.Commit()

except OperationCanceledException:
    MessageBox.Show("De gebruiker heeft de actie gestopt.",
                    "Copy Parameters to Parameters From/To | Scholten BIM Consultancy",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

except Exception as e:
    MessageBox.Show("Er is een fout opgetreden:\n{}".format(str(e)),
                    "Copy Parameters to Parameters From/To | Scholten BIM Consultancy",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)

finally:
    if 't' in locals() and t.HasStarted():
        t.RollBack()
