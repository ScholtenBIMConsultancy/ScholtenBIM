# -*- coding: utf-8 -*-

__title__ = "Copy Parameter to Parameter From/To (Type)"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 17.02.2025
__________________________________________________________________
Description:
SHIFT-CLICK to display options.

Met deze tool kan je parameters van een object uit je huidige model kopieren naar andere objecten uit je huidige model. Deze parameters zullen 1-op-1 naar dezelfde parameters worden weggeschreven.
Houdt de Shift knop ingedrukt bij het uitvoeren van deze actie en je kan de parameters aanpassen.
__________________________________________________________________
How-to:

-> SHIFT-Click om parameters te selecteren.
-> Single-Click om parameters uitlezen/wegschrijven tussen de elementen.
__________________________________________________________________
Last update:

- [01.04.2025] - 1.2 Juiste manier van config.py gebruikt, geen veranderde werking van het script.
- [04.03.2025] - 1.1 Update transaction & empty config file 
- [17.02.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
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
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, Keys, Control
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

# Initialiseer de transactievariabele
t = None

# Functie om parameterwaarde en opslagtype op te halen
def get_parameter_info(element, param_name):
    param = element.Symbol.LookupParameter(param_name) if element.Symbol else None
    if param:
        value = None
        if param.StorageType == StorageType.String:
            value = param.AsString()
        elif param.StorageType == StorageType.ElementId:
            value = param.AsElementId().IntegerValue
        elif param.StorageType == StorageType.Integer:
            value = param.AsInteger()
        elif param.StorageType == StorageType.Double:
            value = param.AsDouble()
        return value, param.StorageType
    return None, None

# Functie om parameterwaarde in te stellen
def set_parameter_value(element, param_name, value, storage_type):
    param = element.Symbol.LookupParameter(param_name) if element.Symbol else None
    if param and value is not None:
        if storage_type == StorageType.String:
            param.Set(value)
        elif storage_type == StorageType.ElementId:
            param.Set(ElementId(value))
        elif storage_type == StorageType.Integer:
            param.Set(int(value))
        elif storage_type == StorageType.Double:
            param.Set(float(value))
        else:
            raise TypeError("Unsupported parameter storage type")

# Functie om parameters te laden uit config.json
def load_parameters_from_config():
    if not os.path.exists(config_path):
        MessageBox.Show("Config-bestand niet gevonden. De gebruiker heeft de actie gestopt.", 
                        "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()
    
    if os.stat(config_path).st_size == 0:
        MessageBox.Show("Config-bestand is leeg. De gebruiker heeft de actie gestopt.", 
                        "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()
    
    with open(config_path, 'r') as config_file:
        try:
            params = json.load(config_file)
            if not params:
                MessageBox.Show("Geen parameters gevonden in het config-bestand. De gebruiker heeft de actie gestopt.", 
                                "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", 
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                sys.exit()
            return params
        except json.JSONDecodeError:
            MessageBox.Show("Config-bestand bevat ongeldige JSON. De gebruiker heeft de actie gestopt.", 
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
    
    source_values = {param_name: get_parameter_info(source_element, param_name) for param_name in selected_params}
    
    with forms.WarningBar(title="Pick target element"):
        target_references = uidoc.Selection.PickObjects(ObjectType.Element, ExcludeRevitLinks(), "Selecteer de doelobjecten")
    target_elements = [doc.GetElement(ref.ElementId) for ref in target_references]

    t = Transaction(doc, "Copy Parameter to Parameter From/To")
    t.Start()

    for element in target_elements:
        category_name = element.Category.Name
        for param_name, (value, storage_type) in source_values.items():
            try:
                if element.Symbol.LookupParameter(param_name):
                    set_parameter_value(element, param_name, value, storage_type)
                else:
                    errors.setdefault(param_name, {}).setdefault(category_name, 0)
                    errors[param_name][category_name] += 1
            except Exception:
                errors.setdefault(param_name, {}).setdefault(category_name, 0)
                errors[param_name][category_name] += 1
    if errors:
        sorted_errors = sorted(errors.items())
        error_messages = ["Parameter '{}': ontbreekt in categorieën: {}.".format(param, ", ".join(["{} ({} object(en))".format(category, count) for category, count in categories.items()])) for param, categories in sorted_errors]
        MessageBox.Show("\n".join(error_messages), "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    
    if t is not None and t.HasStarted():
        t.Commit()
except OperationCanceledException:
    MessageBox.Show("De gebruiker heeft de actie gestopt.", "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
finally:
    if t is not None and t.HasStarted():
        t.RollBack()

