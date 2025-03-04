# -*- coding: utf-8 -*-

__title__ = "Copy Parameter to Parameter From/To"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 17.02.2025
__________________________________________________________________
Description:
SHIFT-CLICK to display options.

Met deze tool kan je parameters van een object uit je huidige model kopieren naar andere objecten uit je huidige model. Deze parameters zullen 1-op-1 naar dezelfde parameters worden weggeschreven.
Houdt de Shift knop ingedrukt bij het uitvoeren van deze actie en je kan de parameters aanpassen.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

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
from pyrevit import forms, script, revit

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
    param = element.LookupParameter(param_name)
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
        else:
            raise TypeError("Unsupported parameter storage type")

# Functie om parameters te selecteren met een PyRevit Form
def select_parameters(element):
    params = [p.Definition.Name for p in element.Parameters]
    params.sort()
    selected_params = forms.SelectFromList.show(params, multiselect=True, title="Copy Parameters to Parameters From/To| Scholten BIM Consultancy")
    if not selected_params:
        MessageBox.Show("Geen parameters geselecteerd. De gebruiker heeft de actie gestopt.", "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()
    return selected_params

# Functie om parameters op te slaan in config.json
def save_parameters_to_config(params):
    with open(config_path, 'w') as config_file:
        json.dump(params, config_file)

# Functie om parameters te laden uit config.json
def load_parameters_from_config():
    if not os.path.exists(config_path):
        MessageBox.Show("Config-bestand niet gevonden. De gebruiker heeft de actie gestopt.", 
                        "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()
    
    if os.stat(config_path).st_size == 0:
        MessageBox.Show("Config-bestand is leeg. De gebruiker heeft de actie gestopt.", 
                        "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()
    
    with open(config_path, 'r') as config_file:
        try:
            params = json.load(config_file)
            if not params:
                MessageBox.Show("Geen parameters gevonden in het config-bestand. De gebruiker heeft de actie gestopt.", 
                                "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", 
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
                sys.exit()
            return params
        except json.JSONDecodeError:
            MessageBox.Show("Config-bestand bevat ongeldige JSON. De gebruiker heeft de actie gestopt.", 
                            "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            sys.exit()

# Aangepaste filter om Revit-links uit te sluiten
class ExcludeRevitLinks(ISelectionFilter):
    def AllowElement(self, element):
        return not isinstance(element, RevitLinkInstance)
    
    def AllowReference(self, reference, position):
        return True

try:
    if (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
        with forms.WarningBar (title="Pick reference element"):
            selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een object om de parameters uit te lezen.")
        element = doc.GetElement(selected_ref.ElementId)
        selected_params = select_parameters(element)
        save_parameters_to_config(selected_params)
    else:
        selected_params = load_parameters_from_config()
        with forms.WarningBar (title="Pick reference element"):
            source_reference = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer het bronobject")
        source_element = doc.GetElement(source_reference.ElementId)
        
        source_values = {param_name: get_parameter_info(source_element, param_name) for param_name in selected_params}
        
        with forms.WarningBar (title="Pick target element"):
            target_references = uidoc.Selection.PickObjects(ObjectType.Element, ExcludeRevitLinks(), "Selecteer de doelobjecten")
        target_elements = [doc.GetElement(ref.ElementId) for ref in target_references]

        t = Transaction(doc, "Copy Parameter to Parameter From/To")
        t.Start()

        for element in target_elements:
            category_name = element.Category.Name
            for param_name, (value, storage_type) in source_values.items():
                try:
                    if element.LookupParameter(param_name):
                        set_parameter_value(element, param_name, value, storage_type)
                    else:
                        errors.setdefault(category_name, {}).setdefault(param_name, 0)
                        errors[category_name][param_name] += 1
                except Exception:
                    errors.setdefault(category_name, {}).setdefault(param_name, 0)
                    errors[category_name][param_name] += 1

        if errors:
            error_messages = ["Categorie '{}': Parameter '{}' ontbreekt in {} object(en).".format(category, param, count) for category, params in errors.items() for param, count in params.items()]
            MessageBox.Show("\n".join(error_messages), "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        else:
            if t is not None and t.HasStarted():
                t.Commit()
except OperationCanceledException:
    MessageBox.Show("De gebruiker heeft de actie gestopt.", "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
finally:
    if t is not None and t.HasStarted():
        t.RollBack()