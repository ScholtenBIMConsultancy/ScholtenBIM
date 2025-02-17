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
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, Keys, Control
from pyrevit import forms

# Voeg deze regel toe
from Autodesk.Revit.DB import StorageType

# Huidig document en UI ophalen
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument

# Configuratiebestand pad
config_path = os.path.join(os.path.dirname(__file__), 'config.json')

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
    return selected_params

# Functie om parameters op te slaan in config.json
def save_parameters_to_config(params):
    with open(config_path, 'w') as config_file:
        json.dump(params, config_file)

# Functie om parameters te laden uit config.json
def load_parameters_from_config():
    if os.path.exists(config_path):
        with open(config_path, 'r') as config_file:
            return json.load(config_file)
    return []

try:
    # Controleer of Shift is ingedrukt
    if (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
        # Vraag de gebruiker om een object te selecteren voor het uitlezen van parameters
        MessageBox.Show("Selecteer een object om de parameters uit te lezen.", "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een object om de parameters uit te lezen.")
        element = doc.GetElement(selected_ref.ElementId)
        
        # Parameters selecteren en opslaan in config.json
        selected_params = select_parameters(element)
        save_parameters_to_config(selected_params)
    else:
        # Parameters laden uit config.json
        selected_params = load_parameters_from_config()

        # Selecteer het bronobject
        source_reference = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer het bronobject")
        source_element = doc.GetElement(source_reference.ElementId)

        # Parameterwaarden en opslagtypes ophalen van het bronobject
        source_values = {}
        for param_name in selected_params:
            value, storage_type = get_parameter_info(source_element, param_name)
            source_values[param_name] = (value, storage_type)

        # Selecteer de doelobjecten
        target_references = uidoc.Selection.PickObjects(ObjectType.Element, "Selecteer de doelobjecten")
        target_elements = [doc.GetElement(ref.ElementId) for ref in target_references]

        # Transactie starten
        t = Transaction(doc, "Kopieer parameterwaarden")
        t.Start()

        errors = {}

        # Parameterwaarden instellen voor de doelobjecten
        for element in target_elements:
            category_name = element.Category.Name
            for param_name, (value, storage_type) in source_values.items():
                try:
                    if element.LookupParameter(param_name):
                        set_parameter_value(element, param_name, value, storage_type)
                    else:
                        if category_name not in errors:
                            errors[category_name] = {}
                        if param_name not in errors[category_name]:
                            errors[category_name][param_name] = 0
                        errors[category_name][param_name] += 1
                except Exception as e:
                    if category_name not in errors:
                        errors[category_name] = {}
                    if param_name not in errors[category_name]:
                        errors[category_name][param_name] = 0
                    errors[category_name][param_name] += 1

        if errors:
            error_messages = []
            for category, params in errors.items():
                for param, count in params.items():
                    error_messages.append("Categorie '{}': Parameter '{}' ontbreekt in {} object(en).".format(category, param, count))
            MessageBox.Show("\n".join(error_messages), "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        else:
            print("Parameterwaarden succesvol gekopieerd.")

        # Transactie voltooien
        t.Commit()
except OperationCanceledException:
    MessageBox.Show("De gebruiker heeft de actie gestopt.", "Copy Parameters to Parameters From/To| Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    if (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
        sys.exit()