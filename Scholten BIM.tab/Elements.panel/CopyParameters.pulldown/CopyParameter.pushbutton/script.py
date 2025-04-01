# -*- coding: utf-8 -*-

__title__ = "Copy Parameter to Parameter"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 29.01.2025
__________________________________________________________________
Description:
SHIFT-CLICK to display options.

Met deze tool kan je een parameter van een object uit je huidige model kopieren naar andere parameters van een ander object uit je huidige model.
Houdt de Shift knop ingedrukt bij het uitvoeren van deze actie en je kan de uit te lezen en weg te schrijven parameter aanpassen.
__________________________________________________________________
How-to:
 
-> Run het script.
__________________________________________________________________
Last update:

- [01.04.2025] - 1.2 Juiste manier van config.py gebruikt, geen veranderde werking van het script.
- [11.02.2025] - 1.1 Icons toegevoegd aan de meldingen en de mogelijkheid om uit een pulldown menu de uit te lezen en weg te schrijven parameters te selecteren.
- [29.01.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
import os
import json
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

config_file_path = os.path.join(os.path.dirname(__file__), 'config.json')

def load_config():
    if os.path.exists(config_file_path):
        with open(config_file_path, 'r') as f:
            return json.load(f)
    return {"read_param": "Comments", "write_param": ["Mark"]}

config = load_config()

try:
    # Gebruik parameters uit configuratiebestand
    read_param_name = config["read_param"]
    write_param_names = config["write_param"]

    # Stap 1: Selecteer een element om uit te lezen
    with forms.WarningBar(title="Pick element"):
        selected_element = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een element uit het huidige model om uit te lezen.")
    element = doc.GetElement(selected_element.ElementId)

    # Lees de waarde van de opgegeven parameter
    read_param = element.LookupParameter(read_param_name)
    if read_param:
        read_value = read_param.AsValueString() or read_param.AsString()
        if not read_value:
            MessageBox.Show("De waarde van de parameter '{}' is leeg. Het script wordt gestopt.".format(read_param_name), "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sys.exit()
    else:
        MessageBox.Show("Parameter '{}' niet gevonden. Het script wordt gestopt.".format(read_param_name), "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        sys.exit()

    # Stap 2: Selecteer meerdere elementen om de parameters naar te schrijven
    with forms.WarningBar(title="Pick element(s)"):
        try:
            selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecteer één of meerdere elementen om de parameters naar te schrijven.")
        except OperationCanceledException:
            MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            sys.exit()  # Stop het script zonder foutmelding als de selectie wordt geannuleerd
    target_elements = [doc.GetElement(ref.ElementId) for ref in selected_refs]

    # Schrijf de waarden naar de opgegeven parameters
    invalid_elements = {param: [] for param in write_param_names}
    with Transaction(doc, "Copy Parameters to Parameters") as t:
        t.Start()
        for target_element in target_elements:
            for write_param_name in write_param_names:
                write_param = target_element.LookupParameter(write_param_name)
                if write_param:
                    write_param.Set(read_value)
                else:
                    category_name = target_element.Category.Name if target_element.Category else "N/A"
                    invalid_elements[write_param_name].append((category_name, target_element.Id))
        t.Commit()

    # Toon een MessageBox met de elementen die de parameter niet bevatten, gegroepeerd per parameter
    missing_elements_info = ""
    for param, elements in invalid_elements.items():
        if elements:
            elements_info = "\n".join(["• {} (ID: {})".format(category_name, id) for category_name, id in elements])
            missing_elements_info += "Parameter '{}':\n{}\n\n".format(param, elements_info)

    if missing_elements_info:
        MessageBox.Show("De volgende elementen bevatten de parameter(s) niet:\n\n{}".format(missing_elements_info), "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    else:
        MessageBox.Show("Waarden geschreven naar de geselecteerde parameters.", "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
except OperationCanceledException:
    MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)