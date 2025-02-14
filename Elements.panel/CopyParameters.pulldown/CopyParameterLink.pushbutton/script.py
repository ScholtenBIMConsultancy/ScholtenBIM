# -*- coding: utf-8 -*-

__title__ = "Copy Parameters to Parameters From Link"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 20.12.2024
__________________________________________________________________
Description:
SHIFT-CLICK to display options.

Met deze tool kan je parameters van een object uit een link kopieren naar een object uit je huidige model.
Houdt de Shift knop ingedrukt bij het uitvoeren van deze actie en je kan de uit te lezen en weg te schrijven parameters aanpassen.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [11.02.2025] - 1.2 Icons toegevoegd aan de meldingen
- [29.01.2025] - 1.1 Melding toegevoegd voor lege parameterwaarde en script stopt. Waarde altijd omgezet naar string.
- [20.12.2024] - 1.0 RELEASE
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
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Drawing import Icon
from System.Windows.Forms import Application, Form, Label, TextBox, Button, MessageBox, Keys, Control, FormBorderStyle, DialogResult, FormStartPosition, MessageBox, MessageBoxButtons, DialogResult, MessageBoxIcon
from pyrevit import forms

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

config_file_path = os.path.join(os.path.dirname(__file__), 'config.json')

def get_linked_element(doc, ref):
    link_instance = doc.GetElement(ref)
    if isinstance(link_instance, RevitLinkInstance):
        link_doc = link_instance.GetLinkDocument()
        link_element_id = ref.LinkedElementId
        return link_doc.GetElement(link_element_id)
    return None

def load_config():
    if os.path.exists(config_file_path):
        with open(config_file_path, 'r') as f:
            return json.load(f)
    return {"read_param": "Comments", "write_param": ["Mark"]}

def save_config(config):
    with open(config_file_path, 'w') as f:
        json.dump(config, f)

config = load_config()

try:
    if (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
        # Vraag de gebruiker om een object te selecteren voor het uitlezen van parameters
        MessageBox.Show("Selecteer een object uit een link om de parameters uit te lezen.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
        selected_ref = uidoc.Selection.PickObject(ObjectType.LinkedElement, "Selecteer een object uit een link om de parameters uit te lezen.")
        element = get_linked_element(doc, selected_ref)
        if element is None:
            MessageBox.Show("Het geselecteerde element is geen gelinkt element.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sys.exit()

        # Lees de parameters van het geselecteerde element uit
        parameters = element.Parameters
        param_names = [param.Definition.Name for param in parameters]
        param_names.sort()  # Sorteer de parameters alfabetisch

        # Voeg "Handmatig invoeren" bovenaan de lijst toe
        param_names.insert(0, "Handmatig invoeren...")

        # Maak een pulldown menu met de parameters
        read_param_name = forms.SelectFromList.show(param_names, title="Selecteer een parameter om uit te lezen", button_name="Selecteer")
        if read_param_name == "Handmatig invoeren...":
            read_param_name = forms.ask_for_string(prompt="Voer de naam van de parameter in om uit te lezen:", title="Handmatige invoer")

        # Vraag de gebruiker om een object te selecteren voor het wegschrijven van parameters
        MessageBox.Show("Selecteer een object om de parameters naar te schrijven.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een object om de parameters naar te schrijven.")
        target_element = doc.GetElement(selected_ref.ElementId)

        # Lees de parameters van het geselecteerde element uit
        parameters = target_element.Parameters
        param_names = [param.Definition.Name for param in parameters if not param.IsReadOnly]
        param_names.sort()  # Sorteer de parameters alfabetisch

        # Voeg "Handmatig invoeren" bovenaan de lijst toe
        param_names.insert(0, "Handmatig invoeren...")

        # Maak een pulldown menu met de parameters (meervoudige selectie)
        write_param_names = forms.SelectFromList.show(param_names, title="Selecteer parameters om naar te schrijven", button_name="Selecteer", multiselect=True)
        if "Handmatig invoeren..." in write_param_names:
            write_param_names.remove("Handmatig invoeren...")
            manual_param_name = forms.ask_for_string(prompt="Voer de naam van de parameter in om naar te schrijven:", title="Handmatige invoer")
            write_param_names.append(manual_param_name)

        if read_param_name and write_param_names:
            config["read_param"] = read_param_name
            config["write_param"] = write_param_names
            save_config(config)
            MessageBox.Show("Parameters opgeslagen. Voer het script opnieuw uit zonder Shift ingedrukt te houden.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            sys.exit()  # Stop het script na het opslaan van de parameters
    else:
        # Gebruik parameters uit configuratiebestand
        read_param_name = config.get("read_param", "")
        write_param_names = config.get("write_param", [])

        if not read_param_name or not write_param_names:
            MessageBox.Show("Geen parameters gevonden in het configuratiebestand. Voer het script opnieuw uit met Shift ingedrukt om parameters in te stellen.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sys.exit()

        # Stap 1: Selecteer een element om uit te lezen
        MessageBox.Show("Selecteer een element uit een link om uit te lezen.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
        try:
            selected_element = uidoc.Selection.PickObject(ObjectType.LinkedElement, "Selecteer een element uit een link om uit te lezen.")
        except OperationCanceledException:
            MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            sys.exit()  # Stop het script zonder foutmelding als de selectie wordt geannuleerd
        element = get_linked_element(doc, selected_element)
        if element is None:
            MessageBox.Show("Het geselecteerde element is geen gelinkt element.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sys.exit()

        # Lees de waarde van de opgegeven parameter
        read_param = element.LookupParameter(read_param_name)
        if read_param:
            read_value = read_param.AsValueString() or read_param.AsString()
            if not read_value:
                MessageBox.Show("De waarde van de parameter '{}' is leeg. Het script wordt gestopt.".format(read_param_name), "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
                sys.exit()
        else:
            MessageBox.Show("Parameter '{}' niet gevonden. Het script wordt gestopt.".format(read_param_name), "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sys.exit()

        # Stap 2: Selecteer meerdere elementen om de parameters naar te schrijven
        MessageBox.Show("Selecteer één of meerdere elementen om de parameters naar te schrijven.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Question)
        try:
            selected_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Selecteer meerdere elementen om de parameters naar te schrijven.")
        except OperationCanceledException:
            MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            sys.exit()  # Stop het script zonder foutmelding als de selectie wordt geannuleerd
        target_elements = [doc.GetElement(ref.ElementId) for ref in selected_refs]

        # Schrijf de waarden naar de opgegeven parameters
        invalid_elements = {param: [] for param in write_param_names}
        with Transaction(doc, "Copy Parameters to Parameters From Link") as t:
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
            MessageBox.Show("De volgende elementen bevatten de parameter(s) niet:\n\n{}".format(missing_elements_info), "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        else:
            MessageBox.Show("Waarden geschreven naar de geselecteerde parameters.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
except OperationCanceledException:
    MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Copy Parameters to Parameters From Link | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)