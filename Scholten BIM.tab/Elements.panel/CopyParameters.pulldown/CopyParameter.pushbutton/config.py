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
    # Vraag de gebruiker om een object te selecteren voor het uitlezen van parameters
    with forms.WarningBar(title="Pick element"):
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een object om de parameters uit te lezen.")
    element = doc.GetElement(selected_ref.ElementId)

    # Lees de parameters van het geselecteerde element uit
    parameters = element.Parameters
    param_names = [param.Definition.Name for param in parameters]
    param_names.sort()  # Sorteer de parameters alfabetisch

    # Voeg "Handmatig invoeren" bovenaan de lijst toe
    param_names.insert(0, "Handmatig invoeren...")

    # Maak een pulldown menu met de parameters (enkelvoudige selectie)
    read_param_name = forms.SelectFromList.show(param_names, title="Selecteer een parameter om uit te lezen", button_name="Selecteer", multiselect=False)
    if read_param_name == "Handmatig invoeren...":
        read_param_name = forms.ask_for_string(prompt="Voer de naam van de parameter in om uit te lezen:", title="Handmatige invoer")

    # Vraag de gebruiker om een ander object te selecteren voor het wegschrijven van parameters
    with forms.WarningBar(title="Pick element(s)"):
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een ander object om de parameters naar te schrijven.")
    target_element = doc.GetElement(selected_ref.ElementId)

    # Lees de parameters van het geselecteerde element uit
    parameters = target_element.Parameters
    write_param_names = [param.Definition.Name for param in parameters if not param.IsReadOnly]
    write_param_names.sort()  # Sorteer de parameters alfabetisch

    # Voeg "Handmatig invoeren" bovenaan de lijst toe
    write_param_names.insert(0, "Handmatig invoeren...")

    # Maak een pulldown menu met de parameters (meervoudige selectie)
    write_param_names = forms.SelectFromList.show(write_param_names, title="Selecteer parameters om naar te schrijven", button_name="Selecteer", multiselect=True)
    if "Handmatig invoeren..." in write_param_names:
        write_param_names.remove("Handmatig invoeren...")
        manual_param_name = forms.ask_for_string(prompt="Voer de naam van de parameter in om naar te schrijven:", title="Handmatige invoer")
        write_param_names.append(manual_param_name)

    if read_param_name and write_param_names:
        config["read_param"] = read_param_name
        config["write_param"] = write_param_names
        save_config(config)
        MessageBox.Show("Parameters opgeslagen. Voer het script opnieuw uit zonder Shift ingedrukt te houden.", "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
        sys.exit()  # Stop het script na het opslaan van de parameters
except OperationCanceledException:
    MessageBox.Show("De bewerking is onderbroken door de gebruiker.", "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    sys.exit()  # Stop het script zonder foutmelding als de selectie wordt geannuleerd
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Copy Parameter to Parameter | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
