__title__ = "Copy Parameter to Parameter"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version = 1.0
Datum    = 20.12.2024
__________________________________________________________________
Omschrijving:

Met deze tool kan je parameter van een object uit een link kopieren naar een object uit je huidige model.
Houdt de Shift knop ingedrukt bij het uitvoeren van deze actie en je kan de uit te lezen en weg te schrijven parameter aanpassen.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [20.12.2024] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
import os
import json
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Drawing import Icon
from System.Windows.Forms import Application, Form, Label, TextBox, Button, MessageBox, Keys, Control, FormBorderStyle, DialogResult, FormStartPosition

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

config_file_path = os.path.join(os.path.dirname(__file__), 'config.json')

class ParameterForm(Form):
    def __init__(self, title, read_default, write_default):
        self.Text = title
        self.Width = 400
        self.Height = 250
        self.FormBorderStyle = FormBorderStyle.FixedDialog  # Maak de form niet aanpasbaar
        self.StartPosition = FormStartPosition.CenterScreen  # Plaats de form in het midden van het scherm

        self.label1 = Label()
        self.label1.Text = 'Parameter uitlezen:'
        self.label1.Top = 20
        self.label1.Left = 20
        self.label1.Width = 150  # Maak de label breder
        self.Controls.Add(self.label1)

        self.textbox1 = TextBox()
        self.textbox1.Top = 45
        self.textbox1.Left = 20
        self.textbox1.Width = 350  # Maak de textbox breder
        self.textbox1.Text = read_default
        self.Controls.Add(self.textbox1)

        self.label2 = Label()
        self.label2.Text = 'Parameter wegschrijven:'
        self.label2.Top = 90
        self.label2.Left = 20
        self.label2.Width = 150  # Maak de label breder
        self.Controls.Add(self.label2)

        self.textbox2 = TextBox()
        self.textbox2.Top = 115
        self.textbox2.Left = 20
        self.textbox2.Width = 350  # Maak de textbox breder
        self.textbox2.Text = write_default
        self.Controls.Add(self.textbox2)

        self.button = Button()
        self.button.Text = 'OK'
        self.button.Top = 160
        self.button.Left = 20
        self.button.Click += self.on_button_click
        self.Controls.Add(self.button)

    def on_button_click(self, sender, event):
        self.read_param = self.textbox1.Text
        self.write_param = self.textbox2.Text
        self.DialogResult = DialogResult.OK

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
    return {"read_param": "Comments", "write_param": "Mark"}

def save_config(config):
    with open(config_file_path, 'w') as f:
        json.dump(config, f)

config = load_config()

try:
    if (Control.ModifierKeys & Keys.Shift) == Keys.Shift:
        # Toon formulier om parameters in te vullen
        form = ParameterForm("Parameters instellen", config["read_param"], config["write_param"])
        
        if form.ShowDialog() == DialogResult.OK:
            config["read_param"] = form.read_param
            config["write_param"] = form.write_param
            save_config(config)
            MessageBox.Show("Parameters opgeslagen. Voer het script opnieuw uit zonder Shift ingedrukt te houden.")
            sys.exit()  # Stop het script na het opslaan van de parameters
    else:
        # Gebruik parameters uit configuratiebestand
        read_param_name = config["read_param"]
        write_param_name = config["write_param"]

        # Stap 1: Selecteer een element om uit te lezen
        MessageBox.Show("Selecteer een element om uit te lezen.")
        selected_element = uidoc.Selection.PickObject(ObjectType.LinkedElement, "Selecteer een element om uit te lezen.")
        element = get_linked_element(doc, selected_element)

        # Lees de waarde van de opgegeven parameter
        read_param = element.LookupParameter(read_param_name)
        if read_param:
            read_value = read_param.AsString()
        else:
            MessageBox.Show("Parameter {} niet gevonden.".format(read_param_name))
            read_value = None

        if read_value:
            # Stap 2: Selecteer een element om de gegevens naar te schrijven
            MessageBox.Show("Selecteer een element om de gegevens naar te schrijven.")
            selected_element = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een element om de gegevens naar te schrijven.")
            target_element = doc.GetElement(selected_element.ElementId)

            # Schrijf de waarde naar de opgegeven parameter
            write_param = target_element.LookupParameter(write_param_name)
            if write_param:
                with Transaction(doc, "Schrijf parameterwaarde") as t:
                    t.Start()
                    write_param.Set(read_value)
                    t.Commit()
                MessageBox.Show("Waarde '{}' geschreven naar parameter '{}'.".format(read_value, write_param_name))
            else:
                MessageBox.Show("Parameter {} niet gevonden.".format(write_param_name))
except Exception as e:
    MessageBox.Show("Er is een fout opgetreden: {}".format(e))