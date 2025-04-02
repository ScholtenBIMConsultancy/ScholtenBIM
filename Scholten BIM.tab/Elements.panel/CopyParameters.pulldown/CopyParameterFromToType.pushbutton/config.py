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
from RevitServices.Persistence import DocumentManager
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

# Functie om parameters te selecteren met een PyRevit Form
def select_parameters(element_type):
    params = [p.Definition.Name for p in element_type.Parameters]
    params.sort()
    selected_params = forms.SelectFromList.show(params, multiselect=True, title="Copy Parameters to Parameters From/To | Scholten BIM Consultancy")
    if not selected_params:
        MessageBox.Show("Geen parameters geselecteerd. De gebruiker heeft de actie gestopt.", "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
        return None
    return selected_params

# Functie om parameters op te slaan in config.json
def save_parameters_to_config(params):
    with open(config_path, 'w') as config_file:
        json.dump(params, config_file)

try:
    with forms.WarningBar(title="Pick reference element"):
        selected_ref = uidoc.Selection.PickObject(ObjectType.Element, "Selecteer een object om de parameters uit te lezen.")
    element = doc.GetElement(selected_ref.ElementId)
    element_type = doc.GetElement(element.GetTypeId())
    selected_params = select_parameters(element_type)
    if selected_params:
        save_parameters_to_config(selected_params)
except Exception as e:
    if "The user aborted the pick operation" in str(e):
        # Geen actie nodig, gewoon afsluiten
        pass
    else:
        MessageBox.Show(str(e), "Copy Parameters to Parameters From/To | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
