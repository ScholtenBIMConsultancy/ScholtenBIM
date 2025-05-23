# -*- coding: utf-8 -*-

__title__ = "Change dimension offset"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 12.05.2025
__________________________________________________________________
Description:
SHIFT CLICK

Met deze tool worden de tekst posities van de geselecteerde dimensions aangepast met een voorgedefineerde waarde.
__________________________________________________________________
How-to:

-> D.m.v. Shift ingedrukt te houden en het script te runnen
     is het mogelijk om een default waarde in te stellen.
-> Selecteer één of meerdere dimensions.
-> Run het script.
__________________________________________________________________
Last update:

- [12.05.2025] - 1.2 Hij past nu juist de dimensions in een section aan.
- [25.02.2025] - 1.1 Het is nu ook mogelijk om de dimensions in een section aan te passen.
- [14.02.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Text aanpassing voor schuine dimensions.
__________________________________________________________________
"""

import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import json
import os
import sys
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, MessageBox, MessageBoxIcon, MessageBoxButtons

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

view_direction = doc.ActiveView.ViewDirection

# Functie om de schaal van de actieve view uit te lezen
def get_active_view_scale(doc):
    active_view = doc.ActiveView
    return active_view.Scale

class DistanceForm(Form):
    def __init__(self):
        self.Text = "Change dimension offset | Scholten BIM Consultancy"
        self.Width = 350
        self.Height = 175
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        self.label = Label()
        self.label.Text = "Voer de afstandswaarde in (mm):"
        self.label.Top = 20
        self.label.Left = 20
        self.label.Width = 200
        self.Controls.Add(self.label)

        self.textBox = TextBox()
        self.textBox.Top = 50
        self.textBox.Left = 20
        self.textBox.Width = 100
        self.textBox.Text = "4"  # Default waarde instellen op 4 mm
        self.Controls.Add(self.textBox)

        self.apply_button = Button()
        self.apply_button.Text = "OK"
        self.apply_button.Top = 90
        self.apply_button.Left = 20
        self.apply_button.Click += self.on_apply_button_click
        self.Controls.Add(self.apply_button)

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Top = 90
        self.cancel_button.Left = 100
        self.cancel_button.Click += self.on_cancel_button_click
        self.Controls.Add(self.cancel_button)

        # Event afvangen bij sluiten van het formulier
        self.FormClosing += self.on_form_closing

    def on_apply_button_click(self, sender, event):
        try:
            self.distance_value = float(self.textBox.Text)
        except ValueError:
            MessageBox.Show("Ongeldige invoer. De standaardwaarde van 4 mm wordt gebruikt.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            self.distance_value = 4
        save_distance_value(self.distance_value)  # Sla altijd de waarde op
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel_button_click(self, sender, event):
        MessageBox.Show("Actie geannuleerd door gebruiker.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def on_form_closing(self, sender, event):
        save_distance_value(4)  # Sla de standaardwaarde op als het formulier wordt gesloten

# Functie om de afstandswaarde op te slaan in config.json in dezelfde directory als het script
def save_distance_value(value):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")
    config = {"distance_value": value}
    with open(config_path, "w") as f:
        json.dump(config, f)

# Functie om de afstandswaarde te laden uit config.json in dezelfde directory als het script
def load_distance_value():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r") as f:
            config = json.load(f)
            return config.get("distance_value", None)
    else:
        return None

# Functie om te controleren of een dimensie horizontaal is
def is_horizontal(dimension, view_type, view_direction):
    line = dimension.Curve
    if view_type == ViewType.Section:
        if abs(view_direction.X) != 0:
            return abs(line.Direction.Y) != 0
        elif abs(view_direction.Y) != 0:
            return abs(line.Direction.X) != 0
    else:
        return abs(line.Direction.X) > abs(line.Direction.Y) and abs(line.Direction.Z) < 0.01
# Functie om te controleren of een dimensie verticaal is
def is_vertical(dimension, view_type, view_direction):
    line = dimension.Curve
    if view_type == ViewType.Section:
        if abs(view_direction.Y) != 0:
            return abs(line.Direction.Z) != 0
        elif abs(view_direction.X) != 0:
            return abs(line.Direction.Z) != 0
    else:
        return abs(line.Direction.Y) > abs(line.Direction.X) and abs(line.Direction.X) < 0.01

# Check of Shift is ingedrukt tijdens het uitvoeren van het script
shift_pressed = System.Windows.Forms.Control.ModifierKeys == System.Windows.Forms.Keys.Shift

if shift_pressed:
    form = DistanceForm()
    if form.ShowDialog() == DialogResult.OK:
        distance_value = form.distance_value
        save_distance_value(distance_value)
    else:
        MessageBox.Show("Ongeldige invoer. De standaardwaarde van 4 mm wordt gebruikt.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        distance_value = 4  # Default waarde van 4 mm
else:
    distance_value = load_distance_value()
    if distance_value is None:
        MessageBox.Show("Geen geldige afstandswaarde gevonden in config.json. Script wordt beëindigd.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()

    selected_element_ids = uidoc.Selection.GetElementIds()

    if not selected_element_ids:
        MessageBox.Show("Geen elementen geselecteerd. Selecteer een of meer dimensies en probeer opnieuw.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()

    selected_dimensions = []
    for el_id in selected_element_ids:
        element = doc.GetElement(el_id)
        if isinstance(element, Dimension):
            selected_dimensions.append(element)

    if not selected_dimensions:
        MessageBox.Show("Geen dimensies geselecteerd. Selecteer één of meer dimensies en probeer opnieuw.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()

    # Functie om het type view te controleren en een melding weer te geven
    def check_view_type():
        active_view = doc.ActiveView
        view_type = active_view.ViewType
        return view_type

    # Functie om een melding te geven en het script te stoppen als de dimensie niet horizontaal of verticaal is
    def check_dimension_orientation(dimension, view_type, view_direction):
        if not (is_horizontal(dimension, view_type, view_direction) or is_vertical(dimension, view_type, view_direction)):
            MessageBox.Show("De geselecteerde dimensie is niet horizontaal of verticaal. Het script wordt beëindigd.", "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            sys.exit()

    # Controleer de oriëntatie van elke geselecteerde dimensie
    for dimension in selected_dimensions:
        view_type = check_view_type()
        check_dimension_orientation(dimension, view_type, view_direction)

    def adjust_text_position(dimension, view_type):
        scale = get_active_view_scale(doc)  # Schaal hier ophalen
        distance_in_feet = (distance_value * scale) / 304.8  # Omzetten van mm naar feet en vermenigvuldigen met schaal
        if view_type == ViewType.Section:
            if dimension.NumberOfSegments >= 2:
                adjust_multiple_segments_section(dimension, distance_in_feet)
            else:
                adjust_single_segment_section(dimension, distance_in_feet)
        else:
            if dimension.NumberOfSegments >= 2:
                adjust_multiple_segments(dimension, distance_in_feet)
            else:
                adjust_single_segment(dimension, distance_in_feet)

    def adjust_single_segment(dimension, distance_in_feet):
        dimension.ResetTextPosition()
        text_position = dimension.TextPosition
        if is_horizontal(dimension, ViewType.FloorPlan, view_direction):
            new_position = XYZ(text_position.X, text_position.Y - distance_in_feet, text_position.Z)
        else:
            new_position = XYZ(text_position.X + distance_in_feet, text_position.Y, text_position.Z)
        dimension.TextPosition = new_position
        doc.Regenerate()  # Zorg ervoor dat de nieuwe positie wordt toegepast

    def adjust_multiple_segments(dimension, distance_in_feet):
        for segment in dimension.Segments:
            segment.ResetTextPosition()
            text_position = segment.TextPosition
            if is_horizontal(dimension, ViewType.FloorPlan, view_direction):
                new_position = XYZ(text_position.X, text_position.Y - distance_in_feet, text_position.Z)
            else:
                new_position = XYZ(text_position.X + distance_in_feet, text_position.Y, text_position.Z)
            segment.TextPosition = new_position
        doc.Regenerate()  # Zorg ervoor dat de nieuwe positie wordt toegepast

    def adjust_single_segment_section(dimension, distance_in_feet):
        dimension.ResetTextPosition()
        text_position = dimension.TextPosition

        if is_horizontal(dimension, ViewType.Section, view_direction):
            new_position = XYZ(text_position.X, text_position.Y, text_position.Z - distance_in_feet)
        else:
            # Controleer de richting van de view
            if view_direction.Y < 0:
                new_position = XYZ(text_position.X + distance_in_feet, text_position.Y, text_position.Z)
            elif view_direction.Y > 0:
                new_position = XYZ(text_position.X - distance_in_feet, text_position.Y, text_position.Z)
            elif view_direction.X < 0:
                new_position = XYZ(text_position.X, text_position.Y - distance_in_feet, text_position.Z)
            elif view_direction.X > 0:
                new_position = XYZ(text_position.X, text_position.Y + distance_in_feet, text_position.Z)
            else:
                # Fallback: geen richting herkend
                new_position = text_position

        dimension.TextPosition = new_position
        doc.Regenerate()  # Zorg ervoor dat de nieuwe positie wordt toegepast
        
    def adjust_multiple_segments_section(dimension, distance_in_feet):
        for segment in dimension.Segments:
            segment.ResetTextPosition()
            text_position = segment.TextPosition

            if is_horizontal(dimension, ViewType.Section, view_direction):
                new_position = XYZ(text_position.X, text_position.Y, text_position.Z - distance_in_feet)
            else:
                # Controleer de richting van de view
                if view_direction.Y < 0:
                    new_position = XYZ(text_position.X + distance_in_feet, text_position.Y, text_position.Z)
                elif view_direction.Y > 0:
                    new_position = XYZ(text_position.X - distance_in_feet, text_position.Y, text_position.Z)
                elif view_direction.X < 0:
                    new_position = XYZ(text_position.X, text_position.Y - distance_in_feet, text_position.Z)
                elif view_direction.X > 0:
                    new_position = XYZ(text_position.X, text_position.Y + distance_in_feet, text_position.Z)
                else:
                    # Fallback indien geen duidelijke richting
                    new_position = text_position

            segment.TextPosition = new_position

        doc.Regenerate()  # Zorg ervoor dat de nieuwe positie wordt toegepast


    # Gebruik de Revit transactiebeheerder
    t = Transaction(doc, "Change dimension offset")
    t.Start()
    try:
        view_type = check_view_type()  # Controleer het viewtype en toon de MessageBox één keer
        for dimension in selected_dimensions:
            adjust_text_position(dimension, view_type)
        t.Commit()
    except Exception as e:
        t.RollBack()
        MessageBox.Show("Fout: {}".format(e), "Change dimension offset | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error) 