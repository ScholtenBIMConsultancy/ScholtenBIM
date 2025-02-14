# -*- coding: utf-8 -*-

__title__ = "Change LineStyle and Select"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 14-02-2025
__________________________________________________________________
Description:

Met deze tool worden de detail lines van nog te selecteren detail item(s) omgezet naar een LineStyle naar keuze.
__________________________________________________________________
How-to:

-> Run het script.
-> Selecteer één of meerdere objecten.
-> Kies een LineStyle.
-> Bevestig met Apply.
__________________________________________________________________
Last update:

- [14.02.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Zorgen dat Revit op voorgrond blijft.
__________________________________________________________________
"""

import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import Application, Form, Button, ComboBox, Label, MessageBox, MessageBoxButtons, MessageBoxIcon, MessageBoxOptions, MessageBoxDefaultButton, GroupBox, Keys
from System.Drawing import Point, Size

# Importeren van de ISelectionFilter interface
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

# Huidig document ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class DetailItemSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element.Category.Id.IntegerValue == int(BuiltInCategory.OST_DetailComponents)
    
    def AllowReference(self, refer, point):
        return False

class CustomForm(Form):
    def __init__(self):
        super(CustomForm, self).__init__()
        self.KeyPreview = True
        self.KeyDown += self.on_key_down

    def on_key_down(self, sender, event):
        if event.KeyCode == Keys.Escape:
            event.SuppressKeyPress = True

class LineStyleForm(CustomForm):
    def __init__(self):
        super(LineStyleForm, self).__init__()
        self.Text = "Change LineStyle | Scholten BIM Consultancy"
        self.Width = 350
        self.Height = 260   
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        self.selected_ids = []

        self.select_label = Label()
        self.select_label.Text = "Selecteer Element(en):"
        self.select_label.Top = 20
        self.select_label.Left = 20
        self.select_label.Width = 200
        self.Controls.Add(self.select_label)

        self.select_button = Button()
        self.select_button.Text = "Select"
        self.select_button.Top = 50
        self.select_button.Left = 20
        self.select_button.Click += self.on_select_button_click
        self.Controls.Add(self.select_button)

        self.separator = GroupBox()
        self.separator.Top = 90
        self.separator.Left = 20
        self.separator.Width = 340
        self.separator.Height = 2
        self.separator.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        self.separator.TabStop = False
        self.Controls.Add(self.separator)

        self.label = Label()
        self.label.Text = "Kies een Line Style:"
        self.label.Top = 110
        self.label.Left = 20
        self.label.Width = 200
        self.Controls.Add(self.label)

        self.comboBox = ComboBox()
        self.comboBox.Top = 140
        self.comboBox.Left = 20
        self.comboBox.Width = 300
        self.Controls.Add(self.comboBox)

        self.apply_button = Button()
        self.apply_button.Text = "Apply"
        self.apply_button.Top = 180
        self.apply_button.Left = 20
        self.apply_button.Click += self.on_apply_button_click
        self.Controls.Add(self.apply_button)

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Top = 180
        self.cancel_button.Left = 100
        self.cancel_button.Click += self.on_cancel_button_click
        self.Controls.Add(self.cancel_button)

        self.load_line_styles()

    def load_line_styles(self):
        line_styles = FilteredElementCollector(doc).OfClass(GraphicsStyle).ToElements()
        unique_styles = {style.Name: style for style in line_styles}
        sorted_styles = sorted(unique_styles.values(), key=lambda x: x.Name)
        for style in sorted_styles:
            self.comboBox.Items.Add(style.Name)
        self.comboBox.SelectedItem = "<Invisible lines>"

    def on_select_button_click(self, sender, event):
        try:
            self.selected_ids = uidoc.Selection.PickObjects(ObjectType.Element, DetailItemSelectionFilter(), "Selecteer Detail Items")
            self.BringToFront()
        except Exception as e:
            if "The user aborted the pick operation" in str(e):
                MessageBox.Show("De selectieactie is geannuleerd door de gebruiker.", "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            else:
                MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            self.BringToFront()
            return

    def on_apply_button_click(self, sender, event):
        if not self.selected_ids:
            MessageBox.Show("Geen elementen geselecteerd.", "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        selected_style_name = self.comboBox.SelectedItem
        invisible_line_style = None
        line_styles = FilteredElementCollector(doc).OfClass(GraphicsStyle).ToElements()
        for style in line_styles:
            if style.Name == selected_style_name:
                invisible_line_style = style
                break

        if invisible_line_style is None:
            MessageBox.Show("LineStyle '{}' niet gevonden.".format(selected_style_name), "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        t = Transaction(doc, "Update Line Style")
        try:
            t.Start()
            detail_item_found = False
            for ref in self.selected_ids:
                element = doc.GetElement(ref.ElementId)
                if isinstance(element, CurveElement):
                    element.LineStyle = invisible_line_style
                    ElementTransformUtils.MoveElement(doc, element.Id, XYZ(0, 0, 0.001))
                    ElementTransformUtils.MoveElement(doc, element.Id, XYZ(0, 0, -0.001))
                    detail_item_found = True
                else:
                    dependent_elements = element.GetDependentElements(None)
                    for dep_elem_id in dependent_elements:
                        dep_element = doc.GetElement(dep_elem_id)
                        if isinstance(dep_element, CurveElement):
                            dep_element.LineStyle = invisible_line_style
                            ElementTransformUtils.MoveElement(doc, dep_elem_id, XYZ(0, 0, 0.001))
                            ElementTransformUtils.MoveElement(doc, dep_elem_id, XYZ(0, 0, -0.001))
                            detail_item_found = True
            if not detail_item_found:
                MessageBox.Show("Geen detail items gevonden.", "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                t.RollBack()
                return
            t.Commit()
            MessageBox.Show("LineStyles zijn aangepast.", "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            self.Close()
        except Exception as e:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)

    def on_cancel_button_click(self, sender, event):
        self.Close()

# Formulier weergeven
form = LineStyleForm()
Application.Run(form)