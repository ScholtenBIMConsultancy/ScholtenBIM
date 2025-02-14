# -*- coding: utf-8 -*-

__title__ = "Change LineStyle on Selection"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 14.02.2025
__________________________________________________________________
Description:

Met deze tool worden de detail lines van de geselecteerde detail item(s) omgezet naar een LineStyle naar keuze.
__________________________________________________________________
How-to:

-> Selecteer één of meerdere Detail Items.
-> Run het script.
-> Selecteer een LineStyle.
-> Bevestig met Apply.
__________________________________________________________________
Last update:

- [14.02.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import Application, Form, Button, ComboBox, Label, MessageBox, MessageBoxButtons, MessageBoxIcon, MessageBoxOptions, MessageBoxDefaultButton, GroupBox
from System.Drawing import Point, Size
import System

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

class LineStyleForm(Form):
    def __init__(self, selected_ids):
        self.Text = "Change LineStyle | Scholten BIM Consultancy"
        self.Width = 350
        self.Height = 175   
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        self.selected_ids = selected_ids

        self.label = Label()
        self.label.Text = "Kies een Line Style:"
        self.label.Top = 20
        self.label.Left = 20
        self.label.Width = 200
        self.Controls.Add(self.label)

        self.comboBox = ComboBox()
        self.comboBox.Top = 50
        self.comboBox.Left = 20
        self.comboBox.Width = 300
        self.Controls.Add(self.comboBox)

        self.apply_button = Button()
        self.apply_button.Text = "Apply"
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

        self.load_line_styles()

    def load_line_styles(self):
        line_styles = FilteredElementCollector(doc).OfClass(GraphicsStyle).ToElements()
        unique_styles = {style.Name: style for style in line_styles}
        sorted_styles = sorted(unique_styles.values(), key=lambda x: x.Name)
        for style in sorted_styles:
            self.comboBox.Items.Add(style.Name)
        self.comboBox.SelectedItem = "<Invisible lines>"

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
            for element_id in self.selected_ids:
                element = doc.GetElement(element_id)
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
            MessageBox.Show("LineStyles zijn aangepast.", "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            self.Close()
        except Exception as e:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
            self.Close()

    def on_cancel_button_click(self, sender, event):
        self.Close()

# Huidige selectie ophalen
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    MessageBox.Show("Geen elementen geselecteerd.", "Change LineStyle | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
else:
    form = LineStyleForm(selected_ids)
    Application.Run(form)