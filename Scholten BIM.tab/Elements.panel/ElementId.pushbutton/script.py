# -*- coding: utf-8 -*-

__title__ = "Get Linked Id's"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.2
Datum    = 10.11.2025
__________________________________________________________________
Description:

Met deze tool kan je element id's uitlezen van gelinkte elementen.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [10.11.2025] - 1.2 GUI verbeterd met kopieer functies.
- [07.04.2025] - 1.1 Toevoeging uitlezen welk linked file het betreft.
- [31.03.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""


import clr
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import RevitLinkInstance, BuiltInParameter
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import script
from pyrevit.forms import WPFWindow
from System.Windows import Clipboard, Thickness, TextWrapping, HorizontalAlignment, GridLength
from System.Windows.Controls.Primitives import Popup, PlacementMode
from System.Windows.Threading import DispatcherTimer
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon
import traceback

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class LinkedIdsWindow(WPFWindow):
    def __init__(self, xaml_path, element_infos):
        WPFWindow.__init__(self, xaml_path)

        for i, (info_text, element_id, parts) in enumerate(element_infos):
            self.add_element_info(parts, element_id, i)

        self.btn_close.Click += self.close_action
        self.btn_copy_by_category.Click += self.copy_by_category_action
        self.element_infos = element_infos

    def add_element_info(self, parts, element_id, index):
        from System.Windows.Controls import Grid, ColumnDefinition, TextBlock, Button, StackPanel, Orientation
        from System.Windows.Media import Brushes

        panel = Grid(Margin=Thickness(0, 3, 0, 3))
        panel.ColumnDefinitions.Add(ColumnDefinition())
        panel.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(220)))

        base_color = Brushes.WhiteSmoke if index % 2 == 1 else Brushes.Transparent
        panel.Background = base_color

        def on_hover(sender, args):
            sender.Background = Brushes.LightGray
        def on_leave(sender, args):
            sender.Background = base_color
        panel.MouseEnter += on_hover
        panel.MouseLeave += on_leave

        tb = TextBlock(TextWrapping=TextWrapping.Wrap, Width=500)
        tb.Inlines.Add(self._make_inline("Category: ", bold=True))
        tb.Inlines.Add(self._make_inline(parts["Category"] + "\n"))
        tb.Inlines.Add(self._make_inline("Family: ", bold=True))
        tb.Inlines.Add(self._make_inline(parts["Family"] + "\n"))
        tb.Inlines.Add(self._make_inline("Linked File: ", bold=True))
        tb.Inlines.Add(self._make_inline(parts["Linked File"] + "\n"))
        tb.Inlines.Add(self._make_inline("ID: ", bold=True))
        tb.Inlines.Add(self._make_inline(str(parts["ID"])))
        Grid.SetColumn(tb, 0)

        btn_copy_line = Button(Content="Kopieer regel", Width=100, Height=30)
        def copy_line(sender, args):
            Clipboard.SetText(
                "Category: {0}\nFamily: {1}\nLinked File: {2}\nID: {3}".format(
                    parts["Category"], parts["Family"], parts["Linked File"], parts["ID"]
                )
            )
            self.show_temporary_popup("Regel gekopieerd!", sender)

        btn_copy_id = Button(Content="Kopieer ID", Width=100, Height=30)
        def copy_id(sender, args):
            Clipboard.SetText(str(parts["ID"]))
            self.show_temporary_popup("ID gekopieerd!", sender)

        btn_copy_line.Click += copy_line
        btn_copy_id.Click += copy_id

        btn_panel = StackPanel(Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Right)
        btn_panel.Children.Add(btn_copy_line)
        btn_panel.Children.Add(btn_copy_id)
        Grid.SetColumn(btn_panel, 1)

        panel.Children.Add(tb)
        panel.Children.Add(btn_panel)
        self.elementList.Children.Add(panel)

    def show_temporary_popup(self, message, target_button, duration_ms=1500):
        from System.Windows.Controls import TextBlock
        from System.Windows.Media import Brushes
        import System

        popup = Popup()
        popup.PlacementTarget = target_button
        popup.Placement = PlacementMode.Bottom
        popup.AllowsTransparency = True
        popup.StaysOpen = False

        tb = TextBlock()
        tb.Text = message
        tb.Background = Brushes.LightYellow
        tb.Foreground = Brushes.Black
        tb.Padding = Thickness(10)
        tb.Margin = Thickness(5)
        popup.Child = tb
        popup.IsOpen = True

        timer = DispatcherTimer()
        timer.Interval = System.TimeSpan.FromMilliseconds(duration_ms)
        def close_popup(sender, e):
            popup.IsOpen = False
            timer.Stop()
        timer.Tick += close_popup
        timer.Start()

    def copy_by_category_action(self, sender, args):
        categories = sorted(set(info[2]["Category"] for info in self.element_infos))
        xaml_popup = script.get_bundle_file('SelectCategory.xaml')
        popup = SelectCategoryWindow(xaml_popup, categories)
        popup.ShowDialog()

        if popup.selected_category:
            ids = [str(info[2]["ID"]) for info in self.element_infos if info[2]["Category"] == popup.selected_category]
            if ids:
                Clipboard.SetText(";".join(ids))
                self.show_temporary_popup("Gekopieerd: {0} ID's voor '{1}'".format(len(ids), popup.selected_category), sender)

    def _make_inline(self, text, bold=False):
        from System.Windows.Documents import Run
        from System.Windows import FontWeights
        run = Run(text)
        if bold:
            run.FontWeight = FontWeights.Bold
        return run

    def close_action(self, sender, args):
        self.Close()


class SelectCategoryWindow(WPFWindow):
    def __init__(self, xaml_path, categories):
        WPFWindow.__init__(self, xaml_path)
        for cat in categories:
            self.combo_categories.Items.Add(cat)

        if self.combo_categories.Items.Count > 0:
            self.combo_categories.SelectedIndex = 0

        self.btn_ok.Click += self.ok_action
        self.btn_cancel.Click += self.cancel_action
        self.selected_category = None

    def ok_action(self, sender, args):
        self.selected_category = self.combo_categories.SelectedItem
        self.Close()

    def cancel_action(self, sender, args):
        self.selected_category = None
        self.Close()


try:
    try:
        selected_refs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Selecteer elementen in gelinkte bestanden")
    except:
        selected_refs = []

    element_infos = []
    for ref in selected_refs:
        link_instance = doc.GetElement(ref)
        if isinstance(link_instance, RevitLinkInstance):
            link_doc = link_instance.GetLinkDocument()
            linked_element = link_doc.GetElement(ref.LinkedElementId)
            link_type = link_instance.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
            category = linked_element.Category.Name
            family_type_name = linked_element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
            parts = {
                "Category": category,
                "Family": family_type_name,
                "Linked File": link_type,
                "ID": linked_element.Id
            }
            element_infos.append((None, linked_element.Id, parts))

    xaml_path = script.get_bundle_file('LinkedIds.xaml')

    if not xaml_path:
        MessageBox.Show(
            "LinkedIds.xaml niet gevonden in de bundle.\nZorg dat LinkedIds.xaml in dezelfde map staat als dit script.",
            "Linked IDs | Scholten BIM Consultancy",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
    elif not element_infos:
        MessageBox.Show(
            "Geen gekoppelde elementen gevonden.",
            "Linked IDs | Scholten BIM Consultancy",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    else:
        LinkedIdsWindow(xaml_path, element_infos).ShowDialog()

except Exception as ex:
    MessageBox.Show(
        "Er ging iets mis:\n\n{0}\n\nTraceback:\n{1}".format(ex, traceback.format_exc()),
        "Linked IDs | Scholten BIM Consultancy",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
    )