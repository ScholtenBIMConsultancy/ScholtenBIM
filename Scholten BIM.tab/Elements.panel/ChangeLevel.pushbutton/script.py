# -*- coding: utf-8 -*-

__title__ = "Change Level"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 14.10.2025
__________________________________________________________________
Description:

Met deze tool kan je op basis van geselecteerde elementen de level aanpassen zonder dat het element van positie veranderd.
__________________________________________________________________
How-to:

-> Selecteer 1 of meerdere elementen.
-> Run het script.
__________________________________________________________________
Last update:

- [14.10.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

from Autodesk.Revit.DB import BuiltInParameter, Level, FamilyInstance, MEPCurve
from Autodesk.Revit.DB.Mechanical import Space
from Autodesk.Revit.DB import FilteredElementCollector, UV
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

# Haal geselecteerde elementen
selected_ids = list(uidoc.Selection.GetElementIds())

if not selected_ids:
    forms.alert(
    "Geen elementen geselecteerd. Selecteer eerst elementen en start het script opnieuw.",
    title="Change Level | Scholten BIM Consultancy"
)
    raise SystemExit

# ----------------------------------------------------
# Functies
# ----------------------------------------------------
def apply_level(level):
    """Pas het geselecteerde level toe op de geselecteerde elementen"""
    changed_count = 0
    with revit.Transaction("Change Level", doc):
        for el_id in selected_ids:
            el = doc.GetElement(el_id)
            if isinstance(el, FamilyInstance):
                # Bereken offset
                current_level = doc.GetElement(el.LevelId)
                elevation_param = el.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                element_elev = elevation_param.AsDouble() if elevation_param else 0
                current_elev = current_level.Elevation if current_level else 0
                new_level_elev = level.Elevation
                new_offset = current_elev + element_elev - new_level_elev

                # Zet level
                level_param = el.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                if level_param:
                    level_param.Set(level.Id)

                # Zet offset
                offset_param = el.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
                if offset_param:
                    offset_param.Set(new_offset)

                changed_count += 1

            elif isinstance(el, MEPCurve):
                el.ReferenceLevel = level
                changed_count += 1

            elif isinstance(el, Space):
                point = el.Location.Point
                newspace = doc.Create.NewSpace(level, UV(point.X, point.Y))
                changed_count += 1

#    forms.alert("{} elementen aangepast naar level '{}'".format(changed_count, level.Name))

# ----------------------------------------------------
# GUI klasse
# ----------------------------------------------------
from pyrevit.forms import WPFWindow

class LevelChangerWindow(WPFWindow):
    def __init__(self, xaml_file_name):
        WPFWindow.__init__(self, xaml_file_name)
        # Vul combobox met levels
        self.levels = list(FilteredElementCollector(doc).OfClass(Level))
        if not self.levels:
            forms.alert("Geen Levels gevonden in het project.")
            self.Close()
            return
        self.combobox_levels.ItemsSource = self.levels
        self.combobox_levels.SelectedIndex = 0

        # Toon aantal geselecteerde elementen in label
        count = len(selected_ids)
        self.label_selection.Content = "{} {}".format(count, "element" if count == 1 else "elementen") + " geselecteerd"

    def button_apply_click(self, sender, e):
        """Pas het geselecteerde level toe en sluit venster"""
        level = self.combobox_levels.SelectedItem
        if not level:
            forms.alert("Selecteer een level uit het pulldown menu.")
            return
        apply_level(level)
        self.Close()

# ----------------------------------------------------
# Start GUI
# ----------------------------------------------------
LevelChangerWindow("ReferenceLevelSelection.xaml").ShowDialog()