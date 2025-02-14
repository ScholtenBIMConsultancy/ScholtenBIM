# -*- coding: utf-8 -*-

__title__ = "Change to <Invisible Lines>"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 14-02-2025
__________________________________________________________________
Description:

Met deze tool worden de detail lines van de geselecteerde detail item(s) omgezet naar <Invisible Lines>.
__________________________________________________________________
How-to:

-> Selecteer één of meerdere Detail Items.
-> Run het script.
__________________________________________________________________
Last update:

- [14.02.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import sys
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

# Huidig document ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Functie om de LineStyle van geselecteerde detail lines te updaten
def update_line_style():
    # LineStyle ophalen met de naam "<Invisible lines>"
    invisible_line_style = None
    line_styles = FilteredElementCollector(doc).OfClass(GraphicsStyle).ToElements()
    for style in line_styles:
        if style.Name == "<Invisible lines>":
            invisible_line_style = style
            break
    
    if invisible_line_style is None:
        MessageBox.Show("LineStyle '<Invisible lines>' niet gevonden.", "Change to <Invisible Lines> | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        sys.exit()
    
    selected_ids = uidoc.Selection.GetElementIds()
    
    if not selected_ids:
        MessageBox.Show("Geen elementen geselecteerd.", "Change to <Invisible Lines> | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        sys.exit()
    
    # Transactie starten
    t = Transaction(doc, "Update Line Style")
    try:
        t.Start()
        detail_item_found = False
        for elem_id in selected_ids:
            element = doc.GetElement(elem_id)
            if isinstance(element, CurveElement):
                element.LineStyle = invisible_line_style
                # Element tijdelijk verplaatsen en terugzetten
                ElementTransformUtils.MoveElement(doc, elem_id, XYZ(0, 0, 0.001))
                ElementTransformUtils.MoveElement(doc, elem_id, XYZ(0, 0, -0.001))
                detail_item_found = True
            else:
                dependent_elements = element.GetDependentElements(None)
                for dep_elem_id in dependent_elements:
                    dep_element = doc.GetElement(dep_elem_id)
                    if isinstance(dep_element, CurveElement):
                        dep_element.LineStyle = invisible_line_style
                        # Element tijdelijk verplaatsen en terugzetten
                        ElementTransformUtils.MoveElement(doc, dep_elem_id, XYZ(0, 0, 0.001))
                        ElementTransformUtils.MoveElement(doc, dep_elem_id, XYZ(0, 0, -0.001))
                        detail_item_found = True
        if not detail_item_found:
            MessageBox.Show("Geen Detail Item geselecteerd.", "Change to <Invisible Lines> | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            t.RollBack()
            sys.exit()
        t.Commit()
    except Exception as e:
        t.RollBack()
        MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Change to <Invisible Lines> | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        sys.exit()
    finally:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()

# Functie aanroepen
update_line_style()