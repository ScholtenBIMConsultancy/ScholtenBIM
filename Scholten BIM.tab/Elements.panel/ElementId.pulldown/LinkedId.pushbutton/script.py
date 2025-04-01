# -*- coding: utf-8 -*-

__title__ = "Get Linked Id's"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 31.03.2025
__________________________________________________________________
Description:

Met deze tool kan je element id's uitlezen van gelinkte elementen.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [31.03.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, RevitLinkInstance, FamilyInstance, BuiltInParameter
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import output, forms

# Actief document en view ophalen
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# Functie om gelinkte elementen op te halen
def get_linked_element(doc, ref):
    link_instance = doc.GetElement(ref)
    if isinstance(link_instance, RevitLinkInstance):
        link_doc = link_instance.GetLinkDocument()
        link_element_id = ref.LinkedElementId
        return link_doc.GetElement(link_element_id)
    return None

# Gelinkte bestanden ophalen
linked_docs = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()

# Gebruik de WarningBar om elementen te selecteren
with forms.WarningBar(title="Pick element(s)"):
    try:
        selected_refs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Selecteer elementen in het gelinkte bestand")
    except:
        # Script stoppen zonder foutmelding als selectie wordt geannuleerd
        selected_refs = []

# ID's van geselecteerde elementen ophalen
selected_elements = [get_linked_element(doc, ref) for ref in selected_refs if get_linked_element(doc, ref)]

# Element-ID's weergeven met output_window in het gewenste formaat
if selected_elements:
    output_window = output.get_output()
    output_window.print_md("##==== Linked Element Details: ====##")
    for elem in selected_elements:
        category = elem.Category.Name
        family_type_name_param = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
        family_type_name = family_type_name_param.AsValueString() if family_type_name_param else "N/A"
        linked_id = elem.Id
        output_window.print_md("**Category:** {}, **Family (Type) Name:** {}".format(category, family_type_name))
        output_window.print_md("**Linked Element ID:** {}".format(linked_id))
        output_window.print_md("=" * 100)