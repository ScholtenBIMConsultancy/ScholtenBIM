# -*- coding: utf-8 -*-

__title__ = "Create Floorplan of Room"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 26.06.2025
__________________________________________________________________
Description:

Met deze tool kan je op basis van te selecteren room(s) een floorplan aan laten maken.
__________________________________________________________________
How-to:

-> Run het script en selecteer room(s) in het gelinkte model.
__________________________________________________________________
Last update:

- [26.06.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Eventueel view-template instellen
- Zorgen dat alleen een room kan worden geselecteerd uit linked model.
__________________________________________________________________
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk.Revit.DB as RDB
from pyrevit import revit, forms, script
from Autodesk.Revit.UI.Selection import ObjectType

# Document en UI-document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

OFFSET = 200.0 / 304.8  # 200mm offset in feet


def get_room_from_linked_element(reference):
    linked_inst = doc.GetElement(reference.ElementId)
    if not isinstance(linked_inst, RDB.RevitLinkInstance):
        return None, None, None
    linked_doc = linked_inst.GetLinkDocument()
    room = linked_doc.GetElement(reference.LinkedElementId)
    if room.Category is None or room.Category.Id != RDB.ElementId(RDB.BuiltInCategory.OST_Rooms):
        return None, None, None
    return room, linked_doc, linked_inst


def ensure_unique_name(base_name):
    collector = RDB.FilteredElementCollector(doc).OfClass(RDB.View)
    names = {v.Name for v in collector}
    name = base_name
    counter = 1
    while name in names:
        name = "{} Copy {}".format(base_name, counter)
        counter += 1
    return name


def create_floorplan_view(room, linked_doc, linked_inst):
    try:
        # Room nummer en naam
        num_param = room.get_Parameter(RDB.BuiltInParameter.ROOM_NUMBER)
        room_number = num_param.AsString() if num_param else str(room.Id.IntegerValue)
        name_param = room.get_Parameter(RDB.BuiltInParameter.ROOM_NAME)
        room_name = name_param.AsString() if name_param else ''

        # View naam
        if room_name:
            base_name = "{} {}".format(room_number, room_name)
        else:
            base_name = room_number
        view_name = ensure_unique_name(base_name)

        # Probeer level via linked model
        level = None
        try:
            linked_level = linked_doc.GetElement(room.LevelId)
            if isinstance(linked_level, RDB.Level):
                # Zoek eerder gemaakte Level in host
                level_name = linked_level.Name
                level = next(
                    (l for l in RDB.FilteredElementCollector(doc).OfClass(RDB.Level)
                     if l.Name == level_name),
                    None
                )
        except:
            level = None

        # Fallback op active view level
        if not level:
            active_view = uidoc.ActiveView
            try:
                gen_level_param = active_view.GenLevel
                if gen_level_param:
                    level = doc.GetElement(gen_level_param.Id)
                    output.print_md("ℹ️ Gebruik Level van actieve view: {}".format(level.Name))
            except:
                level = None

        if not level:
            output.print_md("❌ Kan Level niet bepalen voor Room {}".format(room.Id.IntegerValue))
            return None

        # FloorPlan view type
        vft = next(
            (v for v in RDB.FilteredElementCollector(doc)
                   .OfClass(RDB.ViewFamilyType)
                   .WhereElementIsElementType()
             if v.ViewFamily == RDB.ViewFamily.FloorPlan),
            None
        )
        if not vft:
            output.print_md('❌ Geen FloorPlan ViewFamilyType gevonden.')
            return None

        # Create view
        view = RDB.ViewPlan.Create(doc, vft.Id, level.Id)
        name_p = view.get_Parameter(RDB.BuiltInParameter.VIEW_NAME)
        if name_p:
            name_p.Set(view_name)

        # Bepaal bounding box van room in host coords
        bbox = room.get_BoundingBox(None)
        transform = linked_inst.GetTotalTransform()
        min_pt = transform.OfPoint(bbox.Min)
        max_pt = transform.OfPoint(bbox.Max)
        crop_min = RDB.XYZ(min_pt.X - OFFSET, min_pt.Y - OFFSET, min_pt.Z)
        crop_max = RDB.XYZ(max_pt.X + OFFSET, max_pt.Y + OFFSET, max_pt.Z)

        # Instellen crop box
        crop_box = RDB.BoundingBoxXYZ()
        crop_box.Min = crop_min
        crop_box.Max = crop_max
        view.CropBox = crop_box
        view.CropBoxActive = True
        view.CropBoxVisible = False

        output.print_md('✅ FloorPlan aangemaakt en gecropt: {}'.format(view_name))
        return view

    except Exception as e:
        import traceback
        traceback.print_exc()
        output.print_md('❌ Fout bij Room {}: {}: {}'.format(room.Id.IntegerValue, type(e).__name__, e))
        return None

# Selectie UI
with forms.WarningBar(title='Selecteer ruimte(s) in gelinkte model(len)'):
    try:
        selected = uidoc.Selection.PickObjects(ObjectType.LinkedElement, 'Select Room(s)')
    except:
        forms.alert('Geen selectie gemaakt. Script gestopt.')
        selected = []

# Verwerking selectie
if selected:
    with revit.Transaction('Create Room Floorplans'):
        for ref in selected:
            room, linked_doc, linked_inst = get_room_from_linked_element(ref)
            if not room:
                continue
            create_floorplan_view(room, linked_doc, linked_inst)
