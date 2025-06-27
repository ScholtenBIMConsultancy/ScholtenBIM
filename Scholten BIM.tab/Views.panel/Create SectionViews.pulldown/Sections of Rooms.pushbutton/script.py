# -*- coding: utf-8 -*-

__title__ = "Create Sections of Room"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 26.06.2025
__________________________________________________________________
Description:

Met deze tool kan je op basis van te selecteren room(s) 4 sections aan laten maken.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [26.06.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Zorgen dat als de views al bestaan er of een duplicate komt of de ruimte wordt overgeslagen.
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
uidoc = __revit__.ActiveUIDocument# -*- coding: utf-8 -*-

__title__ = "Create Room Sections"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 26.06.2025
__________________________________________________________________
Description:

Met deze tool kan je op basis van te selecteren room(s) 4 sections aan laten maken.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [26.06.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Zorgen dat als de views al bestaan er of een duplicate komt of de ruimte wordt overgeslagen.
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

# Instellingen
FAR_CLIP_MARGIN = 200.0 / 304.8  # 200mm voor ver-clip in feet
VERTICAL_MARGIN = 200.0 / 304.8  # 200mm voor onder- en bovenclip in feet


def get_room_from_linked_element(reference):
    linked_inst = doc.GetElement(reference.ElementId)
    if not isinstance(linked_inst, RDB.RevitLinkInstance):
        return None
    linked_doc = linked_inst.GetLinkDocument()
    elem = linked_doc.GetElement(reference.LinkedElementId)
    if elem.Category is None or elem.Category.Id != RDB.ElementId(RDB.BuiltInCategory.OST_Rooms):
        return None
    return elem


def create_elevation_view(room, direction, suffix):
    try:
        num_param = room.get_Parameter(RDB.BuiltInParameter.ROOM_NUMBER)
        room_number = num_param.AsString() if num_param else str(room.Id.IntegerValue)
        name_param = room.get_Parameter(RDB.BuiltInParameter.ROOM_NAME)
        room_name = name_param.AsString() if name_param else ''

        if room_name:
            view_name = '{0} {1}_{2}'.format(room_number, room_name, suffix)
        else:
            view_name = '{0}_{1}'.format(room_number, suffix)

        bbox = room.get_BoundingBox(None)
        if not bbox:
            output.print_md('⚠️ Geen bounding box voor Room {0}'.format(room.Id.IntegerValue))
            return None
        center = RDB.XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0
        )

        view_dir = RDB.XYZ(direction.X, direction.Y, direction.Z).Normalize()
        up = RDB.XYZ.BasisZ
        right = up.CrossProduct(view_dir).Normalize()

        transform = RDB.Transform.Identity
        transform.Origin = center
        transform.BasisX = right
        transform.BasisY = up
        transform.BasisZ = view_dir

        width = abs(bbox.Max.X - bbox.Min.X)
        length = abs(bbox.Max.Y - bbox.Min.Y)
        height = abs(bbox.Max.Z - bbox.Min.Z)

        if abs(direction.X) > 0:  # Oost-West
            hor_ext = length / 2.0
            view_depth = width / 2.0 + FAR_CLIP_MARGIN
        else:  # Noord-Zuid
            hor_ext = width / 2.0
            view_depth = length / 2.0 + FAR_CLIP_MARGIN

        half_height = height / 2.0

        min_pt = RDB.XYZ(-hor_ext - FAR_CLIP_MARGIN, -half_height - VERTICAL_MARGIN, 0.0)
        max_pt = RDB.XYZ(hor_ext + FAR_CLIP_MARGIN, half_height + VERTICAL_MARGIN, view_depth)

        section_box = RDB.BoundingBoxXYZ()
        section_box.Transform = transform
        section_box.Min = min_pt
        section_box.Max = max_pt
        try:
            section_box.MinEnabled = True
            section_box.MaxEnabled = True
        except:
            pass

        vft = next(
            (v for v in RDB.FilteredElementCollector(doc)
                   .OfClass(RDB.ViewFamilyType)
                   .WhereElementIsElementType()
             if v.ViewFamily == RDB.ViewFamily.Section),
            None
        )
        if not vft:
            output.print_md('❌ Geen Section ViewFamilyType gevonden.')
            return None

        view = RDB.ViewSection.CreateSection(doc, vft.Id, section_box)
        name_p = view.get_Parameter(RDB.BuiltInParameter.VIEW_NAME)
        if name_p:
            name_p.Set(view_name)

        try:
            offset_param = view.get_Parameter(RDB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
            if offset_param and offset_param.StorageType == RDB.StorageType.Double:
                offset_param.Set(view_depth)
        except:
            output.print_md('⚠️ Kon VIEWER_BOUND_OFFSET_FAR niet instellen.')

        output.print_md('✅ Section aangemaakt: {0}'.format(view_name))
        return view

    except Exception as e:
        import traceback
        traceback.print_exc()
        output.print_md('❌ Fout bij Room {0}: {1}: {2}'.format(room.Id.IntegerValue, type(e).__name__, e))
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
    with revit.Transaction('Create Room Sections'):
        for ref in selected:
            room = get_room_from_linked_element(ref)
            if not room:
                continue
            directions = [RDB.XYZ(1, 0, 0), RDB.XYZ(0, 1, 0), RDB.XYZ(-1, 0, 0), RDB.XYZ(0, -1, 0)]
            for i, direction in enumerate(directions):
                create_elevation_view(room, direction, chr(65 + i))
output = script.get_output()

# Instellingen
FAR_CLIP_MARGIN = 200.0 / 304.8  # 200mm voor ver-clip in feet
VERTICAL_MARGIN = 200.0 / 304.8  # 200mm voor onder- en bovenclip in feet


def get_room_from_linked_element(reference):
    linked_inst = doc.GetElement(reference.ElementId)
    if not isinstance(linked_inst, RDB.RevitLinkInstance):
        return None
    linked_doc = linked_inst.GetLinkDocument()
    elem = linked_doc.GetElement(reference.LinkedElementId)
    if elem.Category is None or elem.Category.Id != RDB.ElementId(RDB.BuiltInCategory.OST_Rooms):
        return None
    return elem


def create_elevation_view(room, direction, suffix):
    try:
        num_param = room.get_Parameter(RDB.BuiltInParameter.ROOM_NUMBER)
        room_number = num_param.AsString() if num_param else str(room.Id.IntegerValue)
        name_param = room.get_Parameter(RDB.BuiltInParameter.ROOM_NAME)
        room_name = name_param.AsString() if name_param else ''

        if room_name:
            view_name = '{0} {1}_{2}'.format(room_number, room_name, suffix)
        else:
            view_name = '{0}_{1}'.format(room_number, suffix)

        bbox = room.get_BoundingBox(None)
        if not bbox:
            output.print_md('⚠️ Geen bounding box voor Room {0}'.format(room.Id.IntegerValue))
            return None
        center = RDB.XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0
        )

        view_dir = RDB.XYZ(direction.X, direction.Y, direction.Z).Normalize()
        up = RDB.XYZ.BasisZ
        right = up.CrossProduct(view_dir).Normalize()

        transform = RDB.Transform.Identity
        transform.Origin = center
        transform.BasisX = right
        transform.BasisY = up
        transform.BasisZ = view_dir

        width = abs(bbox.Max.X - bbox.Min.X)
        length = abs(bbox.Max.Y - bbox.Min.Y)
        height = abs(bbox.Max.Z - bbox.Min.Z)

        if abs(direction.X) > 0:  # Oost-West
            hor_ext = length / 2.0
            view_depth = width / 2.0 + FAR_CLIP_MARGIN
        else:  # Noord-Zuid
            hor_ext = width / 2.0
            view_depth = length / 2.0 + FAR_CLIP_MARGIN

        half_height = height / 2.0

        min_pt = RDB.XYZ(-hor_ext - FAR_CLIP_MARGIN, -half_height - VERTICAL_MARGIN, 0.0)
        max_pt = RDB.XYZ(hor_ext + FAR_CLIP_MARGIN, half_height + VERTICAL_MARGIN, view_depth)

        section_box = RDB.BoundingBoxXYZ()
        section_box.Transform = transform
        section_box.Min = min_pt
        section_box.Max = max_pt
        try:
            section_box.MinEnabled = True
            section_box.MaxEnabled = True
        except:
            pass

        vft = next(
            (v for v in RDB.FilteredElementCollector(doc)
                   .OfClass(RDB.ViewFamilyType)
                   .WhereElementIsElementType()
             if v.ViewFamily == RDB.ViewFamily.Section),
            None
        )
        if not vft:
            output.print_md('❌ Geen Section ViewFamilyType gevonden.')
            return None

        view = RDB.ViewSection.CreateSection(doc, vft.Id, section_box)
        name_p = view.get_Parameter(RDB.BuiltInParameter.VIEW_NAME)
        if name_p:
            name_p.Set(view_name)

        try:
            offset_param = view.get_Parameter(RDB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
            if offset_param and offset_param.StorageType == RDB.StorageType.Double:
                offset_param.Set(view_depth)
        except:
            output.print_md('⚠️ Kon VIEWER_BOUND_OFFSET_FAR niet instellen.')

        output.print_md('✅ Section aangemaakt: {0}'.format(view_name))
        return view

    except Exception as e:
        import traceback
        traceback.print_exc()
        output.print_md('❌ Fout bij Room {0}: {1}: {2}'.format(room.Id.IntegerValue, type(e).__name__, e))
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
    with revit.Transaction('Create Room Sections'):
        for ref in selected:
            room = get_room_from_linked_element(ref)
            if not room:
                continue
            directions = [RDB.XYZ(1, 0, 0), RDB.XYZ(0, 1, 0), RDB.XYZ(-1, 0, 0), RDB.XYZ(0, -1, 0)]
            for i, direction in enumerate(directions):
                create_elevation_view(room, direction, chr(65 + i))