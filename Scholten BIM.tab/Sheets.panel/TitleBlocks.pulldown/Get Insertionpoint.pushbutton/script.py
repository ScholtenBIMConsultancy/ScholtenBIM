# -*- coding: utf-8 -*-

__title__ = "Get / Set Insertionpoint TitleBlock"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 11.02.2025
__________________________________________________________________
Description:

Met deze tool kan je controleren of de TitleBlock op insertionpoint 0,0,0 staan. Mocht dit niet zo zijn krijg je de vraag of je dit wilt laten verplaatsen. Dit zal dan gebeuren inclusief alle viewports en revisionclouds.
__________________________________________________________________
How-to:

-> Run het script.
__________________________________________________________________
Last update:

- [11.02.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- 
__________________________________________________________________
"""


import clr
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitNodes')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, XYZ
from Autodesk.Revit.UI import TaskDialog
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, DialogResult, MessageBoxIcon

# Haal het huidige document op
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Controleer of er een actieve view is
active_view = doc.ActiveView
if active_view is None:
    MessageBox.Show("Geen actieve view gevonden.", "Get & Set Insertionpoint TitleBlock | Scholten BIM Consultancy")
else:
    # Verzamel alle titleblocks in de actieve view
    collector = FilteredElementCollector(doc, active_view.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)
    titleblocks = collector.ToElements()

    # Controleer of er precies één titleblock is
    if len(titleblocks) == 0:
        MessageBox.Show("Geen titleblock gevonden in de actieve view.", "Get & Set Insertionpoint TitleBlock | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    elif len(titleblocks) > 1:
        MessageBox.Show("Er zijn meerdere titleblocks gevonden in de actieve view. Zorg dat er maximaal 1 titleblock aanwezig is.", "Get & Set Insertionpoint TitleBlock | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
    else:
        # Haal het eerste (en enige) titleblock op
        titleblock = titleblocks[0]

        # Haal het insertion point van het titleblock op
        insertion_point = titleblock.Location.Point

        # Converteer de coördinaten naar millimeters
        insertion_point_mm = (insertion_point.X * 304.8, insertion_point.Y * 304.8, insertion_point.Z * 304.8)

        # Formatteer de coördinaten van het insertion point in millimeters
        insertion_point_str = "X: {0:.2f} mm, Y: {1:.2f} mm, Z: {2:.2f} mm".format(insertion_point_mm[0], insertion_point_mm[1], insertion_point_mm[2])

        # Controleer of de coördinaten al 0 zijn
        if insertion_point_mm[0] == 0 and insertion_point_mm[1] == 0 and insertion_point_mm[2] == 0:
            MessageBox.Show("De waarden van de X, Y en Z as zijn al 0.", "Get & Set Insertionpoint TitleBlock | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
            sys.exit()  # Stop het script als de coördinaten al 0 zijn
        else:
            # Toon de coördinaten in een message box en vraag of de elementen verplaatst moeten worden
            message = "Insertion Point: {0}\n\nWil je de TitleBlock en viewports verplaatsen zodat de waarden van de X, Y en Z as weer 0 zijn?".format(insertion_point_str)
            result = MessageBox.Show(message, "Get & Set Insertionpoint TitleBlock | Scholten BIM Consultancy", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            if result == DialogResult.Yes:
                # Bereken het verschil in coördinaten
                offset = XYZ(-insertion_point.X, -insertion_point.Y, -insertion_point.Z)

                # Start een transactie
                t = Transaction(doc, "Verplaats TitleBlock en Viewports")
                t.Start()

                # Verzamel alle viewports en revision clouds in de actieve view
                viewports = FilteredElementCollector(doc, active_view.Id).OfCategory(BuiltInCategory.OST_Viewports).ToElements()
                revision_clouds = FilteredElementCollector(doc, active_view.Id).OfCategory(BuiltInCategory.OST_RevisionClouds).ToElements()

                # Unpin alle viewports
                for viewport in viewports:
                    if viewport.Pinned:
                        viewport.Pinned = False

                # Verplaats de titleblock, viewports en revision clouds
                titleblock.Location.Move(offset)
                for viewport in viewports:
                    viewport.Location.Move(offset)
                for cloud in revision_clouds:
                    cloud.Location.Move(offset)

                # Pin alle viewports weer
                for viewport in viewports:
                    viewport.Pinned = True

                # Pin de sheet
                active_view.Pinned = True

                # Commit de transactie
                t.Commit()

                MessageBox.Show("TitleBlock, viewports en sheet zijn verplaatst en gepind.", "Get & Set Insertionpoint TitleBlock | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)