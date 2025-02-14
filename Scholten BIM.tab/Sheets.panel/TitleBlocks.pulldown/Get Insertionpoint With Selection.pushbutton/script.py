# -*- coding: utf-8 -*-

__title__ = "Get / Set Insertionpoint TitleBlock with Selection"
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
clr.AddReference('pyrevit')

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, XYZ, ViewSheet, ElementTransformUtils, ScheduleSheetInstance, LocationPoint
from Autodesk.Revit.UI import TaskDialog
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult
from pyrevit import forms

# Haal het huidige document op
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Verzamel alle sheets in het document
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

# Maak een lijst van sheet nummers en namen
sheet_display_names = ["{} - {}".format(sheet.SheetNumber, sheet.Name) for sheet in sheets]

# Sorteer de lijst van sheet namen op alfabetische volgorde
sheet_display_names.sort()

# Toon het formulier en haal de geselecteerde sheets op
selected_sheet_display_names = forms.SelectFromList.show(sheet_display_names, multiselect=True, title="Selecteer Sheets")

if selected_sheet_display_names:
    selected_sheets = [sheet for sheet in sheets if "{} - {}".format(sheet.SheetNumber, sheet.Name) in selected_sheet_display_names]

    count_zero_insertion = 0  # Teller voor sheets met (0,0,0) coördinaten
    sheets_modified = False  # Vlag om bij te houden of er sheets zijn aangepast
    sheets_with_deviation = []  # Lijst om sheets met afwijking op te slaan

    for selected_sheet in selected_sheets:
        # Verzamel alle titleblocks in de geselecteerde sheet
        collector = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)
        titleblocks = collector.ToElements()

        # Controleer of er precies één titleblock is
        if len(titleblocks) == 0:
            MessageBox.Show("Geen titleblock gevonden in de geselecteerde sheet.", "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        elif len(titleblocks) > 1:
            MessageBox.Show("Er zijn meerdere titleblocks gevonden in de geselecteerde sheet. Zorg dat er maximaal 1 titleblock aanwezig is.", "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)
        else:
            # Haal het eerste (en enige) titleblock op
            titleblock = titleblocks[0]

            # Controleer of het titleblock een LocationPoint heeft
            if isinstance(titleblock.Location, LocationPoint):
                # Haal het insertion point van het titleblock op
                insertion_point = titleblock.Location.Point

                # Converteer de coördinaten naar millimeters
                insertion_point_mm = (insertion_point.X * 304.8, insertion_point.Y * 304.8, insertion_point.Z * 304.8)

                # Controleer of de coördinaten al 0 zijn
                if insertion_point_mm[0] == 0 and insertion_point_mm[1] == 0 and insertion_point_mm[2] == 0:
                    count_zero_insertion += 1
                else:
                    sheets_with_deviation.append("{} - {}".format(selected_sheet.SheetNumber, selected_sheet.Name))

    if sheets_with_deviation:
        # Toon de coördinaten in een message box en vraag of de elementen verplaatst moeten worden
        message = "Onderstaande sheets staan niet op 0,0,0, wil je elementen verplaatsen zodat de waarden weer 0 zijn?\n\n• {}".format("\n• ".join(sheets_with_deviation))
        result = MessageBox.Show(message, "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        if result == DialogResult.Yes:
            # Start een transactie
            t = Transaction(doc, "Verplaats en Pin Elementen")
            t.Start()

            try:
                for selected_sheet in selected_sheets:
                    # Verzamel alle titleblocks in de geselecteerde sheet
                    collector = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)
                    titleblocks = collector.ToElements()

                    # Controleer of er precies één titleblock is
                    if len(titleblocks) == 1:
                        # Haal het eerste (en enige) titleblock op
                        titleblock = titleblocks[0]

                        # Controleer of het titleblock een LocationPoint heeft
                        if isinstance(titleblock.Location, LocationPoint):
                            # Haal het insertion point van het titleblock op
                            insertion_point = titleblock.Location.Point

                            # Converteer de coördinaten naar millimeters
                            insertion_point_mm = (insertion_point.X * 304.8, insertion_point.Y * 304.8, insertion_point.Z * 304.8)

                            # Controleer of de coördinaten al 0 zijn
                            if not (insertion_point_mm[0] == 0 and insertion_point_mm[1] == 0 and insertion_point_mm[2] == 0):
                                # Bereken het verschil in coördinaten
                                offset_mm = XYZ(-insertion_point.X, -insertion_point.Y, -insertion_point.Z)

                                # Verzamel alle viewports, revision clouds en schedules op de sheet
                                viewports = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_Viewports).ToElements()
                                revision_clouds = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_RevisionClouds).ToElements()
                                schedules = FilteredElementCollector(doc, selected_sheet.Id).OfClass(ScheduleSheetInstance).ToElements()

                                # Unpin alle elementen
                                titleblock.Pinned = False
                                for viewport in viewports:
                                    viewport.Pinned = False
                                for cloud in revision_clouds:
                                    cloud.Pinned = False
                                for schedule in schedules:
                                    schedule.Pinned = False

                                # Verplaats de titleblock, viewports, revision clouds en losstaande schedules
                                ElementTransformUtils.MoveElement(doc, titleblock.Id, offset_mm)
                                for viewport in viewports:
                                    ElementTransformUtils.MoveElement(doc, viewport.Id, offset_mm)
                                for cloud in revision_clouds:
                                    ElementTransformUtils.MoveElement(doc, cloud.Id, offset_mm)
                                for schedule in schedules:
                                    # Controleer of de schedule zich binnen het title block bevindt
                                    schedule_location = schedule.Location.Point
                                    if not (insertion_point.X - 0.1 < schedule_location.X < insertion_point.X + 0.1 and
                                            insertion_point.Y - 0.1 < schedule_location.Y < insertion_point.Y + 0.1):
                                        ElementTransformUtils.MoveElement(doc, schedule.Id, offset_mm)

                                sheets_modified = True  # Markeer dat er sheets zijn aangepast

                # Pin alle elementen op de geselecteerde sheets
                for selected_sheet in selected_sheets:
                    titleblocks = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).ToElements()
                    viewports = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_Viewports).ToElements()
                    revision_clouds = FilteredElementCollector(doc, selected_sheet.Id).OfCategory(BuiltInCategory.OST_RevisionClouds).ToElements()
                    schedules = FilteredElementCollector(doc, selected_sheet.Id).OfClass(ScheduleSheetInstance).ToElements()

                    for titleblock in titleblocks:
                        titleblock.Pinned = True
                    for viewport in viewports:
                        viewport.Pinned = True
                    for cloud in revision_clouds:
                        cloud.Pinned = True
                    for schedule in schedules:
                        schedule.Pinned = True

                # Commit de transactie
                t.Commit()

                MessageBox.Show("Elementen op de sheets zijn verplaatst en gepind.", "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)

            except Exception as e:
                t.RollBack()
                MessageBox.Show("Er is een fout opgetreden: {}".format(e), "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Error)

    # Toon de juiste melding afhankelijk van de situatie
    if count_zero_insertion == len(selected_sheets):
        MessageBox.Show("De geselecteerde sheets staan al gepositioneerd op 0,0,0.", "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)
    elif sheets_modified:
        MessageBox.Show("De overige sheets staan al gepositioneerd op 0,0,0.", "Get & Set Insertionpoint TitleBlock With Selection | Scholten BIM Consultancy", MessageBoxButtons.OK, MessageBoxIcon.Information)