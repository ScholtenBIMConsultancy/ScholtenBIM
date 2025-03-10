import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def adjust_text_width(text_element):
    text = text_element.Text
    # Bereken de breedte op basis van de lengte van de tekst en een lagere factor
    width = len(text) * 0.006  # Lagere factor voor nauwkeurigheid
    text_element.Width = width

try:
    # Start een transactie
    t = Transaction(doc, "Adjust Text Width")
    t.Start()

    # Selecteer de tekstnotities die je wilt aanpassen
    selected_ids = uidoc.Selection.GetElementIds()

    for text_element_id in selected_ids:
        text_element = doc.GetElement(text_element_id)
        if isinstance(text_element, TextElement):
            adjust_text_width(text_element)

    # Eindig de transactie
    t.Commit()
except Exception as e:
    print("Er is een fout opgetreden: {}".format(e))
    t.RollBack()