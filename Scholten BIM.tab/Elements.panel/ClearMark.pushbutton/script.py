# -*- coding: utf-8 -*-

__title__ = "Clear Mark"
__author__ = "Scholten BIM Consultancy"
__doc__ = """Version   = 1.0
Datum    = 05.11.2025
__________________________________________________________________
Description:

Met deze tool kan je op basis van selectie, active view of gehele model parameter Mark leeghalen.
__________________________________________________________________
How-to:
-> Run het script.
__________________________________________________________________
Last update:

- [05.11.2025] - 1.0 RELEASE
__________________________________________________________________
To-do:

- Tips/Tricks?.
__________________________________________________________________
"""

from pyrevit import revit, DB, forms, script
from pyrevit.forms import WPFWindow
import clr
import traceback

# MessageBox
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

doc = revit.doc
uidoc = revit.uidoc


# ------------------------------
# Filter: "ALL_MODEL_MARK is niet leeg"
# ------------------------------
def _nonempty_mark_filter():
    """
    ElementParameterFilter die alleen elementen doorlaat waarvan Mark (ALL_MODEL_MARK)
    niet gelijk is aan een lege string. Werkt op API's met de 3-arg FilterStringRule.
    """
    bip = DB.BuiltInParameter.ALL_MODEL_MARK
    pvp = DB.ParameterValueProvider(DB.ElementId(bip))
    # Evaluator: string equals
    evaluator = DB.FilterStringEquals()
    # Maak regel: Mark == ""  (let op: hier GEEN caseSensitive-argument)
    rule = DB.FilterStringRule(pvp, evaluator, "")
    # Invert de regel -> NOT(Mark == "") => Mark != ""
    return DB.ElementParameterFilter(rule, True)

# ------------------------------
# Collector helpers + snelle telling
# ------------------------------
def _collector_active_view_only_filled():
    av = doc.ActiveView
    if av is None:
        return None
    col = DB.FilteredElementCollector(doc, av.Id).WhereElementIsNotElementType()
    return col.WherePasses(_nonempty_mark_filter())

def _collector_whole_model_only_filled():
    col = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
    return col.WherePasses(_nonempty_mark_filter())

def _collector_count(col):
    if col is None:
        return 0
    try:
        return col.GetElementCount()
    except Exception:
        try:
            return len(list(col))
        except Exception:
            try:
                return len(col.ToElements())
            except Exception:
                return 0


# ------------------------------
# Elementen verzamelen (altijd: Mark gevuld)
# ------------------------------
def collect_selection_only_filled():
    ids = list(uidoc.Selection.GetElementIds())
    if not ids:
        MessageBox.Show(
            "Geen elementen geselecteerd. Selecteer elementen en probeer opnieuw.",
            "Clear Mark | Scholten BIM Consultancy",
            MessageBoxButtons.OK, MessageBoxIcon.Information
        )
        return None
    elems = [doc.GetElement(i) for i in ids]
    out = []
    for e in elems:
        try:
            p = e.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
            if p and p.HasValue:
                s = p.AsString()
                if s and s.strip() != "":
                    out.append(e)
        except Exception as ex:
            try:
                eid = e.Id.IntegerValue if e else "?"
            except Exception:
                eid = "?"
            print("Selectiecontrole fout op {0}: {1}".format(eid, ex))
    return out

def collect_active_view_only_filled():
    col = _collector_active_view_only_filled()
    if col is None:
        MessageBox.Show(
            "Er is geen actieve view beschikbaar.",
            "Clear Mark | Scholten BIM Consultancy",
            MessageBoxButtons.OK, MessageBoxIcon.Information
        )
        return None
    return col.ToElements()

def collect_whole_model_only_filled():
    col = _collector_whole_model_only_filled()
    return col.ToElements()


# ------------------------------
# Clear Mark met voortgang + annuleren (rollback)
# ------------------------------
def clear_mark_with_progress(elements):
    if not elements:
        return 0, False

    total = len(elements)
    changed = 0
    cancelled = False

    t = DB.Transaction(doc, "Clear Mark")
    t.Start()
    try:
        steps = total if total > 0 else 1
        with forms.ProgressBar(title="Clear Mark | Bezig met verwerken...",
                               cancellable=True, step_count=steps) as pb:
            for idx, elem in enumerate(elements, 1):
                # Cancel?
                try:
                    if hasattr(pb, 'cancelled') and pb.cancelled:
                        cancelled = True
                        break
                except Exception:
                    pass

                try:
                    param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
                    if param and param.HasValue:
                        s = param.AsString()
                        if s and s.strip() != "":
                            if not param.IsReadOnly:
                                param.Set("")
                                changed += 1
                            else:
                                print("Mark is read-only voor element {0}".format(
                                    elem.Id.IntegerValue))
                except Exception as e:
                    print("Kon Mark leegmaken op element {0}: {1}".format(
                        elem.Id.IntegerValue, e))

                # Progress
                try:
                    pb.update_progress(idx, steps)
                except Exception:
                    pass
    finally:
        if cancelled:
            try:
                t.RollBack()
            except Exception as ex:
                print("Rollback mislukte: {0}".format(ex))
        else:
            try:
                t.Commit()
            except Exception as ex:
                print("Commit mislukte: {0}".format(ex))

    return changed, cancelled


def show_result_message(changed, cancelled, scope_label):
    if cancelled:
        msg = "Actie geannuleerd.\n\n{0} element(en) waren al leeggemaakt.".format(changed)
    else:
        if changed == 0:
            msg = "Geen elementen met gevulde Mark gevonden in {0}.".format(scope_label)
        elif changed == 1:
            msg = "Er is 1 element aangepast in {0}.".format(scope_label)
        else:
            msg = "Er zijn {0} elementen aangepast in {1}.".format(changed, scope_label)

    MessageBox.Show(
        msg,
        "Clear Mark | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Information
    )


# ------------------------------
# GUI
# ------------------------------
class ClearMarkWindow(WPFWindow):
    def __init__(self, xaml_path):
        WPFWindow.__init__(self, xaml_path)

        # Buttons
        self.btn_run.Click += self.run_click
        self.btn_cancel.Click += self.cancel_click
        self.btn_refresh.Click += self.refresh_count

        # Radio -> live telling
        self.rb_selection.Checked += self._scope_changed
        self.rb_activeview.Checked += self._scope_changed
        self.rb_wholemodel.Checked += self._scope_changed

        # Init
        self._refresh_count_safe()

        # Failsafe
        self.btn_run.IsEnabled = True
        self.btn_cancel.IsEnabled = True
        self.btn_refresh.IsEnabled = True

    # ---------- UI helpers ----------
    def _set_count_text(self, text):
        try:
            self.txt_count.Text = text
        except Exception:
            pass

    def _scope_changed(self, sender, e):
        self._refresh_count_safe()

    def refresh_count(self, sender, e):
        self._refresh_count_safe()

    def _refresh_count_safe(self):
        try:
            if self.rb_selection.IsChecked:
                ids = list(uidoc.Selection.GetElementIds())
                if not ids:
                    self._set_count_text("Telling: 0 (geen selectie)")
                    return
                cnt = 0
                for i in ids:
                    el = doc.GetElement(i)
                    if not el:
                        continue
                    p = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
                    if p and p.HasValue:
                        s = p.AsString()
                        if s and s.strip() != "":
                            cnt += 1
                self._set_count_text("Telling: {0} (Selectie)".format(cnt))

            elif self.rb_activeview.IsChecked:
                col = _collector_active_view_only_filled()
                if col is None:
                    self._set_count_text("Telling: — (geen actieve view)")
                    return
                cnt = _collector_count(col)
                self._set_count_text("Telling: {0} (Active View)".format(cnt))

            else:
                col = _collector_whole_model_only_filled()
                cnt = _collector_count(col)
                self._set_count_text("Telling: {0} (Whole Model)".format(cnt))
        except Exception as ex:
            self._set_count_text("Telling: — (fout)")
            MessageBox.Show(
                "Kon telling niet bepalen:\n\n{0}".format(ex),
                "Clear Mark | Scholten BIM Consultancy",
                MessageBoxButtons.OK, MessageBoxIcon.Warning
            )

    # ---------- Actions ----------
    def cancel_click(self, sender, e):
        self.Close()

    def run_click(self, sender, e):
        try:
            if self.rb_selection.IsChecked:
                elems = collect_selection_only_filled()
                scope_label = "Selectie"
                if elems is None:
                    return
            elif self.rb_activeview.IsChecked:
                elems = collect_active_view_only_filled()
                scope_label = "Active View"
                if elems is None:
                    return
            else:
                elems = collect_whole_model_only_filled()
                scope_label = "Whole Model"

            if not elems:
                MessageBox.Show(
                    "Geen elementen met gevulde Mark gevonden in {0}.".format(scope_label),
                    "Clear Mark | Scholten BIM Consultancy",
                    MessageBoxButtons.OK, MessageBoxIcon.Information
                )
                return

            changed, cancelled = clear_mark_with_progress(elems)
            show_result_message(changed, cancelled, scope_label)
            self.Close()
        except Exception as ex:
            MessageBox.Show(
                "Fout tijdens uitvoeren:\n\n{0}\n\nTraceback:\n{1}".format(
                    ex, traceback.format_exc()),
                "Clear Mark | Scholten BIM Consultancy",
                MessageBoxButtons.OK, MessageBoxIcon.Error
            )


# ------------------------------
# Start GUI
# ------------------------------
try:
    xaml_path = script.get_bundle_file('ClearMark.xaml')
    if not xaml_path:
        MessageBox.Show(
            "ClearMark.xaml niet gevonden in de bundle.\n"
            "Zorg dat ClearMark.xaml in dezelfde map staat als dit script.",
            "Clear Mark | Scholten BIM Consultancy",
            MessageBoxButtons.OK, MessageBoxIcon.Warning
        )
    else:
        ClearMarkWindow(xaml_path).ShowDialog()
except Exception as ex:
    MessageBox.Show(
        "Er ging iets mis bij het openen van het venster:\n\n{0}\n\nTraceback:\n{1}".format(
            ex, traceback.format_exc()),
        "Clear Mark | Scholten BIM Consultancy",
        MessageBoxButtons.OK, MessageBoxIcon.Error
    )