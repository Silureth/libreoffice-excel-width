"""Show / set LibreOffice Calc column widths in Excel (and NPOI) units.

Macros:
  excel_width_panel  - small always-open window that follows the selected
                       column(s) live and lets you type a width in Excel units
  show_excel_width   - one-shot message box
"""
import unohelper
from com.sun.star.awt.MessageBoxType import INFOBOX
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.util import XModifyListener
from com.sun.star.awt import XActionListener, XTopWindowListener, XKeyListener
from com.sun.star.awt.Key import RETURN

MDW_PX = 7       # max digit width: Calibri/Carlito 11pt @ 96 DPI
PADDING_PX = 5   # Excel adds 5 px of cell padding to the visible width

_panels = {}     # keep panels alive, keyed by document


def _doc_key(doc):
    try:
        return doc.RuntimeUID
    except Exception:
        return doc.URL or str(doc.Title)


# ---------------------------------------------------------------- conversion
def _px(hmm):                    # 1/100 mm -> px @ 96 DPI
    return round(hmm / 2540.0 * 96)


def _hmm(px):                    # px @ 96 DPI -> 1/100 mm
    return round(px * 2540.0 / 96)


def _excel_ui(px):
    return max(0.0, (px - PADDING_PX) / MDW_PX)


def _npoi(px):
    return int(px * 256 / MDW_PX)    # Excel truncates


def _px_from_excel_ui(ui):
    return round(ui * MDW_PX) + PADDING_PX


# ---------------------------------------------------------------- helpers
def _selected_columns(doc):
    """(sheet, [column indexes]) for the current selection, or (None, [])."""
    sel = doc.CurrentSelection
    if sel is None:
        return None, []
    if hasattr(sel, "RangeAddresses"):        # multi-selection
        addrs = list(sel.RangeAddresses)
    elif hasattr(sel, "RangeAddress"):
        addrs = [sel.RangeAddress]
    else:                                     # shape, chart, ...
        return None, []
    if not addrs:
        return None, []
    sheet = doc.Sheets.getByIndex(addrs[0].Sheet)
    cols = sorted({c for a in addrs for c in range(a.StartColumn, a.EndColumn + 1)})
    return sheet, cols


def _describe(doc):
    sheet, idx = _selected_columns(doc)
    if not idx:
        return "No cells selected."
    columns = sheet.Columns
    rows = []
    for c in idx[:200]:
        col = columns.getByIndex(c)
        rows.append((col.Name, _px(col.Width)))
    widths = {px for _, px in rows}
    if len(widths) == 1:
        px = rows[0][1]
        name = rows[0][0] if len(idx) == 1 else f"{rows[0][0]}:{rows[-1][0]} ({len(idx)} cols)"
        return f"{name}\nExcel {_excel_ui(px):.2f}   NPOI {_npoi(px)}   {px} px"
    lines = [f"{n}: Excel {_excel_ui(px):.2f}   NPOI {_npoi(px)}   {px} px" for n, px in rows[:4]]
    if len(idx) > 4:
        lines.append(f"... {len(idx) - 4} more (widths differ)")
    return "\n".join(lines)


# ---------------------------------------------------------------- panel
class _Panel(unohelper.Base, XSelectionChangeListener, XModifyListener,
             XActionListener, XTopWindowListener, XKeyListener):

    def __init__(self, doc, ctx):
        self.doc = doc
        self.key = _doc_key(doc)
        self.controller = doc.CurrentController
        smgr = ctx.ServiceManager

        dm = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        dm.Title = "Excel column width"
        dm.Width, dm.Height = 190, 72
        dm.Closeable = True
        dm.Moveable = True

        def add(kind, name, x, y, w, h, **props):
            m = dm.createInstance("com.sun.star.awt.UnoControl%sModel" % kind)
            m.PositionX, m.PositionY, m.Width, m.Height = x, y, w, h
            for k, v in props.items():
                setattr(m, k, v)
            dm.insertByName(name, m)

        add("FixedText", "info", 4, 4, 182, 44, MultiLine=True)
        add("FixedText", "lbl", 4, 55, 34, 12, Label="Set Excel:")
        add("Edit", "input", 38, 53, 40, 14)
        add("Button", "set", 82, 53, 40, 14, Label="Apply")
        add("Button", "refresh", 126, 53, 40, 14, Label="Refresh")

        self.dlg = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        self.dlg.setModel(dm)
        tk = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        self.dlg.createPeer(tk, self.controller.Frame.ContainerWindow)

        self.dlg.getControl("set").setActionCommand("set")
        self.dlg.getControl("refresh").setActionCommand("refresh")
        self.dlg.getControl("set").addActionListener(self)
        self.dlg.getControl("refresh").addActionListener(self)
        self.dlg.getControl("input").addKeyListener(self)
        self.dlg.addTopWindowListener(self)

        self.controller.addSelectionChangeListener(self)
        self.doc.addModifyListener(self)

        self.update()
        self.dlg.setVisible(True)

    # -- refresh the label
    def update(self):
        try:
            text = _describe(self.doc)
        except Exception as e:                # never let a listener blow up
            text = "Error: %s" % e
        self.dlg.getControl("info").getModel().Label = text

    def apply(self):
        raw = self.dlg.getControl("input").getText().strip().replace(",", ".")
        try:
            ui = float(raw)
        except ValueError:
            return
        sheet, idx = _selected_columns(self.doc)
        if not idx:
            return
        hmm = _hmm(_px_from_excel_ui(ui))
        for c in idx:
            sheet.Columns.getByIndex(c).Width = hmm
        self.update()

    def close(self):
        _panels.pop(self.key, None)
        for fn in (lambda: self.controller.removeSelectionChangeListener(self),
                   lambda: self.doc.removeModifyListener(self),
                   lambda: self.dlg.setVisible(False),
                   lambda: self.dlg.dispose()):
            try:
                fn()
            except Exception:
                pass

    # -- XSelectionChangeListener / XModifyListener
    def selectionChanged(self, ev):
        self.update()

    def modified(self, ev):
        self.update()

    def disposing(self, ev):                  # document or view closed
        self.close()

    # -- XActionListener
    def actionPerformed(self, ev):
        if ev.ActionCommand == "set":
            self.apply()
        else:
            self.update()

    # -- XKeyListener (Enter in the input box = Apply)
    def keyPressed(self, ev):
        if ev.KeyCode == RETURN:
            self.apply()

    def keyReleased(self, ev):
        pass

    # -- XTopWindowListener
    def windowClosing(self, ev):
        self.close()

    def windowOpened(self, ev): pass
    def windowClosed(self, ev): pass
    def windowMinimized(self, ev): pass
    def windowNormalized(self, ev): pass
    def windowActivated(self, ev): pass
    def windowDeactivated(self, ev): pass


# ---------------------------------------------------------------- macros
def excel_width_panel(*args):
    doc = XSCRIPTCONTEXT.getDocument()
    old = _panels.get(_doc_key(doc))
    if old is not None:                       # toggle: second call closes it
        old.close()
        return
    panel = _Panel(doc, XSCRIPTCONTEXT.getComponentContext())
    _panels[panel.key] = panel


def show_excel_width(*args):
    doc = XSCRIPTCONTEXT.getDocument()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    tk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    box = tk.createMessageBox(doc.CurrentController.Frame.ContainerWindow,
                              INFOBOX, 1, "Excel column width", _describe(doc))
    box.execute()


g_exportedScripts = (excel_width_panel, show_excel_width)
