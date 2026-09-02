"""TRI Label Maker - enter and print Avery 5209 file folder labels.

See CLAUDE.md for the authoritative sheet geometry and functional spec.
"""

import copy
import json
import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, simpledialog, ttk

try:
    import win32con
    import win32gui
    import win32print
    import win32ui
    HAVE_WIN32 = True
except ImportError:
    HAVE_WIN32 = False

APP_NAME = "TRI Label Maker"
APP_VERSION = "1.2"
ICON_FILE = "TRI_LabelMaker_icon.ico"
TUTORIAL_FILE = "TRI_LabelMaker_tutorial_narrated.wmv"

# ---------------------------------------------------------------------------
# Avery 5209 geometry (inches) - authoritative values from CLAUDE.md
# ---------------------------------------------------------------------------
SHEET_W = 4.0
SHEET_H = 6.0
LABEL_W = 3.4375
LABEL_H = 2.0 / 3.0
NUM_LABELS = 7
MARGIN_TOP = 0.65    # sheet top margin above label 1
MARGIN_LEFT = 0.3    # sheet left margin
MARGIN_RIGHT = SHEET_W - MARGIN_LEFT - LABEL_W  # 0.2625, right edge to labels

# Alignment test page (plain letter paper, 8.5x11): the 4x6 sheet outline is
# drawn centered horizontally and dropped below the unprintable top edge.
ALIGN_PAGE_W = 8.5
ALIGN_BOX_TOP = 0.25
ALIGN_BOX_LEFT = (ALIGN_PAGE_W - SHEET_W) / 2   # 2.25
ALIGN_TARGET_LR = ALIGN_BOX_LEFT                # 2.25 on both sides
ALIGN_TARGET_BOTTOM = ALIGN_BOX_TOP + SHEET_H   # 6.25 from the paper top

FONT_NAME = "Lucida Sans Typewriter"
FALLBACK_FONT = "Courier New"
AUTO_MAX_PT = 40
AUTO_MIN_PT = 4
MAX_LINES = 2
TAB_GAP = 0.15       # minimum gap between the two tab columns (inches)
V_PAD = 0.03         # vertical inset inside a label (inches)
PREVIEW_SCALE = 90   # preview canvas pixels per inch

# Dark theme palette (visual reference: DarkThemeStyle.png)
CLR_BG = "#1f1f1f"             # window background
CLR_HEADER = "#161616"         # header / status bar
CLR_ROW = "#2e2e2e"            # idle label row
CLR_ROW_ACTIVE = "#3a3a3a"     # focused label row
CLR_FIELD = "#2e2e2e"          # entry/listbox background
CLR_BTN = "#383838"            # secondary button
CLR_BTN_HOVER = "#484848"
CLR_FG = "#e6e6e6"             # primary text
CLR_DIM = "#9b9b9b"            # secondary text
CLR_ACCENT = "#f0a638"         # amber accent
CLR_ACCENT_ACTIVE = "#ffb84d"
CLR_ON_ACCENT = "#1a1a1a"      # text on amber
CLR_PREVIEW_BG = "#191919"     # preview panel backdrop

# Entry rows mirror the preview's vertical geometry so row N sits level with
# label N on the sheet: same pitch, same top offset (see _build_ui).
ROW_PITCH = round(LABEL_H * PREVIEW_SCALE)   # 60 px per label row
ROW_GAP = 2                                  # visual gap inside the pitch
PREVIEW_PAD = 14                             # canvas inset around the sheet

DEFAULT_STYLE = {
    "font": FONT_NAME,    # per-label font family
    "h_align": "center",  # left | center | right | tab
    "v_align": "middle",  # top | middle | bottom
    "margin": 0.10,       # inches inset used by left/right/tab alignment
    "gutter": 0.0,        # vertical gap between line 1 and line 2 (inches)
    # one entry per line (labels hold MAX_LINES lines): size 0 = auto-size
    "lines": [{"size": 0, "bold": False}, {"size": 0, "bold": False}],
}

DEFAULT_FORMATS = {
    "1": {"name": "Job folders", "style": copy.deepcopy(DEFAULT_STYLE)},
    "2": {"name": "Vendor folders", "style": copy.deepcopy(DEFAULT_STYLE)},
    "3": {"name": "Commercial folders", "style": copy.deepcopy(DEFAULT_STYLE)},
}
FORMAT_KEYS = ("1", "2", "3")

CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")), "TRI Label Maker")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")

# DeviceCapabilities / DEVMODE constants (defined locally so a missing
# win32con attribute can never break printing)
DC_PAPERS = 2
DC_PAPERSIZE = 3
DMPAPER_USER = 256
DMPAPER_LETTER = 1
DM_ORIENTATION = 0x1
DM_PAPERSIZE = 0x2
DM_PAPERLENGTH = 0x4
DM_PAPERWIDTH = 0x8
PHYSICALOFFSETX = 112
PHYSICALOFFSETY = 113


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def is_blank(text):
    return text.strip() == ""


def normalize_style(style):
    """Return a complete style dict, upgrading pre-per-line styles
    (which had a single label-wide "font_size") to the current shape."""
    s = copy.deepcopy(DEFAULT_STYLE)
    if not isinstance(style, dict):
        return s
    for key in ("h_align", "v_align", "margin", "gutter"):
        if key in style:
            s[key] = style[key]
    if style.get("font"):
        s["font"] = str(style["font"])
    old_lines = style.get("lines")
    if isinstance(old_lines, list) and old_lines:
        for i in range(MAX_LINES):
            src = old_lines[min(i, len(old_lines) - 1)]
            if isinstance(src, dict):
                s["lines"][i] = {"size": int(src.get("size", 0) or 0),
                                 "bold": bool(src.get("bold", False))}
    else:
        old_size = int(style.get("font_size", 0) or 0)
        for line in s["lines"]:
            line["size"] = old_size
    return s


# ---------------------------------------------------------------------------
# Layout engine - shared by the screen preview and the printer so both stay
# WYSIWYG-consistent. A "measurer" maps a point size to width/line-height in
# inches; TkMeasurer backs the preview, GdiMeasurer backs the printer DC.
# ---------------------------------------------------------------------------

def _line_fits(line, style, m, usable_w):
    if style.get("h_align") == "tab" and "\t" in line:
        left, right = line.split("\t", 1)
        return m.width(left) + TAB_GAP + m.width(right.replace("\t", " ")) <= usable_w
    return m.width(line.replace("\t", " ")) <= usable_w


def layout_label(text, style, measurer):
    """Lay out one label (max MAX_LINES lines, each with its own size/bold).
    Returns runs as (x_inches, y_inches, string, pt, bold) relative to the
    label's top-left corner."""
    lines = text.split("\n")[:MAX_LINES]
    family = style.get("font") or None  # None = the measurer's default
    style_lines = style.get("lines") or DEFAULT_STYLE["lines"]
    margin = float(style.get("margin", 0.10))
    gutter = float(style.get("gutter", 0.0))
    usable_w = max(0.2, LABEL_W - 2 * margin)
    usable_h = LABEL_H - 2 * V_PAD - gutter * (len(lines) - 1)

    specs = []  # (size, bold) per text line
    for i in range(len(lines)):
        src = style_lines[min(i, len(style_lines) - 1)]
        specs.append((int(src.get("size", 0) or 0), bool(src.get("bold"))))

    # Fixed-size lines claim their height first; auto lines share the rest.
    pts = [size if size > 0 else None for size, _ in specs]
    fixed_h = sum(measurer(pts[i], specs[i][1], family).line_height
                  for i in range(len(lines)) if pts[i])
    autos = [i for i in range(len(lines)) if pts[i] is None]
    share = ((usable_h - fixed_h) / len(autos)) if autos else 0.0
    for i in autos:
        bold = specs[i][1]
        pts[i] = AUTO_MIN_PT
        for cand in range(AUTO_MAX_PT, AUTO_MIN_PT - 1, -1):
            m = measurer(cand, bold, family)
            if m.line_height <= share and _line_fits(lines[i], style, m,
                                                     usable_w):
                pts[i] = cand
                break

    metrics = [measurer(pts[i], specs[i][1], family)
               for i in range(len(lines))]
    total_h = sum(m.line_height for m in metrics) + gutter * (len(lines) - 1)
    va = style.get("v_align", "middle")
    if va == "top":
        y = V_PAD
    elif va == "bottom":
        y = LABEL_H - V_PAD - total_h
    else:
        y = (LABEL_H - total_h) / 2.0

    runs = []
    ha = style.get("h_align", "center")
    for i, line in enumerate(lines):
        m = metrics[i]
        pt, bold = pts[i], specs[i][1]
        if ha == "tab" and "\t" in line:
            left, right = line.split("\t", 1)
            right = right.replace("\t", " ")
            if left:
                runs.append((margin, y, left, pt, bold))
            if right:
                runs.append((LABEL_W - margin - m.width(right), y, right,
                             pt, bold))
        elif line:
            clean = line.replace("\t", " ")
            w = m.width(clean)
            if ha in ("left", "tab"):
                x = margin
            elif ha == "right":
                x = LABEL_W - margin - w
            else:
                x = (LABEL_W - w) / 2.0
            runs.append((x, y, clean, pt, bold))
        y += m.line_height + gutter
    return runs


class TkMeasurer:
    """Measures text in inches using Tk fonts rendered at PREVIEW_SCALE px/inch."""

    def __init__(self, family):
        self.family = family
        self._cache = {}

    def __call__(self, pt, bold=False, family=None):
        fam = family or self.family
        key = (fam, pt, bold)
        if key not in self._cache:
            px = max(1, round(pt * PREVIEW_SCALE / 72.0))
            self._cache[key] = _TkFontMetrics(tkfont.Font(
                family=fam, size=-px,
                weight="bold" if bold else "normal"))
        return self._cache[key]


class _TkFontMetrics:
    def __init__(self, font):
        self.font = font
        self.line_height = font.metrics("linespace") / PREVIEW_SCALE

    def width(self, s):
        return self.font.measure(s) / PREVIEW_SCALE


class GdiMeasurer:
    """Measures text in inches on a printer DC."""

    def __init__(self, dc, family, dpi_x, dpi_y):
        self.dc = dc
        self.family = family
        self.dpi_x = dpi_x
        self.dpi_y = dpi_y
        self._cache = {}

    def __call__(self, pt, bold=False, family=None):
        fam = family or self.family
        key = (fam, pt, bold)
        if key not in self._cache:
            font = win32ui.CreateFont({
                "name": fam,
                "height": -round(pt * self.dpi_y / 72.0),
                "weight": 700 if bold else 400,
            })
            self._cache[key] = _GdiFontMetrics(self.dc, font,
                                               self.dpi_x, self.dpi_y)
        return self._cache[key]


class _GdiFontMetrics:
    def __init__(self, dc, font, dpi_x, dpi_y):
        self.dc = dc
        self.font = font
        self.dpi_x = dpi_x
        old = dc.SelectObject(font)
        self.line_height = dc.GetTextExtent("Ag")[1] / dpi_y
        dc.SelectObject(old)

    def width(self, s):
        old = self.dc.SelectObject(self.font)
        w = self.dc.GetTextExtent(s)[0] / self.dpi_x
        self.dc.SelectObject(old)
        return w


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

def default_config():
    return {
        "formats": copy.deepcopy(DEFAULT_FORMATS),
        "default_format": None,
        "printer": None,
        "calibration": {"x": 0.0, "y": 0.0},
        "autosave": None,
        "label_sets": {},
    }


def load_config():
    cfg = default_config()
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        for key in cfg:
            if key in data:
                cfg[key] = data[key]
    except (OSError, ValueError):
        pass
    for key in FORMAT_KEYS:  # heal partial/missing format slots
        if key not in cfg["formats"] or "style" not in cfg["formats"][key]:
            cfg["formats"][key] = copy.deepcopy(DEFAULT_FORMATS[key])
        cfg["formats"][key]["style"] = normalize_style(
            cfg["formats"][key]["style"])
    if cfg["default_format"] not in (None,) + FORMAT_KEYS:
        cfg["default_format"] = None
    return cfg


def save_config(cfg):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except OSError as e:
        messagebox.showerror(APP_NAME, "Could not save settings:\n%s" % e)


# ---------------------------------------------------------------------------
# Printing (raw GDI via pywin32)
# ---------------------------------------------------------------------------

def compute_calibration(cur_x, cur_y, meas_left, meas_right=None,
                        meas_bottom=None):
    """Derive corrected print offsets from ruler measurements of the
    letter-paper alignment test printed with the current offsets.

    meas_left:   paper left edge -> box left side (target ALIGN_TARGET_LR)
    meas_right:  paper right edge -> box right side (target ALIGN_TARGET_LR);
                 optional, averaged with the left measurement
    meas_bottom: paper top edge -> box bottom line (target
                 ALIGN_TARGET_BOTTOM); optional, Y unchanged if absent
    """
    dxs = [meas_left - ALIGN_TARGET_LR]
    if meas_right is not None:
        dxs.append(ALIGN_TARGET_LR - meas_right)
    dx = sum(dxs) / len(dxs)
    dy = (meas_bottom - ALIGN_TARGET_BOTTOM) if meas_bottom is not None else 0.0
    return cur_x - dx, cur_y - dy


def list_printers():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        return [p["pPrinterName"] for p in win32print.EnumPrinters(flags, None, 2)]
    except Exception:
        return [p[2] for p in win32print.EnumPrinters(flags, None, 1)]


def _find_4x6_paper(printer_name):
    """Return the driver's paper code for 4x6 in media, or None."""
    try:
        h = win32print.OpenPrinter(printer_name)
        try:
            port = win32print.GetPrinter(h, 2).get("pPortName") or ""
        finally:
            win32print.ClosePrinter(h)
        codes = win32print.DeviceCapabilities(printer_name, port, DC_PAPERS)
        sizes = win32print.DeviceCapabilities(printer_name, port, DC_PAPERSIZE)
        for code, size in zip(codes, sizes):
            try:
                cx, cy = int(size[0]), int(size[1])
            except (TypeError, ValueError, IndexError, KeyError):
                continue
            # 4" x 6" = 1016 x 1524 in 0.1 mm units; allow 3 mm tolerance
            if abs(cx - 1016) <= 30 and abs(cy - 1524) <= 30:
                return code
    except Exception:
        pass
    return None


def _create_printer_dc(printer_name, letter=False):
    """Printer DC forced to portrait media — 4x6 for labels, or letter for
    the alignment test — never via the driver dialog."""
    devmode = None
    try:
        h = win32print.OpenPrinter(printer_name)
        try:
            devmode = win32print.GetPrinter(h, 2)["pDevMode"]
        finally:
            win32print.ClosePrinter(h)
    except Exception:
        devmode = None
    if devmode is not None:
        fields = DM_ORIENTATION
        devmode.Orientation = 1  # portrait
        if letter:
            devmode.PaperSize = DMPAPER_LETTER
            fields |= DM_PAPERSIZE
        else:
            code = _find_4x6_paper(printer_name)
            if code is not None:
                devmode.PaperSize = code
                fields |= DM_PAPERSIZE
            else:
                devmode.PaperSize = DMPAPER_USER
                devmode.PaperWidth = 1016   # 0.1 mm units
                devmode.PaperLength = 1524
                fields |= DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH
        devmode.Fields = devmode.Fields | fields
    hdc = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
    return win32ui.CreateDCFromHandle(hdc)


def print_sheet(printer_name, rows, calibration):
    """Print one 4x6 sheet. rows = [(text, style)] for the 7 positions;
    blank rows print nothing (partial-sheet reuse)."""
    dc = _create_printer_dc(printer_name)
    try:
        dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
        off_x = dc.GetDeviceCaps(PHYSICALOFFSETX)
        off_y = dc.GetDeviceCaps(PHYSICALOFFSETY)
        cal_x = float(calibration.get("x", 0.0))
        cal_y = float(calibration.get("y", 0.0))

        def px(x_in):
            return round((x_in + cal_x) * dpi_x) - off_x

        def py(y_in):
            return round((y_in + cal_y) * dpi_y) - off_y

        measurer = GdiMeasurer(dc, FONT_NAME, dpi_x, dpi_y)

        dc.StartDoc(APP_NAME + " - folder labels")
        try:
            dc.StartPage()
            dc.SetBkMode(win32con.TRANSPARENT)
            for i in range(NUM_LABELS):
                top = MARGIN_TOP + i * LABEL_H
                text, style = rows[i]
                if is_blank(text):
                    continue
                fam = style.get("font") or None
                for rx, ry, s, pt, bold in layout_label(text, style, measurer):
                    dc.SelectObject(measurer(pt, bold, fam).font)
                    dc.TextOut(px(MARGIN_LEFT + rx), py(top + ry), s)
            dc.EndPage()
            dc.EndDoc()
        except Exception:
            try:
                dc.AbortDoc()
            except Exception:
                pass
            raise
    finally:
        dc.DeleteDC()


def print_alignment_sheet(printer_name, calibration):
    """Print the calibration test on plain letter paper, all on one page:
    the 4x6 sheet outline centered horizontally (left/right 2.25"), its top
    dropped to 0.25" (the paper's very top is unprintable), the 7 label
    boxes inside it, and the measuring worksheet below."""
    dc = _create_printer_dc(printer_name, letter=True)
    try:
        dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)
        off_x = dc.GetDeviceCaps(PHYSICALOFFSETX)
        off_y = dc.GetDeviceCaps(PHYSICALOFFSETY)
        cal_x = float(calibration.get("x", 0.0))
        cal_y = float(calibration.get("y", 0.0))

        def px(x_in):
            return round((x_in + cal_x) * dpi_x) - off_x

        def py(y_in):
            return round((y_in + cal_y) * dpi_y) - off_y

        def px0(x_in):  # no calibration offset
            return round(x_in * dpi_x) - off_x

        def py0(y_in):
            return round(y_in * dpi_y) - off_y

        def rect(l, t, r, b, fx, fy):
            dc.MoveTo((fx(l), fy(t)))
            dc.LineTo((fx(r), fy(t)))
            dc.LineTo((fx(r), fy(b)))
            dc.LineTo((fx(l), fy(b)))
            dc.LineTo((fx(l), fy(t)))

        DASH_LEN = 0.25   # dashes drawn manually: 1/4" on, 1/8" off
        DASH_GAP = 0.125

        def dashed_rect(l, t, r, b, fx, fy):
            for y_line in (t, b):
                x = l
                while x < r:
                    seg = min(x + DASH_LEN, r)
                    dc.MoveTo((fx(x), fy(y_line)))
                    dc.LineTo((fx(seg), fy(y_line)))
                    x = seg + DASH_GAP
            for x_line in (l, r):
                y_pos = t
                while y_pos < b:
                    seg = min(y_pos + DASH_LEN, b)
                    dc.MoveTo((fx(x_line), fy(y_pos)))
                    dc.LineTo((fx(x_line), fy(seg)))
                    y_pos = seg + DASH_GAP

        measurer = GdiMeasurer(dc, FONT_NAME, dpi_x, dpi_y)
        dc.StartDoc(APP_NAME + " - alignment test")
        try:
            dc.StartPage()
            dc.SetBkMode(win32con.TRANSPARENT)
            # dashed 4x6 outline with NO offset (raw printer position),
            # then the solid outline WITH the current offset applied —
            # the gap between them shows what the calibration is doing
            dashed_rect(ALIGN_BOX_LEFT, ALIGN_BOX_TOP,
                        ALIGN_BOX_LEFT + SHEET_W, ALIGN_TARGET_BOTTOM,
                        px0, py0)
            rect(ALIGN_BOX_LEFT, ALIGN_BOX_TOP,
                 ALIGN_BOX_LEFT + SHEET_W, ALIGN_TARGET_BOTTOM, px, py)
            # caption inside the outline's top margin
            m8 = measurer(8)
            dc.SelectObject(m8.font)
            dc.TextOut(px(ALIGN_BOX_LEFT + 0.08),
                       py(ALIGN_BOX_TOP + 0.16),
                       "%s alignment test  cal x=%+.3f y=%+.3f" %
                       (APP_NAME, cal_x, cal_y))
            # the 7 label boxes, positioned as on the real sheet
            m10 = measurer(10)
            for i in range(NUM_LABELS):
                top = ALIGN_BOX_TOP + MARGIN_TOP + i * LABEL_H
                rect(ALIGN_BOX_LEFT + MARGIN_LEFT, top,
                     ALIGN_BOX_LEFT + MARGIN_LEFT + LABEL_W, top + LABEL_H,
                     px, py)
                s = str(i + 1)
                dc.SelectObject(m10.font)
                dc.TextOut(px(ALIGN_BOX_LEFT + MARGIN_LEFT +
                              (LABEL_W - m10.width(s)) / 2),
                           py(top + (LABEL_H - m10.line_height) / 2), s)
            # legend inside the box's bottom margin, below label 7
            dc.SelectObject(m8.font)
            legend = "Measure the SOLID box.  Dashed box = no offset."
            dc.TextOut(px(ALIGN_BOX_LEFT +
                          (SHEET_W - m8.width(legend)) / 2),
                       py(ALIGN_BOX_TOP + MARGIN_TOP +
                          NUM_LABELS * LABEL_H + 0.12),
                       legend)

            # instructions below the outline, centered on the page
            y = ALIGN_TARGET_BOTTOM + 0.35
            for line in (
                    "Go to the Print menu, click Calibration, enter the "
                    "measurements, and click",
                    "Calculate to automatically calculate the offsets. "
                    "Save, then reprint to verify."):
                dc.TextOut(px((ALIGN_PAGE_W - m8.width(line)) / 2),
                           py(y), line)
                y += m8.line_height * 1.3
            dc.EndPage()
            dc.EndDoc()
        except Exception:
            try:
                dc.AbortDoc()
            except Exception:
                pass
            raise
    finally:
        dc.DeleteDC()


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def style_button(btn, accent=False):
    """Flat dark button; accent=True gives the amber primary style."""
    if accent:
        btn.configure(bg=CLR_ACCENT, fg=CLR_ON_ACCENT,
                      activebackground=CLR_ACCENT_ACTIVE,
                      activeforeground=CLR_ON_ACCENT,
                      font=("Segoe UI", 9, "bold"))
    else:
        btn.configure(bg=CLR_BTN, fg=CLR_FG,
                      activebackground=CLR_BTN_HOVER,
                      activeforeground="#ffffff", font=("Segoe UI", 9))
    btn.configure(relief="flat", bd=0, padx=10, pady=3, cursor="hand2")
    btn._themed = True


def apply_theme(widget):
    """Recursively apply the dark palette to a dialog's tk widgets
    (widgets tagged _themed keep their explicit styling)."""
    for w in widget.winfo_children():
        if not getattr(w, "_themed", False):
            cls = w.winfo_class()
            try:
                if cls == "Frame":
                    w.configure(bg=CLR_BG)
                elif cls == "Labelframe":
                    w.configure(bg=CLR_BG, fg=CLR_DIM, bd=1, relief="groove",
                                font=("Segoe UI", 9))
                elif cls == "Label":
                    w.configure(bg=CLR_BG, fg=CLR_FG, font=("Segoe UI", 9))
                elif cls == "Button":
                    style_button(w)
                elif cls in ("Radiobutton", "Checkbutton"):
                    w.configure(bg=CLR_BG, fg=CLR_FG, selectcolor=CLR_BTN,
                                activebackground=CLR_BG,
                                activeforeground="#ffffff",
                                font=("Segoe UI", 9))
                elif cls == "Entry":
                    w.configure(bg=CLR_FIELD, fg=CLR_FG,
                                insertbackground=CLR_FG, relief="flat")
                elif cls == "Spinbox":
                    w.configure(bg=CLR_FIELD, fg=CLR_FG,
                                insertbackground=CLR_FG, relief="flat",
                                buttonbackground=CLR_BTN)
                elif cls == "Listbox":
                    w.configure(bg=CLR_FIELD, fg=CLR_FG, relief="flat",
                                selectbackground=CLR_ACCENT,
                                selectforeground=CLR_ON_ACCENT)
            except tk.TclError:
                pass
        apply_theme(w)


def enable_dark_titlebar(window):
    """Ask DWM to draw this window's native title bar dark (Win10 1903+)."""
    try:
        import ctypes
        window.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass


def theme_toplevel(top):
    top.configure(bg=CLR_BG)
    apply_theme(top)
    enable_dark_titlebar(top)


# ---------------------------------------------------------------------------
# Tutorial video player (MCI via winmm — stdlib ctypes, no extra deps)
# ---------------------------------------------------------------------------

def _mci(cmd):
    """Send one MCI command string to winmm, returning the result string."""
    import ctypes
    buf = ctypes.create_unicode_buffer(256)
    err = ctypes.windll.winmm.mciSendStringW(cmd, buf, 256, None)
    if err:
        ebuf = ctypes.create_unicode_buffer(256)
        ctypes.windll.winmm.mciGetErrorStringW(err, ebuf, 256)
        raise RuntimeError("MCI: %s (%s)" % (ebuf.value, cmd))
    return buf.value


def tutorial_video_path():
    """The tutorial video ships as a separate file next to the EXE (it is
    not bundled into the onefile EXE); from source it lives in tutorial/."""
    if getattr(sys, "frozen", False):
        cands = (os.path.join(os.path.dirname(sys.executable), TUTORIAL_FILE),)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
        cands = (os.path.join(base, "tutorial", TUTORIAL_FILE),
                 os.path.join(base, TUTORIAL_FILE))
    for path in cands:
        if os.path.exists(path):
            return path
    return None


class TutorialDialog(tk.Toplevel):
    """Non-modal player for the built-in tutorial video: MCI renders into
    a child window over self.video, so the app stays usable while it plays.
    The video is WMV because stock DirectShow (which MCI's MPEGVideo device
    uses) cannot open MP4/H.264."""

    ALIAS = "tri_tutorial"

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Tutorial")
        self.resizable(False, False)
        self._mci_open = False
        self._poll_job = None
        self.muted = False

        self.video = tk.Frame(self, bg="#000000", width=400, height=300)
        self.video.pack()
        self.video.pack_propagate(False)
        self.video._themed = True

        bar = tk.Frame(self, padx=10, pady=8)
        bar.pack(fill="x")
        self.play_btn = tk.Button(bar, text="Pause", width=8,
                                  command=self.toggle_play)
        style_button(self.play_btn, accent=True)
        self.play_btn.pack(side="left")
        self.mute_btn = tk.Button(bar, text="Mute", width=8,
                                  command=self.toggle_mute)
        style_button(self.mute_btn)
        self.mute_btn.pack(side="left", padx=(6, 0))
        self.volume = tk.Scale(
            bar, from_=0, to=100, orient="horizontal", showvalue=False,
            length=150, command=lambda _v: self._apply_volume(),
            bg=CLR_ACCENT, troughcolor=CLR_BTN, highlightthickness=0, bd=0,
            activebackground=CLR_ACCENT_ACTIVE, sliderrelief="flat")
        self.volume.set(80)
        self.volume._themed = True
        self.volume.pack(side="right")
        tk.Label(bar, text="Volume").pack(side="right", padx=(0, 6))

        theme_toplevel(self)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", lambda e: self.close())
        self.transient(app.root)
        self.update()  # realize the video frame so it has an HWND

        path = tutorial_video_path()
        try:
            if path is None:
                raise RuntimeError("tutorial video not found")
            _mci('open "%s" type mpegvideo alias %s parent %d style child'
                 % (path, self.ALIAS, self.video.winfo_id()))
            self._mci_open = True
            src = _mci("where %s source" % self.ALIAS).split()
            vw, vh = int(src[2]), int(src[3])
            scale = min(1.0, (self.winfo_screenwidth() - 80.0) / vw,
                        (self.winfo_screenheight() - 220.0) / vh)
            w, h = int(vw * scale), int(vh * scale)
            self.video.configure(width=w, height=h)
            _mci("put %s window at 0 0 %d %d" % (self.ALIAS, w, h))
            self._apply_volume()
            _mci("play %s" % self.ALIAS)
        except Exception:
            self.close()
            self._offer_external(path)
            return

        self.update_idletasks()
        rx, ry = app.root.winfo_rootx(), app.root.winfo_rooty()
        rw = app.root.winfo_width()
        self.geometry("+%d+%d"
                      % (max(0, rx + (rw - self.winfo_reqwidth()) // 2),
                         max(0, ry - 20)))
        self._poll_job = self.after(300, self._poll)

    def _offer_external(self, path):
        if path is None:
            messagebox.showerror(
                "Tutorial", "The tutorial video was not found.",
                parent=self.app.root)
        elif messagebox.askyesno(
                "Tutorial", "The built-in player could not start.\n"
                "Open the tutorial in your default video player instead?",
                parent=self.app.root):
            os.startfile(path)

    def toggle_play(self):
        try:
            mode = _mci("status %s mode" % self.ALIAS)
            if mode == "playing":
                _mci("pause %s" % self.ALIAS)
                self.play_btn.configure(text="Play")
            else:
                if mode == "stopped":  # replay after reaching the end
                    _mci("seek %s to start" % self.ALIAS)
                _mci("play %s" % self.ALIAS)
                self.play_btn.configure(text="Pause")
        except RuntimeError:
            pass

    def toggle_mute(self):
        try:
            _mci("setaudio %s %s"
                 % (self.ALIAS, "on" if self.muted else "off"))
            self.muted = not self.muted
            if not self.muted:
                self._apply_volume()
            self.mute_btn.configure(text="Unmute" if self.muted else "Mute")
        except RuntimeError:
            pass

    def _apply_volume(self):
        if self._mci_open and not self.muted:
            try:
                _mci("setaudio %s volume to %d"
                     % (self.ALIAS, self.volume.get() * 10))
            except RuntimeError:
                pass

    def _poll(self):
        try:
            if _mci("status %s mode" % self.ALIAS) == "stopped":
                self.play_btn.configure(text="Play")
        except RuntimeError:
            pass
        self._poll_job = self.after(300, self._poll)

    def close(self):
        if self._poll_job is not None:
            self.after_cancel(self._poll_job)
            self._poll_job = None
        if self._mci_open:
            try:
                _mci("close %s" % self.ALIAS)
            except RuntimeError:
                pass
            self._mci_open = False
        self.destroy()


class StyleEditor(tk.LabelFrame):
    """Reusable style controls, used by the per-label dialog and Edit Formats."""

    def __init__(self, parent, style, title="Style"):
        super().__init__(parent, text=title, padx=8, pady=6)
        style = normalize_style(style)
        self.h_align = tk.StringVar(value=style.get("h_align", "center"))
        self.v_align = tk.StringVar(value=style.get("v_align", "middle"))
        self.margin = tk.StringVar(value="%.2f" % style.get("margin", 0.10))

        self.font_var = tk.StringVar(value=style.get("font") or FONT_NAME)
        row = tk.Frame(self)
        row.pack(anchor="w", pady=(0, 4))
        tk.Label(row, text="Font:").pack(side="left")
        families = sorted({f for f in tkfont.families()
                           if not f.startswith("@")})
        ttk.Combobox(row, textvariable=self.font_var, values=families,
                     width=28, state="readonly").pack(side="left",
                                                      padx=(4, 0))

        self.line_vars = []
        for i in range(MAX_LINES):
            ls = style["lines"][i]
            mode = tk.StringVar(value="fixed" if ls["size"] else "auto")
            size = tk.IntVar(value=ls["size"] or 12)
            bold = tk.BooleanVar(value=ls["bold"])
            self.line_vars.append((mode, size, bold))
            row = tk.Frame(self)
            row.pack(anchor="w")
            tk.Label(row, text="Line %d size:" % (i + 1)).pack(side="left")
            tk.Radiobutton(row, text="Auto", variable=mode,
                           value="auto").pack(side="left")
            tk.Radiobutton(row, text="Points:", variable=mode,
                           value="fixed").pack(side="left")
            tk.Spinbox(row, from_=AUTO_MIN_PT, to=72, width=4,
                       textvariable=size).pack(side="left")
            tk.Checkbutton(row, text="Bold", variable=bold).pack(
                side="left", padx=(8, 0))
            # Editing the size (typing or arrows) implies "Points" mode —
            # otherwise a typed size is silently ignored while on Auto.
            # The trace must attach only after the Spinbox exists: giving a
            # Spinbox its textvariable writes to it once at creation, which
            # must not flip the mode by itself.
            size.trace_add("write", lambda *_a, m=mode: m.set("fixed"))

        row = tk.Frame(self)
        row.pack(anchor="w", pady=(4, 0))
        tk.Label(row, text="Horizontal:").pack(side="left")
        for text, val in (("Left", "left"), ("Center", "center"),
                          ("Right", "right"), ("Tab L/R", "tab")):
            tk.Radiobutton(row, text=text, variable=self.h_align,
                           value=val).pack(side="left")

        row = tk.Frame(self)
        row.pack(anchor="w", pady=(4, 0))
        tk.Label(row, text="Vertical:").pack(side="left")
        for text, val in (("Top", "top"), ("Middle", "middle"),
                          ("Bottom", "bottom")):
            tk.Radiobutton(row, text=text, variable=self.v_align,
                           value=val).pack(side="left")

        self.gutter = tk.StringVar(value="%.2f" % style.get("gutter", 0.0))
        row = tk.Frame(self)
        row.pack(anchor="w", pady=(4, 0))
        tk.Label(row, text="Margin (in):").pack(side="left")
        tk.Entry(row, width=6, textvariable=self.margin).pack(side="left")
        tk.Label(row, text="Line gutter (in):").pack(side="left", padx=(10, 0))
        tk.Entry(row, width=6, textvariable=self.gutter).pack(side="left")

    def set_style(self, style):
        """Load a style into the controls (e.g. when picking a format)."""
        style = normalize_style(style)
        self.font_var.set(style.get("font") or FONT_NAME)
        self.h_align.set(style.get("h_align", "center"))
        self.v_align.set(style.get("v_align", "middle"))
        self.margin.set("%.2f" % style.get("margin", 0.10))
        self.gutter.set("%.2f" % style.get("gutter", 0.0))
        for (mode, size, bold), ls in zip(self.line_vars, style["lines"]):
            # size first: its write-trace forces "fixed", so mode must be
            # set afterwards for auto lines to stay auto.
            size.set(ls["size"] or 12)
            mode.set("fixed" if ls["size"] else "auto")
            bold.set(bool(ls["bold"]))

    def get_style(self):
        try:
            margin = max(0.0, min(1.0, float(self.margin.get())))
        except ValueError:
            margin = 0.10
        try:
            gutter = max(0.0, min(0.5, float(self.gutter.get())))
        except ValueError:
            gutter = 0.0
        lines = []
        for mode, size, bold in self.line_vars:
            val = 0
            if mode.get() == "fixed":
                try:
                    val = max(AUTO_MIN_PT, min(72, int(size.get())))
                except tk.TclError:
                    val = 12
            lines.append({"size": val, "bold": bool(bold.get())})
        return {
            "font": self.font_var.get() or FONT_NAME,
            "h_align": self.h_align.get(),
            "v_align": self.v_align.get(),
            "margin": margin,
            "gutter": gutter,
            "lines": lines,
        }


class LabelRow:
    def __init__(self, app, parent, index):
        self.app = app
        self.index = index
        self.preset = None
        self.style = app.default_style()
        self._placeholder = False

        f = tk.Frame(parent, bg=CLR_ROW)
        f.pack(fill="x", pady=ROW_GAP)
        self.frame = f
        self.accent = tk.Frame(f, width=3, bg=CLR_ROW)
        self.accent.pack(side="left", fill="y")
        self.num = tk.Label(f, text=str(index + 1), width=2, bg=CLR_ROW,
                            fg=CLR_DIM, font=("Segoe UI", 9))
        self.num.pack(side="left", padx=(2, 4))
        self.text = tk.Text(f, height=2, width=30, wrap="none", undo=True,
                            bg=CLR_ROW, fg=CLR_FG,
                            insertbackground=CLR_ACCENT,
                            selectbackground=CLR_ACCENT,
                            selectforeground=CLR_ON_ACCENT,
                            relief="flat", bd=0, highlightthickness=0,
                            font=(app.ui_family, 10))
        self.text.pack(side="left", pady=5)
        self.text.tag_configure("ph", foreground=CLR_DIM)
        self.text.bind("<Return>", self._on_enter)
        self.text.bind("<Shift-Return>", self._on_shift_enter)
        self.text.bind("<Tab>", self._on_tab)
        self.text.bind("<Shift-Tab>", self._on_shift_tab)
        self.text.bind("<Control-Tab>", self._on_ctrl_tab)
        self.text.bind("<KeyRelease>", self._on_key_release)
        self.text.bind("<FocusIn>", self._on_focus_in)
        self.text.bind("<FocusOut>", self._on_focus_out)

        # Drag handle (⋮) — grab and drag vertically to reorder labels.
        self.handle = tk.Label(f, text="⋮", bg=CLR_ROW, fg=CLR_DIM,
                               font=("Segoe UI", 11), cursor="fleur", padx=4)
        self.handle.pack(side="right", fill="y", padx=(0, 4))
        self.handle.bind("<ButtonPress-1>",
                         lambda e: app.begin_row_drag(self.index))
        self.handle.bind("<B1-Motion>",
                         lambda e: app.update_row_drag(e.y_root))
        self.handle.bind("<ButtonRelease-1>",
                         lambda e: app.end_row_drag(e.y_root))

        self.btns = tk.Frame(f, bg=CLR_ROW)
        self.btns.pack(side="right", padx=6)
        small = ("Segoe UI", 8)

        def rbtn(text, cmd, padx=(0, 0), width=None):
            b = tk.Button(self.btns, text=text, command=cmd, bg=CLR_BTN,
                          fg=CLR_FG, activebackground=CLR_BTN_HOVER,
                          activeforeground="#ffffff", relief="flat", bd=0,
                          font=small, padx=6, cursor="hand2")
            if width:
                b.configure(width=width)
            b._themed = True
            b.pack(side="left", padx=padx)
            return b

        rbtn("Style…", self.edit_style, padx=(0, 6))
        self.preset_buttons = {}
        for key in FORMAT_KEYS:
            self.preset_buttons[key] = rbtn(
                key, lambda k=key: self.set_preset(k), padx=(1, 1), width=2)
        rbtn("▲", self.dup_up, padx=(6, 1), width=2).configure(fg=CLR_DIM)
        rbtn("▼", self.dup_down, padx=(1, 1), width=2).configure(fg=CLR_DIM)
        rbtn("✕", self.clear, padx=(1, 0), width=2).configure(fg=CLR_DIM)
        self.refresh_preset_buttons()

        if app.config.get("default_format"):
            self.set_preset(app.config["default_format"], silent=True)
        self._show_placeholder()

        # Freeze the row at the preview's label pitch: keep the natural width
        # the children just requested, but force the height to ROW_PITCH so
        # row N stays level with label N on the preview sheet.
        f.update_idletasks()
        f.configure(width=f.winfo_reqwidth(),
                    height=ROW_PITCH - 2 * ROW_GAP)
        f.pack_propagate(False)

    # -- data ------------------------------------------------------------
    def get_text(self):
        if self._placeholder:
            return ""
        return self.text.get("1.0", "end-1c")

    def set_text(self, value):
        value = "\n".join(value.split("\n")[:MAX_LINES])
        self._placeholder = False
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value)
        if self.app.root.focus_get() is not self.text:
            self._show_placeholder()

    def _show_placeholder(self):
        if not self._placeholder and self.text.get("1.0", "end-1c") == "":
            self._placeholder = True
            self.text.insert("1.0", "Empty", "ph")

    def _on_focus_in(self, _event):
        if self._placeholder:
            self._placeholder = False
            self.text.delete("1.0", "end")
        self.app.set_active_row(self.index)

    def _on_focus_out(self, _event):
        self._show_placeholder()

    def set_visual_active(self, active):
        row_bg = CLR_ROW_ACTIVE if active else CLR_ROW
        self.frame.configure(bg=row_bg)
        self.accent.configure(bg=CLR_ACCENT if active else row_bg)
        self.num.configure(bg=row_bg, fg=CLR_FG if active else CLR_DIM)
        self.text.configure(bg=row_bg)
        self.btns.configure(bg=row_bg)
        self.handle.configure(bg=row_bg)

    def dump(self):
        return {"text": self.get_text(),
                "style": copy.deepcopy(self.style),
                "preset": self.preset}

    def load(self, d):
        self.set_text(d.get("text", ""))
        style = d.get("style")
        self.style = normalize_style(style) if style else self.app.default_style()
        self.preset = d.get("preset") if d.get("preset") in FORMAT_KEYS else None
        self.refresh_preset_buttons()

    # -- presets & style -------------------------------------------------
    def set_preset(self, key, silent=False):
        self.preset = key
        self.style = copy.deepcopy(self.app.config["formats"][key]["style"])
        self.refresh_preset_buttons()
        if not silent:
            self.app.schedule_preview()

    def detach_preset(self):
        self.preset = None
        self.refresh_preset_buttons()

    def refresh_preset_buttons(self):
        for key, btn in self.preset_buttons.items():
            if key == self.preset:
                btn.configure(bg=CLR_ACCENT, fg=CLR_ON_ACCENT,
                              activebackground=CLR_ACCENT_ACTIVE,
                              activeforeground=CLR_ON_ACCENT)
            else:
                btn.configure(bg=CLR_BTN, fg=CLR_DIM,
                              activebackground=CLR_BTN_HOVER,
                              activeforeground=CLR_FG)

    def sync_from_preset(self):
        if self.preset:
            self.style = copy.deepcopy(
                self.app.config["formats"][self.preset]["style"])

    def edit_style(self):
        dlg = StyleDialog(self.app.root,
                          "Label %d style" % (self.index + 1), self.style,
                          formats=self.app.config["formats"],
                          preset=self.preset)
        if dlg.result is not None:
            if dlg.preset:
                self.set_preset(dlg.preset)  # picked a format → link to it
            else:
                self.style = dlg.result
                self.detach_preset()  # manual tweak unlinks the preset
                self.app.schedule_preview()

    # -- duplicate -------------------------------------------------------
    def copy_from(self, other):
        self.set_text(other.get_text())
        self.style = copy.deepcopy(other.style)
        self.preset = other.preset
        self.refresh_preset_buttons()
        self.app.schedule_preview()

    def dup_up(self):
        if self.index > 0:
            self.app.rows[self.index - 1].copy_from(self)

    def dup_down(self):
        if self.index < NUM_LABELS - 1:
            self.app.rows[self.index + 1].copy_from(self)

    def clear(self):
        """Clear the label's text; its style and format link are kept."""
        self.set_text("")
        self.app.schedule_preview()

    # -- key handling ----------------------------------------------------
    def _on_enter(self, _event):
        self.app.focus_row((self.index + 1) % NUM_LABELS)
        return "break"

    def _on_tab(self, _event):
        self.app.focus_row((self.index + 1) % NUM_LABELS)
        return "break"

    def _on_shift_tab(self, _event):
        self.app.focus_row((self.index - 1) % NUM_LABELS)
        return "break"

    def _on_shift_enter(self, _event):
        # allow the newline only while under the line limit
        if int(self.text.index("end-1c").split(".")[0]) >= MAX_LINES:
            return "break"
        return None  # default class binding inserts the newline

    def _on_key_release(self, _event):
        lines = self.get_text().split("\n")
        if len(lines) > MAX_LINES:  # e.g. pasted multi-line text
            self.set_text("\n".join(lines[:MAX_LINES]))
            self.text.mark_set("insert", "end")
        self.app.schedule_preview()

    def _on_ctrl_tab(self, _event):
        self.text.insert("insert", "\t")
        self.app.schedule_preview()
        return "break"


class StyleDialog(tk.Toplevel):
    """Per-label style dialog. If `formats` is given, a selector row lets the
    user load one of the global formats into the controls; `self.preset` then
    reports the picked format on OK — unless the controls were tweaked after
    picking, which detaches (same rule as the row's 1/2/3 buttons)."""

    def __init__(self, parent, title, style, formats=None, preset=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.preset = None
        self._snapshot = None
        self._formats = formats
        self._fmt_buttons = {}
        if formats:
            row = tk.Frame(self)
            row.pack(anchor="w", padx=10, pady=(10, 0))
            tk.Label(row, text="Format:").pack(side="left")
            for key in FORMAT_KEYS:
                name = formats[key].get("name", "Format %s" % key)
                b = tk.Button(row, text="%s · %s" % (key, name),
                              command=lambda k=key: self._pick_format(k))
                b.pack(side="left", padx=3)
                self._fmt_buttons[key] = b
        self.editor = StyleEditor(self, style, title="Options")
        self.editor.pack(padx=10, pady=10, fill="x")
        btns = tk.Frame(self)
        btns.pack(pady=(0, 10))
        tk.Button(btns, text="OK", width=9, command=self._ok).pack(
            side="left", padx=4)
        tk.Button(btns, text="Cancel", width=9, command=self.destroy).pack(
            side="left", padx=4)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        theme_toplevel(self)
        if formats and preset in FORMAT_KEYS:
            # Label is already linked: pre-highlight its format and snapshot
            # the current controls so an untouched OK keeps the link.
            self.preset = preset
            self._snapshot = self.editor.get_style()
            self._mark_format()
        self.transient(parent)
        self.grab_set()
        self.wait_window()

    def _pick_format(self, key):
        self.preset = key
        self.editor.set_style(self._formats[key]["style"])
        self._snapshot = self.editor.get_style()
        self._mark_format()

    def _mark_format(self):
        for key, b in self._fmt_buttons.items():
            if key == self.preset:
                b.configure(bg=CLR_ACCENT, fg=CLR_ON_ACCENT,
                            activebackground=CLR_ACCENT_ACTIVE,
                            activeforeground=CLR_ON_ACCENT)
            else:
                b.configure(bg=CLR_BTN, fg=CLR_FG,
                            activebackground=CLR_BTN_HOVER,
                            activeforeground="#ffffff")

    def _ok(self):
        self.result = self.editor.get_style()
        if self._snapshot is None or self.result != self._snapshot:
            self.preset = None  # tweaked after picking → not linked
        self.destroy()


class FormatsDialog(tk.Toplevel):
    """Edit the three global formats and the default format for blank labels."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Edit formats")
        self.resizable(False, False)

        nb = ttk.Notebook(self)
        nb.pack(padx=10, pady=10)
        self.name_vars = {}
        self.editors = {}
        for key in FORMAT_KEYS:
            fmt = app.config["formats"][key]
            frame = tk.Frame(nb, padx=8, pady=8)
            nb.add(frame, text="Format %s" % key)
            row = tk.Frame(frame)
            row.pack(anchor="w", pady=(0, 6))
            tk.Label(row, text="Name:").pack(side="left")
            var = tk.StringVar(value=fmt.get("name", "Format %s" % key))
            tk.Entry(row, width=28, textvariable=var).pack(side="left")
            self.name_vars[key] = var
            editor = StyleEditor(frame, fmt["style"])
            editor.pack(fill="x")
            self.editors[key] = editor

        row = tk.Frame(self)
        row.pack(anchor="w", padx=10)
        tk.Label(row, text="Default format for blank labels:").pack(side="left")
        self.default_var = tk.StringVar(
            value=app.config.get("default_format") or "none")
        tk.Radiobutton(row, text="None", variable=self.default_var,
                       value="none").pack(side="left")
        for key in FORMAT_KEYS:
            tk.Radiobutton(row, text=key, variable=self.default_var,
                           value=key).pack(side="left")

        btns = tk.Frame(self)
        btns.pack(pady=10)
        tk.Button(btns, text="Save", width=9, command=self._save).pack(
            side="left", padx=4)
        tk.Button(btns, text="Cancel", width=9, command=self.destroy).pack(
            side="left", padx=4)
        self.bind("<Escape>", lambda e: self.destroy())
        theme_toplevel(self)
        self.transient(app.root)
        self.grab_set()
        self.wait_window()

    def _save(self):
        for key in FORMAT_KEYS:
            self.app.config["formats"][key] = {
                "name": self.name_vars[key].get().strip() or "Format %s" % key,
                "style": self.editors[key].get_style(),
            }
        val = self.default_var.get()
        self.app.config["default_format"] = val if val in FORMAT_KEYS else None
        save_config(self.app.config)
        self.app.on_formats_changed()
        self.destroy()


class PrinterDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Print")
        self.resizable(False, False)
        self.choice = None

        tk.Label(self, text="Select printer:").pack(
            anchor="w", padx=10, pady=(10, 2))
        self.listbox = tk.Listbox(self, width=48, height=8)
        self.listbox.pack(padx=10)
        printers = list_printers()
        preselect = app.config.get("printer")
        if not preselect:
            try:
                preselect = win32print.GetDefaultPrinter()
            except Exception:
                preselect = None
        for i, name in enumerate(printers):
            self.listbox.insert("end", name)
            if name == preselect:
                self.listbox.selection_set(i)
                self.listbox.see(i)
        if printers and not self.listbox.curselection():
            self.listbox.selection_set(0)

        self.remember = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text="Remember as Quick Print printer",
                       variable=self.remember).pack(anchor="w", padx=10, pady=4)

        btns = tk.Frame(self)
        btns.pack(pady=(2, 10))
        tk.Button(btns, text="Print", width=9, command=self._ok).pack(
            side="left", padx=4)
        tk.Button(btns, text="Cancel", width=9, command=self.destroy).pack(
            side="left", padx=4)
        self.listbox.bind("<Double-Button-1>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        theme_toplevel(self)
        self.transient(app.root)
        self.grab_set()
        self.wait_window()

    def _ok(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        self.choice = self.listbox.get(sel[0])
        if self.remember.get():
            self.app.config["printer"] = self.choice
            save_config(self.app.config)
        self.destroy()


class CalibrationDialog(tk.Toplevel):
    """Set print offsets directly, or calculate them from ruler measurements
    taken off a printed alignment test."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Calibration")
        self.resizable(False, False)
        cal = app.config.get("calibration", {"x": 0.0, "y": 0.0})

        box = tk.LabelFrame(self, text="Offsets (inches)", padx=8, pady=6)
        box.pack(fill="x", padx=10, pady=(10, 4))
        row = tk.Frame(box)
        row.pack(anchor="w")
        tk.Label(row, text="X (+ right):").pack(side="left")
        self.x_var = tk.StringVar(value="%+.3f" % cal.get("x", 0.0))
        tk.Entry(row, width=8, textvariable=self.x_var).pack(
            side="left", padx=(2, 14))
        tk.Label(row, text="Y (+ down):").pack(side="left")
        self.y_var = tk.StringVar(value="%+.3f" % cal.get("y", 0.0))
        tk.Entry(row, width=8, textvariable=self.y_var).pack(
            side="left", padx=2)

        meas = tk.LabelFrame(
            self, text="Calculate from a printed alignment test",
            padx=8, pady=6)
        meas.pack(fill="x", padx=10, pady=4)
        tk.Label(meas, text="Print the test (plain letter paper) with the "
                            "offsets above, then measure:").pack(anchor="w")
        self.m_left = tk.StringVar()
        self.m_right = tk.StringVar()
        self.m_bottom = tk.StringVar()
        for label, var in (
                ("Paper left edge → box left side (target %.3f\"):"
                 % ALIGN_TARGET_LR, self.m_left),
                ("Paper right edge → box right side (target %.3f\", "
                 "optional):" % ALIGN_TARGET_LR, self.m_right),
                ("Paper top edge → box bottom line (target %.3f\", "
                 "optional):" % ALIGN_TARGET_BOTTOM, self.m_bottom)):
            r = tk.Frame(meas)
            r.pack(anchor="w", pady=1)
            tk.Label(r, text=label).pack(side="left")
            tk.Entry(r, width=8, textvariable=var).pack(side="left", padx=4)
        tk.Button(meas, text="Calculate offsets",
                  command=self._calc).pack(anchor="w", pady=(4, 0))

        self.msg = tk.Label(self, text="", wraplength=380, justify="left")
        self.msg.pack(anchor="w", padx=12)

        btns = tk.Frame(self)
        btns.pack(pady=10)
        tk.Button(btns, text="Print alignment test",
                  command=app.alignment_test).pack(side="left", padx=4)
        tk.Button(btns, text="Save", width=9, command=self._save).pack(
            side="left", padx=4)
        tk.Button(btns, text="Cancel", width=9, command=self.destroy).pack(
            side="left", padx=4)
        self.bind("<Escape>", lambda e: self.destroy())
        theme_toplevel(self)
        self.transient(app.root)
        self.grab_set()
        self.wait_window()

    def _calc(self):
        try:
            left = float(self.m_left.get())
            right_s = self.m_right.get().strip()
            right = float(right_s) if right_s else None
            bottom_s = self.m_bottom.get().strip()
            bottom = float(bottom_s) if bottom_s else None
            cur_x = float(self.x_var.get())
            cur_y = float(self.y_var.get())
        except ValueError:
            messagebox.showerror(
                APP_NAME, "Enter the measurements in inches (e.g. 2.25).",
                parent=self)
            return
        nx, ny = compute_calibration(cur_x, cur_y, left, right, bottom)
        self.x_var.set("%+.3f" % nx)
        self.y_var.set("%+.3f" % ny)
        self.msg.configure(
            text="Offsets updated from the measurements. Save to apply, "
                 "then reprint the test to verify.")

    def _save(self):
        try:
            x = float(self.x_var.get())
            y = float(self.y_var.get())
        except ValueError:
            messagebox.showerror(APP_NAME, "Offsets must be numbers "
                                           "(inches).", parent=self)
            return
        self.app.config["calibration"] = {"x": x, "y": y}
        save_config(self.app.config)
        self.app.set_status("Calibration set to x=%+.3f\", y=%+.3f\"" % (x, y))
        self.destroy()


class LoadSetDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Label sets")
        self.resizable(False, False)
        tk.Label(self, text="Saved label sets:").pack(
            anchor="w", padx=10, pady=(10, 2))
        self.listbox = tk.Listbox(self, width=40, height=8)
        self.listbox.pack(padx=10)
        self._fill()
        btns = tk.Frame(self)
        btns.pack(pady=10)
        tk.Button(btns, text="Load", width=9, command=self._load).pack(
            side="left", padx=4)
        tk.Button(btns, text="Delete", width=9, command=self._delete).pack(
            side="left", padx=4)
        tk.Button(btns, text="Close", width=9, command=self.destroy).pack(
            side="left", padx=4)
        self.listbox.bind("<Double-Button-1>", lambda e: self._load())
        self.bind("<Escape>", lambda e: self.destroy())
        theme_toplevel(self)
        self.transient(app.root)
        self.grab_set()
        self.wait_window()

    def _fill(self):
        self.listbox.delete(0, "end")
        for name in sorted(self.app.config.get("label_sets", {})):
            self.listbox.insert("end", name)

    def _selected(self):
        sel = self.listbox.curselection()
        return self.listbox.get(sel[0]) if sel else None

    def _load(self):
        name = self._selected()
        if not name:
            return
        self.app.load_dumps(self.app.config["label_sets"][name])
        self.app.set_status("Loaded label set \"%s\"" % name)
        self.destroy()

    def _delete(self):
        name = self._selected()
        if not name:
            return
        if messagebox.askyesno(APP_NAME, "Delete label set \"%s\"?" % name,
                               parent=self):
            del self.app.config["label_sets"][name]
            save_config(self.app.config)
            self._fill()


class LabelApp:
    def __init__(self, root):
        self.root = root
        self.config = load_config()
        families = set(tkfont.families())
        self.ui_family = FONT_NAME if FONT_NAME in families else FALLBACK_FONT
        self.measurer = TkMeasurer(self.ui_family)
        self._preview_job = None
        self.rows = []
        self._drag_from = None
        self._drag_target = None

        root.title("%s v%s" % (APP_NAME, APP_VERSION))
        try:
            root.iconbitmap(resource_path(ICON_FILE))
        except tk.TclError:
            pass
        root.configure(bg=CLR_BG)
        enable_dark_titlebar(root)
        self.active_row = None
        ttk_style = ttk.Style(root)
        try:
            ttk_style.theme_use("clam")
        except tk.TclError:
            pass
        ttk_style.configure("TNotebook", background=CLR_BG, borderwidth=0)
        ttk_style.configure("TNotebook.Tab", background=CLR_BTN,
                            foreground=CLR_FG, padding=(12, 5))
        ttk_style.map("TNotebook.Tab",
                      background=[("selected", CLR_ACCENT)],
                      foreground=[("selected", CLR_ON_ACCENT)])
        ttk_style.configure("TCombobox", fieldbackground=CLR_FIELD,
                            background=CLR_BTN, foreground=CLR_FG,
                            arrowcolor=CLR_FG, borderwidth=0)
        ttk_style.map("TCombobox",
                      fieldbackground=[("readonly", CLR_FIELD)],
                      foreground=[("readonly", CLR_FG)],
                      selectbackground=[("readonly", CLR_FIELD)],
                      selectforeground=[("readonly", CLR_FG)])
        root.option_add("*TCombobox*Listbox.background", CLR_FIELD)
        root.option_add("*TCombobox*Listbox.foreground", CLR_FG)
        root.option_add("*TCombobox*Listbox.selectBackground", CLR_ACCENT)
        root.option_add("*TCombobox*Listbox.selectForeground", CLR_ON_ACCENT)

        self._build_menu()
        self._build_ui()
        self._restore_autosave()
        self._bind_keys()
        self.update_preview()
        self.focus_row(0)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- construction ----------------------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self.root)

        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="New sheet", accelerator="Ctrl+N",
                           command=self.new_sheet)
        m_file.add_command(label="Start at position…",
                           command=self.start_at_position)
        m_file.add_separator()
        m_file.add_command(label="Save label set…", command=self.save_set)
        m_file.add_command(label="Load label set…",
                           command=lambda: LoadSetDialog(self))
        m_file.add_separator()
        m_file.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=m_file)

        m_print = tk.Menu(menubar, tearoff=0)
        m_print.add_command(label="Print…", accelerator="Ctrl+P",
                            command=self.print_select)
        m_print.add_command(label="Quick Print", accelerator="Ctrl+Shift+P",
                            command=self.quick_print)
        m_print.add_separator()
        m_print.add_command(label="Print alignment test (plain letter paper)",
                            command=self.alignment_test)
        m_print.add_command(label="Calibration…",
                            command=self.calibration_dialog)
        menubar.add_cascade(label="Print", menu=m_print)

        self.m_formats = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Formats", menu=self.m_formats)
        self._rebuild_formats_menu()

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="About / shortcuts", command=self.about)
        menubar.add_cascade(label="Help", menu=m_help)

        for m in (menubar, m_file, m_print, self.m_formats, m_help):
            m.configure(bg="#2a2a2a", fg=CLR_FG, bd=0,
                        activebackground=CLR_ACCENT,
                        activeforeground=CLR_ON_ACCENT)

        self.root.config(menu=menubar)

    def _rebuild_formats_menu(self):
        m = self.m_formats
        m.delete(0, "end")
        m.add_command(label="Edit formats…",
                      command=lambda: FormatsDialog(self))
        m.add_separator()
        for key in FORMAT_KEYS:
            name = self.config["formats"][key].get("name", "Format %s" % key)
            m.add_command(
                label="Apply %s (%s) to all labels" % (key, name),
                command=lambda k=key: self.apply_format_all(k))

    def _build_ui(self):
        header = tk.Frame(self.root, bg=CLR_HEADER)
        header.pack(fill="x")
        icon = tk.Canvas(header, width=24, height=24, bg=CLR_HEADER,
                         highlightthickness=0)
        icon.pack(side="left", padx=(12, 6), pady=8)
        icon.create_polygon(4, 6, 13, 6, 20, 12, 13, 18, 4, 18,
                            fill=CLR_ACCENT, outline="")
        icon.create_oval(7, 10, 11, 14, fill=CLR_HEADER, outline="")
        tk.Label(header, text=APP_NAME, bg=CLR_HEADER, fg="#ffffff",
                 font=("Segoe UI", 11, "bold")).pack(side="left")

        def hbtn(text, cmd, accent=False):
            b = tk.Button(header, text=text, command=cmd)
            style_button(b, accent)
            b.pack(side="right", padx=4, pady=9)
            return b

        hbtn("Quick Print", self.quick_print, accent=True)
        hbtn("Print…", self.print_select)
        hbtn("Formats", lambda: FormatsDialog(self))
        hbtn("Tutorial", self.show_tutorial)
        tk.Frame(self.root, height=1, bg="#000000").pack(fill="x")

        self.status = tk.Label(
            self.root, anchor="w", bg=CLR_HEADER, fg=CLR_DIM, padx=10,
            font=("Segoe UI", 8),
            text="Ready — Enter: next label · Shift+Enter: new line · "
                 "Ctrl+Tab: tab split · Ctrl+1/2/3: format")
        self.status.pack(fill="x", side="bottom")

        body = tk.Frame(self.root, bg=CLR_BG)
        body.pack(fill="both", expand=True)

        panel = tk.Frame(body, bg=CLR_PREVIEW_BG)
        panel.pack(side="right", fill="y")
        tk.Label(panel, text="Preview · 4 × 6 in", bg=CLR_PREVIEW_BG,
                 fg=CLR_DIM, font=("Segoe UI", 9)).pack(
            anchor="w", padx=16, pady=(10, 2))
        self.canvas = tk.Canvas(
            panel, width=int(SHEET_W * PREVIEW_SCALE) + 2 * PREVIEW_PAD + 4,
            height=int(SHEET_H * PREVIEW_SCALE) + 2 * PREVIEW_PAD + 4,
            bg=CLR_PREVIEW_BG, highlightthickness=0)
        self.canvas.pack(padx=12, pady=(0, 12), anchor="n")

        # Left column mirrors the panel's skeleton (caption + content) so the
        # rows below share the canvas's vertical origin; a spacer then drops
        # row 1 to the sheet's first-label offset.
        left = tk.Frame(body, bg=CLR_BG)
        left.pack(side="left", fill="both", expand=True, padx=(12, 8))
        tk.Label(left, text="Labels · Avery 5209", bg=CLR_BG, fg=CLR_DIM,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(10, 2))
        rows_area = tk.Frame(left, bg=CLR_BG)
        rows_area.pack(fill="both", expand=True)
        spacer_h = PREVIEW_PAD + round(MARGIN_TOP * PREVIEW_SCALE) - ROW_GAP
        tk.Frame(rows_area, bg=CLR_BG, height=spacer_h).pack(fill="x")
        for i in range(NUM_LABELS):
            self.rows.append(LabelRow(self, rows_area, i))

    def _bind_keys(self):
        self.root.bind_all("<Control-p>", lambda e: self.print_select())
        self.root.bind_all("<Control-P>", lambda e: self.quick_print())
        self.root.bind_all("<Control-n>", lambda e: self.new_sheet())
        for key in FORMAT_KEYS:
            self.root.bind_all(
                "<Control-Key-%s>" % key,
                lambda e, k=key: self._preset_focused_row(k))

    # -- helpers ---------------------------------------------------------
    def default_style(self):
        df = self.config.get("default_format")
        if df in FORMAT_KEYS:
            return copy.deepcopy(self.config["formats"][df]["style"])
        return copy.deepcopy(DEFAULT_STYLE)

    def blank_dump(self):
        df = self.config.get("default_format")
        return {"text": "", "style": self.default_style(),
                "preset": df if df in FORMAT_KEYS else None}

    def focus_row(self, index):
        self.rows[index].text.focus_set()

    def set_active_row(self, index):
        if self.active_row == index:
            return
        self.active_row = index
        for i, row in enumerate(self.rows):
            row.set_visual_active(i == index)
        self.schedule_preview()

    # -- row drag-reorder ------------------------------------------------
    def _row_at_y(self, y_root):
        """Index of the row under screen-Y, clamped to the row column."""
        for i, row in enumerate(self.rows):
            if y_root < row.frame.winfo_rooty() + row.frame.winfo_height():
                return i
        return NUM_LABELS - 1

    def _refresh_drag_marks(self):
        for i, row in enumerate(self.rows):
            row.set_visual_active(i == self.active_row)
        for idx in (self._drag_from, self._drag_target):
            if idx is not None:
                self.rows[idx].accent.configure(bg=CLR_ACCENT)

    def begin_row_drag(self, index):
        self._drag_from = index
        self._drag_target = index
        self._refresh_drag_marks()
        self.set_status("Moving label %d — drop on its new position"
                        % (index + 1))

    def update_row_drag(self, y_root):
        if self._drag_from is None:
            return
        target = self._row_at_y(y_root)
        if target != self._drag_target:
            self._drag_target = target
            self._refresh_drag_marks()

    def end_row_drag(self, y_root):
        if self._drag_from is None:
            return
        src, dst = self._drag_from, self._row_at_y(y_root)
        self._drag_from = None
        self._drag_target = None
        if dst != src:
            dumps = [row.dump() for row in self.rows]
            dumps.insert(dst, dumps.pop(src))
            for row, d in zip(self.rows, dumps):
                row.load(d)
            self.active_row = dst
            self.set_status("Moved label %d to position %d"
                            % (src + 1, dst + 1))
        else:
            self.set_status("Ready")
        self._refresh_drag_marks()
        self.schedule_preview()

    def _focused_row(self):
        widget = self.root.focus_get()
        for row in self.rows:
            if row.text is widget:
                return row
        return None

    def _preset_focused_row(self, key):
        row = self._focused_row()
        if row is not None:
            row.set_preset(key)

    def show_tutorial(self):
        dlg = getattr(self, "_tutorial", None)
        if dlg is not None and dlg.winfo_exists():
            dlg.lift()
            dlg.focus_set()
            return
        self._tutorial = TutorialDialog(self)

    def set_status(self, msg):
        self.status.configure(text=msg)

    def show_toast(self, title, body, ms=10000):
        """Borderless notification centered over the window; auto-hides
        after `ms` (click dismisses early). Non-modal."""
        t = tk.Toplevel(self.root)
        t.overrideredirect(True)
        t.attributes("-topmost", True)
        border = tk.Frame(t, bg=CLR_ACCENT)
        border.pack(fill="both", expand=True)
        inner = tk.Frame(border, bg=CLR_HEADER, padx=28, pady=18)
        inner.pack(fill="both", expand=True, padx=2, pady=2)
        tk.Label(inner, text=title, bg=CLR_HEADER, fg=CLR_ACCENT,
                 font=("Segoe UI", 12, "bold")).pack()
        tk.Label(inner, text=body, bg=CLR_HEADER, fg=CLR_FG,
                 font=("Segoe UI", 10)).pack(pady=(6, 0))
        t.update_idletasks()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        w, h = t.winfo_reqwidth(), t.winfo_reqheight()
        t.geometry("+%d+%d" % (rx + (rw - w) // 2, ry + (rh - h) // 2))

        def _close(_event=None):
            try:
                t.destroy()
            except tk.TclError:
                pass

        for widget in (t, border, inner):
            widget.bind("<Button-1>", _close)
        t.after(ms, _close)

    def dumps(self):
        return [row.dump() for row in self.rows]

    def load_dumps(self, dumps):
        for row, d in zip(self.rows, dumps):
            row.load(d)
        self.schedule_preview()

    # -- preview ---------------------------------------------------------
    def schedule_preview(self):
        if self._preview_job is not None:
            self.root.after_cancel(self._preview_job)
        self._preview_job = self.root.after(150, self.update_preview)

    def update_preview(self):
        self._preview_job = None
        c = self.canvas
        c.delete("all")
        ox = oy = PREVIEW_PAD

        def sx(x_in):
            return ox + x_in * PREVIEW_SCALE

        def sy(y_in):
            return oy + y_in * PREVIEW_SCALE

        c.create_rectangle(sx(0) + 4, sy(0) + 4, sx(SHEET_W) + 4,
                           sy(SHEET_H) + 4, fill="#0d0d0d", outline="")
        c.create_rectangle(sx(0), sy(0), sx(SHEET_W), sy(SHEET_H),
                           fill="white", outline="#e2e2e2")
        for i in range(NUM_LABELS):
            top = MARGIN_TOP + i * LABEL_H
            active = i == self.active_row
            c.create_rectangle(sx(MARGIN_LEFT), sy(top),
                               sx(MARGIN_LEFT + LABEL_W), sy(top + LABEL_H),
                               outline=CLR_ACCENT if active else "#c8c8c8",
                               dash=(3, 2))
            row = self.rows[i]
            text = row.get_text()
            if is_blank(text):
                continue
            fam = row.style.get("font") or self.ui_family
            for rx, ry, s, pt, bold in layout_label(text, row.style,
                                                    self.measurer):
                px_size = max(1, round(pt * PREVIEW_SCALE / 72.0))
                font = ((fam, -px_size, "bold") if bold
                        else (fam, -px_size))
                c.create_text(sx(MARGIN_LEFT + rx), sy(top + ry), text=s,
                              anchor="nw", font=font)

    # -- formats ---------------------------------------------------------
    def on_formats_changed(self):
        df = self.config.get("default_format")
        for row in self.rows:
            row.sync_from_preset()
            if df and row.preset is None and is_blank(row.get_text()):
                row.set_preset(df, silent=True)
        self._rebuild_formats_menu()
        self.schedule_preview()

    def apply_format_all(self, key):
        for row in self.rows:
            row.set_preset(key, silent=True)
        self.set_status("Applied format %s (%s) to all labels" %
                        (key, self.config["formats"][key].get("name", "")))
        self.schedule_preview()

    # -- sheet operations ------------------------------------------------
    def new_sheet(self):
        if any(not is_blank(r.get_text()) for r in self.rows):
            if not messagebox.askyesno(APP_NAME, "Clear all labels?"):
                return
        for row in self.rows:
            row.load(self.blank_dump())
        self.focus_row(0)
        self.set_status("New sheet")
        self.schedule_preview()

    def start_at_position(self):
        n = simpledialog.askinteger(
            "Start at position",
            "Sheet position for the first entered label (1-7):",
            parent=self.root, minvalue=1, maxvalue=NUM_LABELS)
        if not n or n == 1:
            return
        dumps = self.dumps()
        keep = dumps[:NUM_LABELS - (n - 1)]
        lost = dumps[NUM_LABELS - (n - 1):]
        if any(not is_blank(d["text"]) for d in lost):
            if not messagebox.askyesno(
                    APP_NAME,
                    "Labels at the bottom will be pushed off the sheet "
                    "and lost. Continue?"):
                return
        self.load_dumps([self.blank_dump() for _ in range(n - 1)] + keep)
        self.set_status("Labels shifted to start at position %d" % n)

    # -- label sets ------------------------------------------------------
    def save_set(self):
        name = simpledialog.askstring(
            "Save label set", "Name for this label set:", parent=self.root)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        sets = self.config.setdefault("label_sets", {})
        if name in sets and not messagebox.askyesno(
                APP_NAME, "Replace existing set \"%s\"?" % name):
            return
        sets[name] = self.dumps()
        save_config(self.config)
        self.set_status("Saved label set \"%s\"" % name)

    # -- printing --------------------------------------------------------
    def _require_win32(self):
        if not HAVE_WIN32:
            messagebox.showerror(
                APP_NAME,
                "Printing requires the pywin32 package.\n"
                "Install it with:  pip install pywin32")
            return False
        return True

    def _quick_printer(self):
        name = self.config.get("printer")
        if name:
            return name
        return win32print.GetDefaultPrinter()

    def _do_print(self, printer_name, outline=False):
        if outline:
            try:
                print_alignment_sheet(printer_name,
                                      self.config.get("calibration", {}))
            except Exception as e:
                messagebox.showerror(APP_NAME, "Print failed:\n%s" % e)
                return False
            self.set_status("Sent alignment test to %s" % printer_name)
            return True
        rows_data = [(r.get_text(), copy.deepcopy(r.style)) for r in self.rows]
        if all(is_blank(t) for t, _ in rows_data):
            messagebox.showinfo(APP_NAME, "All labels are blank — "
                                          "nothing to print.")
            return False
        try:
            print_sheet(printer_name, rows_data,
                        self.config.get("calibration", {}))
        except Exception as e:
            messagebox.showerror(APP_NAME, "Print failed:\n%s" % e)
            return False
        count = sum(1 for t, _ in rows_data if not is_blank(t))
        self.set_status("Sent %d label(s) to %s" % (count, printer_name))
        return True

    def print_select(self):
        if not self._require_win32():
            return
        dlg = PrinterDialog(self)
        if dlg.choice:
            self._do_print(dlg.choice)

    def quick_print(self):
        if not self._require_win32():
            return
        try:
            printer = self._quick_printer()
        except Exception as e:
            messagebox.showerror(APP_NAME, "No printer available:\n%s" % e)
            return
        if self._do_print(printer):
            self.show_toast("Printing", "Sending labels to %s" % printer,
                            ms=5000)

    def alignment_test(self):
        if not self._require_win32():
            return
        try:
            printer = self._quick_printer()
        except Exception as e:
            messagebox.showerror(APP_NAME, "No printer available:\n%s" % e)
            return
        if self._do_print(printer, outline=True):
            self.show_toast("Printing Alignment Sheet",
                            "Follow the instruction printed on the sheet")

    def calibration_dialog(self):
        CalibrationDialog(self)

    # -- persistence -----------------------------------------------------
    def _restore_autosave(self):
        saved = self.config.get("autosave")
        if isinstance(saved, list) and saved:
            self.load_dumps(saved[:NUM_LABELS])

    def on_close(self):
        self.config["autosave"] = self.dumps()
        save_config(self.config)
        self.root.destroy()

    def about(self):
        messagebox.showinfo(
            "About " + APP_NAME,
            "%s v%s\nAvery 5209 folder label printing (7 per 4\" x 6\" sheet)\n\n"
            "Shortcuts:\n"
            "  Enter / Tab — next label\n"
            "  Shift+Enter — second line (labels hold max 2 lines)\n"
            "  Ctrl+Tab — insert tab (splits Tab L/R columns)\n"
            "  Ctrl+1 / 2 / 3 — assign format to current label\n"
            "  Ctrl+P — print…   Ctrl+Shift+P — quick print\n"
            "  Ctrl+N — new sheet" % (APP_NAME, APP_VERSION))


def _run_smoke(root, app):
    style = normalize_style({"h_align": "tab"})
    runs = layout_label("Smith, John\t2026", style, app.measurer)
    assert len(runs) == 2, runs
    assert all(AUTO_MIN_PT <= r[3] <= AUTO_MAX_PT for r in runs)

    # per-line size/bold + line limit, and old-style migration
    style2 = normalize_style({"font_size": 14, "h_align": "center"})
    assert style2["lines"][0]["size"] == 14  # migrated label-wide size
    style2["lines"][0] = {"size": 0, "bold": True}
    style2["lines"][1] = {"size": 8, "bold": False}
    runs2 = layout_label("Line one\nLine two longer\nDROPPED", style2,
                         app.measurer)
    assert len(runs2) == 2, runs2
    assert runs2[0][4] is True and runs2[1][3] == 8

    # typing a size must flip that line to fixed mode (regression test)
    ed = StyleEditor(root, normalize_style({}))
    mode1, size1, _bold1 = ed.line_vars[1]
    size1.set(9)
    assert mode1.get() == "fixed"
    assert ed.get_style()["lines"][1]["size"] == 9
    assert ed.get_style()["lines"][0]["size"] == 0  # line 1 untouched: auto
    ed.destroy()

    # calibration math (letter-paper test): print landed 0.05" right and
    # 0.05" high -> correct left and down
    nx, ny = compute_calibration(0.0, 0.0, 2.30, 2.20, 6.20)
    assert abs(nx + 0.05) < 1e-9 and abs(ny - 0.05) < 1e-9, (nx, ny)
    nx2, ny2 = compute_calibration(0.1, -0.2, ALIGN_TARGET_LR,
                                   ALIGN_TARGET_LR, ALIGN_TARGET_BOTTOM)
    assert abs(nx2 - 0.1) < 1e-9 and abs(ny2 + 0.2) < 1e-9  # perfect = keep

    # gutter: second line must shift down by exactly the gutter amount
    style3 = normalize_style({"h_align": "left", "v_align": "top"})
    style3["lines"] = [{"size": 10, "bold": False}, {"size": 10, "bold": False}]
    base = layout_label("AA\nBB", style3, app.measurer)
    style3["gutter"] = 0.2
    spaced = layout_label("AA\nBB", style3, app.measurer)
    assert abs((spaced[1][1] - base[1][1]) - 0.2) < 0.001

    app.rows[0].set_text("keep 1\nkeep 2\ndropped")
    assert app.rows[0].get_text().count("\n") == 1
    app.rows[0].clear()
    assert app.rows[0].get_text() == ""
    app.rows[0].set_text("SMOKE TEST")
    app.show_toast("Smoke", "toast check", ms=50)
    app.update_preview()
    print("SMOKE OK: tab pt=%d, auto bold pt=%d, fixed line2 pt=%d"
          % (runs[0][3], runs2[0][3], runs2[1][3]))
    root.after(200, root.destroy)


def main():
    root = tk.Tk()
    app = LabelApp(root)
    if "--smoke" in sys.argv:
        root.after(400, lambda: _run_smoke(root, app))
    root.mainloop()


if __name__ == "__main__":
    main()
