"""
==============================================================================
NGLESHIE AMANFROM POLYCLINIC LABORATORY — REPORT GENERATOR
==============================================================================
Reads a Mindray BS240 (or similar analyser) PDF and produces a patient report
laid out exactly like the laboratory's Word template:

    • Ghana Health Service logo + centred laboratory name on every page
    • PATH. NO / SPECIMEN DATE / PATIENT NAME / SEX / AGE / SPECIMEN block
    • Red section titles:  RENAL FUNCTION TEST · LIVER FUNCTION TEST ·
                           LIPID PROFILE · GLYCATED HAEMOGLOBIN · URIC ACID
    • Borderless 5-column tables: INVESTIGATION | VALUE | UNIT | FLAG |
                                  REFERENCE INTERVAL
    • "Date printed …" / "Electronically Signed by QC manager" footer
    • "- -END OF REPORT- -" closing line

HOW TO RUN IT
-------------
  Double-click the file  →  the drag-and-drop window opens.

  From a terminal / PowerShell:
    python lab_report.py --demo              (sample renal/liver/lipid report)
    python lab_report.py --demo-hba1c        (sample HbA1c report)
    python lab_report.py "patient.pdf"
    python lab_report.py "patient.pdf" --path-no CHE901/26 --sex FEMALE
                          --age 46 --specimen SERUM
                          --specimen-date 27/04/2026 --output report.pdf

WHAT YOU MAY WANT TO EDIT
-------------------------
  SECTION 2  — laboratory name, signatory line, page geometry, SHOW_FLAGS
  SECTION 6  — reference intervals and units (TEST_LIBRARY)
  SECTION 5  — analyser short codes (TEST_CODE_MAP)
==============================================================================
"""

# ==============================================================================
# SECTION 1 — IMPORTS
# "Importing" means borrowing tools that other people already wrote.
# Think of it like getting equipment out of a toolbox before starting a job.
# ==============================================================================

import argparse     # Reads what you type after "python lab_report.py ..."
import base64       # Decodes the laboratory logo that is embedded in this file
import io           # Lets us treat the decoded logo as if it were a file
import os           # File paths, folders, "does this file exist?"
import re           # Text pattern matching (used all over the parser)
import sys          # Command-line arguments and clean exits
import threading    # Runs PDF generation in the background so the GUI responds

from datetime import datetime
# datetime = the current date and time, used for the "Date printed" footer.

# ── GUI imports (tkinter + optional drag-and-drop) ────────────────────────────
# Wrapped in try/except so the script still runs in CLI mode on machines that
# have no tkinter (e.g. a bare Linux server). Windows and macOS always have it.
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False
    tk = None

# Drag-and-drop is not part of tkinter: it comes from tkinterdnd2, which wraps
# a small Tcl library called tkdnd. Either can be missing — the package may not
# be installed, or the .exe may have been built without tkdnd's own files
# bundled alongside it, which is the usual reason dropping stops working in a
# packaged build while it still works when the script is run directly.
#
# Whatever went wrong is kept in DND_PROBLEM rather than thrown away, so the
# window can say why instead of quietly turning into a click-only program.
DND_PROBLEM = ""
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except Exception as _dnd_error:          # Not just ImportError: a packaged
    DND_AVAILABLE = False                # build can fail here in other ways
    TkinterDnD = None
    DND_FILES = None
    DND_PROBLEM = f"{type(_dnd_error).__name__}: {_dnd_error}"

# ── ReportLab: the library that actually draws the PDF ────────────────────────
from reportlab.lib.pagesizes import letter      # 8.5in × 11in paper
from reportlab.lib import colors                # Named and hex colours
from reportlab.lib.units import inch            # So we can write 0.4*inch
from reportlab.lib.styles import ParagraphStyle # Reusable text styles
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.utils import ImageReader     # Turns the logo bytes into art
from reportlab.platypus import (
    Paragraph,        # A block of text with a style applied to it
    Spacer,           # An invisible vertical gap
    Table,            # A grid of cells
    TableStyle,       # Rules controlling how a Table looks
    KeepTogether,     # Keeps a group of elements on the same page
    PageBreak,        # Starts the next element on a fresh page
    BaseDocTemplate,  # A document whose every page we can customise
    Frame,            # The rectangular content box on each page
    PageTemplate,     # "What does a page look like, and what runs on each one"
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ==============================================================================
# SECTION 2 — LABORATORY CONFIGURATION
# ✏️  These are the only lines most people ever need to change.
# ==============================================================================

LAB_NAME  = "NGLESHIE AMANFROM POLYCLINIC LABORATORY"
SIGNATORY = "Electronically Signed by QC manager"

# The two-digit year suffix used when building a path number like "CHE901/26".
# Leave as None to use the current year automatically.
PATH_NO_YEAR = None

# Set to True to have the script work out High / Low itself and print it in the
# FLAG column in red. Left False so the report matches the Word template, where
# the FLAG column stays blank unless the analyser itself printed a flag.
SHOW_FLAGS = False

# Print "- -END OF REPORT- -" after the last section.
SHOW_END_OF_REPORT = True

# ── The OTHER INVESTIGATIONS section ──────────────────────────────────────────
# Anything the analyser measured that LabScribe does not recognise is collected
# into a final "OTHER INVESTIGATIONS" section, so that a result is never
# silently dropped from a report.
#
# A glycated haemoglobin request is reported on its own, and the odd extra row
# an HbA1c run produces is not wanted on it — so that section is left off
# whenever HbA1c is being reported ON ITS OWN. On a report that carries
# chemistry as well, the section stays, because there it is holding real
# chemistry results that have nowhere else to go.
#
# Whatever is left off the report is still listed in the window, under the
# results for that file, so nothing goes unseen.
SHOW_OTHER_SECTION    = True    # False leaves it off every report
HIDE_OTHER_WITH_HBA1C = True    # Leave it off a glycated-haemoglobin-only report

# ── Two files from the same patient ───────────────────────────────────────────
# The laboratory often runs the chemistry and the glycated haemoglobin as two
# separate jobs, so one patient arrives as two analyser printouts. LabScribe
# recognises them as one patient and prints a single report carrying both.
#
# Set to False to go back to one report per file. When two files do belong to
# one patient, they are still never allowed to overwrite each other: the second
# report is given a distinct file name instead.
MERGE_SAME_PATIENT = True

# What the SPECIMEN line says when a merged report carries chemistry run on
# serum and a glycated haemoglobin run on whole blood.
MERGED_SPECIMEN_JOIN = " / "

# ── One section to a page ─────────────────────────────────────────────────────
# Each section starts a fresh page, so the lipid profile and the glycated
# haemoglobin are never found sharing a sheet. The patient header is drawn on
# every page, so each one stands on its own if it is handed over separately.
#
# Set to False to go back to packing as many sections onto a page as fit.
SECTION_PER_PAGE = True

# ── The application's own name, and who to write to about it ──────────────────
APP_NAME      = "LabScribe"
APP_TAGLINE   = "Clinical chemistry reporting"
APP_VERSION   = "2.4"
APP_PUBLISHER = "McLovin Software Solutions"
APP_EMAIL     = "abdulkhaliqsoule@gmail.com"
APP_PHONE     = "+233 20 960 0338"

# The six lines that print at the top of every report, and how they are
# labelled in the window. Editing a label here changes it in the window only —
# the wording printed on the report is set in SECTION 9.
PATIENT_FIELDS = [
    ("path_no",       "Path no."),
    ("specimen_date", "Specimen date"),
    ("name",          "Patient name"),
    ("sex",           "Sex"),
    ("age",           "Age (years)"),
    ("specimen",      "Specimen"),
]

# ── Page geometry (inches) ────────────────────────────────────────────────────
# These numbers were measured off the Word template so the PDF lines up with it.
PAGE_LEFT    = 0.40    # Left margin — the results table starts here
PAGE_RIGHT   = 0.40    # Right margin
PAGE_TOP     = 2.70    # Big top margin: the header block is drawn inside it
PAGE_BOTTOM  = 1.20    # Bottom margin: the footer line is drawn inside it

HDR_TEXT_X   = 1.00    # Patient lines start 1 inch in, as in the template
HDR_COL2_X   = 5.00    # SPECIMEN DATE / SEX start here
HDR_RULE_END = 7.27    # The horizontal rule stops here

LOGO_X, LOGO_TOP, LOGO_W, LOGO_H = 0.08, 1.08, 0.97, 0.92

RED = colors.HexColor("#FF0000")   # The section-title red used in the template


# ==============================================================================
# SECTION 2B — EMBEDDED GHANA HEALTH SERVICE LOGO
#
# The logo is stored inside this file as base64 text (a way of writing binary
# data using ordinary letters) so the script is a single self-contained file —
# there is no separate image to keep next to it or lose.
#
# ✏️  To use a different logo: save it as a PNG, then either
#       (a) put the file next to this script and name it "lab_logo.png"
#           — the script prefers that file whenever it exists — or
#       (b) replace the text below with the base64 of your own image.
# ==============================================================================

LOGO_FILENAME = "lab_logo.png"   # Optional override sitting beside the script

GHS_LOGO_B64 = """
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
AAAASUVORK5CYII=
"""


def _plain(text):
    """
    Strips the little formatting tags the PDF uses (<sub>, <super>, &lt;) so the
    same text can be shown as ordinary characters in the window.
    """
    text = "" if text is None else str(text)
    text = re.sub(r"</?(?:sub|super|b|i|font[^>]*)>", "", text)
    return (text.replace("&lt;", "<").replace("&gt;", ">")
                .replace("&amp;", "&").strip())


def get_logo_image():
    """
    Returns something ReportLab can draw as the laboratory logo, or None.

    Order of preference:
      1. A file called lab_logo.png sitting next to this script
      2. The Ghana Health Service logo embedded above
    """
    try:
        override = os.path.join(app_directory(), LOGO_FILENAME)
        if os.path.isfile(override):
            return ImageReader(override)
    except Exception:
        pass

    try:
        raw = base64.b64decode("".join(GHS_LOGO_B64.split()))
        return ImageReader(io.BytesIO(raw))
    except Exception:
        return None


# ==============================================================================
# SECTION 3 — FONTS
#
# The Word template uses Cambria (header), Algerian (red section titles),
# Berlin Sans FB (column headings) and Times New Roman (the results).
# Those first three only exist on Windows, so the script looks for them and
# quietly falls back to the fonts built into ReportLab when they are absent.
# The report is laid out identically either way — only the lettering differs.
# ==============================================================================

FONT_DIRS = [
    r"C:\Windows\Fonts",
    os.path.expanduser(r"~\AppData\Local\Microsoft\Windows\Fonts"),
    "/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts"),
    "/usr/share/fonts/truetype/msttcorefonts", "/usr/share/fonts",
]

# Filled in by _register_fonts(); every drawing routine reads these names.
F_HEADER      = "Times-Bold"    # Laboratory name + patient block
F_TITLE       = "Helvetica-Bold"  # Red section titles
F_COLHDR      = "Helvetica-Bold"  # INVESTIGATION / VALUE / UNIT / …
F_BODY        = "Times-Roman"   # Test names, units
F_BODY_BOLD   = "Times-Bold"    # Values and reference intervals
F_BODY_ITALIC = "Times-Italic"  # HbA1c test name
F_BODY_BI     = "Times-BoldItalic"  # eGFR explanatory note

# True once a real TrueType body font has been found. ReportLab's built-in
# fonts cannot draw µ or ≥, so when this stays False those characters are
# swapped for plain-text equivalents rather than silently disappearing.
HAVE_TTF_BODY = False


def _find_font_file(*filenames):
    """
    Returns the full path of the first font file that exists, else None.

    Windows keeps everything in one folder; Linux buries fonts a few folders
    deep, so each location is searched three levels down.
    """
    wanted = {name.lower(): name for name in filenames}

    for directory in FONT_DIRS:
        if not os.path.isdir(directory):
            continue

        # The plain case first — the file sitting right there
        for name in filenames:
            candidate = os.path.join(directory, name)
            if os.path.isfile(candidate):
                return candidate

        # Then a shallow walk, in case it is tucked into sub-folders
        try:
            base_depth = directory.rstrip(os.sep).count(os.sep)
            for root, dirs, files in os.walk(directory):
                if root.count(os.sep) - base_depth >= 3:
                    dirs[:] = []          # deep enough — stop descending
                    continue
                for file_name in files:
                    if file_name.lower() in wanted:
                        return os.path.join(root, file_name)
        except Exception:
            pass

    return None


def _try_register(internal_name, *filenames):
    """Registers a TrueType font if it can be found. Returns True on success."""
    path = _find_font_file(*filenames)
    if not path:
        return False
    try:
        pdfmetrics.registerFont(TTFont(internal_name, path))
        return True
    except Exception:
        return False


# The serif families to try for the results, best match first. Times New Roman
# is what the template uses; Liberation Serif has identical letter widths and
# ships with most Linux systems; DejaVu Serif is the last resort.
BODY_FONT_CANDIDATES = [
    # (regular, bold, italic, bold-italic)
    ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    ("Times New Roman.ttf", "Times New Roman Bold.ttf",
     "Times New Roman Italic.ttf", "Times New Roman Bold Italic.ttf"),
    ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf",
     "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"),
    ("DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf",
     "DejaVuSerif-Italic.ttf", "DejaVuSerif-BoldItalic.ttf"),
]


def _register_fonts():
    """Looks for the template's fonts and points the F_* names at them."""
    global F_HEADER, F_TITLE, F_COLHDR
    global F_BODY, F_BODY_BOLD, F_BODY_ITALIC, F_BODY_BI, HAVE_TTF_BODY

    # ── The results themselves — a serif face ─────────────────────────────────
    for regular, bold, italic, bold_italic in BODY_FONT_CANDIDATES:
        if not (_try_register("LabSerif", regular)
                and _try_register("LabSerif-Bold", bold)
                and _try_register("LabSerif-Italic", italic)
                and _try_register("LabSerif-BoldItalic", bold_italic)):
            continue
        try:
            pdfmetrics.registerFontFamily(
                "LabSerif", normal="LabSerif", bold="LabSerif-Bold",
                italic="LabSerif-Italic", boldItalic="LabSerif-BoldItalic")
        except Exception:
            continue
        F_BODY, F_BODY_BOLD = "LabSerif", "LabSerif-Bold"
        F_BODY_ITALIC, F_BODY_BI = "LabSerif-Italic", "LabSerif-BoldItalic"
        F_HEADER = "LabSerif-Bold"   # A sensible default until Cambria is found
        HAVE_TTF_BODY = True
        break

    # ── Cambria Bold — the laboratory name and the patient block ──────────────
    if _try_register("Cambria-Bold", "cambriab.ttf", "Cambria Bold.ttf"):
        F_HEADER = "Cambria-Bold"

    # ── Algerian — the red section titles ─────────────────────────────────────
    if _try_register("Algerian", "ALGER.TTF", "alger.ttf", "Algerian.ttf"):
        F_TITLE = "Algerian"

    # ── Berlin Sans FB Demi Bold — the column headings ────────────────────────
    if _try_register("BerlinSansFB-Bold", "BRLNSDB.TTF", "brlnsdb.ttf"):
        F_COLHDR = "BerlinSansFB-Bold"


_register_fonts()


# Characters the built-in ReportLab fonts have no glyph for. When no TrueType
# serif could be found, these are swapped for plain equivalents so a unit or a
# reference interval never comes out with a character missing.
PLAIN_SUBSTITUTES = {
    "µ": "u", "μ": "u",          # micro sign, Greek mu
    "≥": "&gt;=", "≤": "&lt;=",  # written as entities: these go inside markup
    "–": "-", "—": "-",          # en dash, em dash
    "≈": "~", "×": "x", "·": ".", "²": "2",
}


def safe_text(text, markup=True):
    """
    Returns text the current fonts can actually draw.

    markup=True  — the text is going into a Paragraph, where "<" and "&" have
                   a special meaning, so replacements are written as entities.
    markup=False — the text is being drawn straight onto the page.
    """
    text = "" if text is None else str(text)
    if HAVE_TTF_BODY:
        return text
    for character, replacement in PLAIN_SUBSTITUTES.items():
        if not markup:
            replacement = replacement.replace("&gt;", ">").replace("&lt;", "<")
        text = text.replace(character, replacement)
    return text


# ==============================================================================
# SECTION 4 — SECTIONS OF THE REPORT
#
# Each entry describes one red-titled block of the report:
#   "title"     — exactly what is printed in red, centred
#   "ref_label" — the wording of the last column heading
#   "show_flag" — whether the FLAG column appears at all
# ==============================================================================

RENAL = "RENAL FUNCTION TEST"
LIVER = "LIVER FUNCTION TEST"
LIPID = "LIPID PROFILE"
HBA1C = "GLYCATED HAEMOGLOBIN"
URIC  = "URIC ACID"
OTHER = "OTHER INVESTIGATIONS"

# "indent" and "cols" are inches, and only needed where a section's columns sit
# somewhere other than the usual five positions — the glycated haemoglobin page
# of the template is laid out differently from the rest, so it says so here.
SECTIONS = {
    RENAL: {"ref_label": "REFERENCE INTERVAL", "show_flag": True, "short": "RENAL",
            "cols": [1.329, 1.480, 1.500, 1.397, 1.997]},
    LIVER: {"ref_label": "REFERENCE INTERVAL", "show_flag": True, "short": "LIVER",
            "cols": [1.682, 1.333, 1.310, 1.142, 2.234]},
    LIPID: {"ref_label": "REFERENCE  VALUE",   "show_flag": True, "short": "LIPID",
            "cols": [2.123, 1.353, 0.922, 0.868, 2.433]},
    HBA1C: {"ref_label": "REFERENCE VALUE",    "show_flag": False, "short": "HbA1c",
            "indent": 0.60, "cols": [2.95, 0.85, 0.90, 2.40]},
    URIC:  {"ref_label": "REFERENCE VALUE",    "show_flag": True, "short": "URIC ACID"},
    OTHER: {"ref_label": "REFERENCE VALUE",    "show_flag": True, "short": "OTHER"},
}

# Top-to-bottom order of the sections in the finished report
SECTION_ORDER = [RENAL, LIVER, LIPID, HBA1C, URIC, OTHER]


# ==============================================================================
# SECTION 5 — TEST LIBRARY
#
# One row per test the laboratory reports. Everything the report needs to know
# about a test lives here, so this is the single place to edit when a reference
# interval changes or a new test is introduced.
#
#   name  — printed in the INVESTIGATION column, and used to order the rows
#   panel — which red-titled section it belongs to
#   unit  — printed in the UNIT column
#   ref   — one or more lines printed in the reference column
#   low / high — optional numeric limits, used only when SHOW_FLAGS is True
#   source_ref — when True, whatever the analyser printed for this test is used
#                instead of the "ref" line above, and "ref" becomes the fallback
#                for when the analyser printed nothing
#
# ✏️  Reference intervals below are the laboratory's own. Creatinine follows
#     the laboratory's Word template and uric acid follows the Tietz adult
#     intervals; both print on two lines, by sex. Every other test uses the
#     single fixed interval written here. No test now defers to the analyser,
#     but the source_ref mechanism remains if one should.
# ==============================================================================

EGFR_STAGES = [
    "Normal Kidney Function: 90+",
    "Mildly reduced Kidney Function: 60 - 89",
    "Moderately reduced Kidney Function: 30 – 59",
    "Severely reduced Kidney Function: 15 - 29",
    "End stage Kidney failure:&lt;15",
]

HBA1C_INTERPRETATION = [
    "NON – DIABETIC:   3.4% – 6.0%",
    "PREDIABETES: 5.7% - 6.4%",
    "DIABETIC:",
    "GOOD CONTROL: 6.1 – 8.0 %",
    "FAIR CONTROL:  8.1 – 10.0 %",
    "POOR CONTROL:   ABOVE 10 %",
]

EGFR_NOTE = (
    "Estimated Glomerular Filtration Rate (GFR) was calculated using the "
    "CKD-EPI (2021) creatinine equation, which does not include a race "
    "variable. eGFR is an estimate and should be interpreted alongside "
    "clinical findings, urine albumin-to-creatinine ratio (uACR), and patient "
    "history. Values may be affected by extremes of muscle mass, malnutrition, "
    "amputations, or acute illness. A single eGFR result is insufficient for "
    "CKD staging; confirmation requires two measurements at least 90 days "
    "apart. For clinical decisions, physician correlation is advised."
)

# ── Comments that depend on the result ────────────────────────────────────────
# The eGFR comment above is attached to its own row and prints under it. These
# ones belong to a whole section and only appear when a result crosses a
# figure, so they are kept here with the figure that triggers them.
#
# ✏️  To reword a comment, edit the text. To change when it appears, edit the
#     number beside "above" in SECTION_COMMENTS.

TG_NOTE = (
    "Triglycerides above 4.5 mmol/L. At this concentration the calculated "
    "LDL-cholesterol is unreliable, and the specimen may be lipaemic, which "
    "can affect other results in this profile. A repeat on a sample taken "
    "after a 9 to 12 hour fast is advised. For clinical decisions, physician "
    "correlation is advised."
)

TEST_LIBRARY = [
    # ── Renal Function ────────────────────────────────────────────────────────
    # Printed on two lines, as the laboratory's own Word template prints it
    {"name": "CREATININE", "panel": RENAL, "unit": "µmol/L",
     "ref": ["Female:  53.0 – 120.0", "Male:    71.0 – 140.0"],
     "low_f": 53.0, "high_f": 120.0, "low_m": 71.0, "high_m": 140.0},

    {"name": "UREA", "panel": RENAL, "unit": "mmol/L",
     "ref": ["1.7 – 8.3"], "low": 1.7, "high": 8.3},

    {"name": "SODIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["135 - 150"], "low": 135, "high": 150},

    {"name": "POTASSIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["3.5 - 5.5"], "low": 3.5, "high": 5.5},

    {"name": "CHLORIDE", "panel": RENAL, "unit": "mmol/L",
     "ref": ["98.0 - 110.0"], "low": 98.0, "high": 110.0},

    {"name": "TOTAL CO2", "panel": RENAL, "unit": "mmol/L",
     "ref": ["22.0 - 29.0"], "low": 22.0, "high": 29.0},

    {"name": "CALCIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["2.15 - 2.55"], "low": 2.15, "high": 2.55},

    {"name": "PHOSPHATE", "panel": RENAL, "unit": "mmol/L",
     "ref": ["0.81 - 1.45"], "low": 0.81, "high": 1.45},

    {"name": "MAGNESIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["0.75 - 1.00"], "low": 0.75, "high": 1.00},

    {"name": "eGFR (CKD-EPI)", "panel": RENAL, "unit": "ml/min/1.73m<super>2</super>",
     "ref": EGFR_STAGES, "low": 90, "high": None, "note": EGFR_NOTE},

    # ── Liver Function ────────────────────────────────────────────────────────
    {"name": "TOTAL PROTEIN", "panel": LIVER, "unit": "g/L",
     "ref": ["60.0 – 87.0"], "low": 60.0, "high": 87.0},

    {"name": "ALBUMIN", "panel": LIVER, "unit": "g/L",
     "ref": ["35.0 – 50.0"], "low": 35.0, "high": 50.0},

    {"name": "GLOBULIN", "panel": LIVER, "unit": "g/L",
     "ref": ["25.0 – 35.0"], "low": 25.0, "high": 35.0},

    {"name": "A/G RATIO", "panel": LIVER, "unit": "",
     "ref": ["1.0 – 2.2"], "low": 1.0, "high": 2.2},

    {"name": "BILIRUBIN (TOTAL)", "panel": LIVER, "unit": "µmol/L",
     "ref": ["1.0 – 22.0"], "low": 1.0, "high": 22.0},

    {"name": "BILIRUBIN (DIRECT)", "panel": LIVER, "unit": "µmol/L",
     "ref": ["1.0 – 9.0"], "low": 1.0, "high": 9.0},

    {"name": "BILIRUBIN (INDIRECT)", "panel": LIVER, "unit": "µmol/L",
     "ref": ["0.0 – 17.0"], "low": 0.0, "high": 17.0},

    {"name": "GAMMA GT", "panel": LIVER, "unit": "U/L",
     "ref": ["5.0 – 50.0"], "low": 5.0, "high": 50.0},

    {"name": "ALT (SGPT)", "panel": LIVER, "unit": "U/L",
     "ref": ["0.0 – 45.0"], "low": 0.0, "high": 45.0},

    {"name": "AST (SGOT)", "panel": LIVER, "unit": "U/L",
     "ref": ["0.0 – 40.0"], "low": 0.0, "high": 40.0},

    {"name": "AST/ALT RATIO", "panel": LIVER, "unit": "",
     "ref": ["0.8 – 1.4"], "low": 0.8, "high": 1.4},

    {"name": "ALKALINE PHOSPHATASE", "panel": LIVER, "unit": "U/L",
     "ref": ["40 – 150"], "low": 40.0, "high": 150.0},

    # ── Lipid Profile ─────────────────────────────────────────────────────────
    {"name": "TOTAL CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["2.80 – 5.20"], "low": 2.80, "high": 5.20},

    {"name": "HDL-CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["≥ 0.90"], "low": 0.90, "high": None},

    {"name": "LDL-CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["&lt;4.00"], "low": None, "high": 4.00},

    {"name": "TRIGLYCERIDES", "panel": LIPID, "unit": "mmol/L",
     "ref": ["0.50 - 1.80"], "low": 0.50, "high": 1.80},

    {"name": "VLDL-CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["0.10 - 1.00"], "low": 0.10, "high": 1.00},

    {"name": "CORONARY RISK", "panel": LIPID, "unit": "",
     "ref": ["&lt;4.0"], "low": None, "high": 4.0},

    {"name": "LDL/HDL RATIO", "panel": LIPID, "unit": "",
     "ref": ["&lt;3.5"], "low": None, "high": 3.5},

    # ── Glycated Haemoglobin ──────────────────────────────────────────────────
    {"name": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)", "panel": HBA1C,
     "unit": "%", "ref": HBA1C_INTERPRETATION,
     "italic": True, "low": 3.4, "high": 6.0},

    # ── Uric Acid ─────────────────────────────────────────────────────────────
    # Tietz adult reference intervals, the values in general laboratory use:
    # men 3.5 - 7.2 mg/dL, women 2.6 - 6.0 mg/dL, converted at 59.48
    {"name": "URIC ACID", "panel": URIC, "unit": "µmol/L",
     "ref": ["Female:  155 – 357", "Male:    208 – 428"],
     "low_f": 155, "high_f": 357, "low_m": 208, "high_m": 428},
]

# Quick lookups built once from the library above
TEST_INFO  = {t["name"]: t for t in TEST_LIBRARY}
TEST_ORDER = {t["name"]: i for i, t in enumerate(TEST_LIBRARY)}


# ==============================================================================
# SECTION 6 — ANALYSER SHORT CODES
#
# The BS240 prints terse codes ("T-bil", "CREA", "γ-GT"). This dictionary turns
# each one into the name used in the report. Everything is lower-case here so
# matching ignores capitalisation.
#
# "__SKIP__" removes the row from the report entirely — used for the HbA1c
# sub-fractions (IFCC, eAG) that the laboratory does not report.
#
# ✏️  Add a line here whenever the analyser produces an abbreviation the report
#     does not recognise. The name on the right must exist in TEST_LIBRARY.
# ==============================================================================

TEST_CODE_MAP = {
    # ── Renal ─────────────────────────────────────────────────────────────────
    "crea": "CREATININE", "creat": "CREATININE", "creatinine": "CREATININE",
    "urea": "UREA", "bun": "UREA", "blood urea nitrogen (bun)": "UREA",
    "na+": "SODIUM", "na": "SODIUM", "sodium": "SODIUM",
    "k+": "POTASSIUM", "k": "POTASSIUM", "potassium": "POTASSIUM",
    "cl-": "CHLORIDE", "cl": "CHLORIDE", "chloride": "CHLORIDE",
    "tco2": "TOTAL CO2", "co2": "TOTAL CO2",
    "carbon dioxide, total": "TOTAL CO2", "bicarbonate": "TOTAL CO2",
    "ca": "CALCIUM", "ca2+": "CALCIUM", "calcium": "CALCIUM",
    "phos": "PHOSPHATE", "phosphate": "PHOSPHATE",
    "mg": "MAGNESIUM", "magnesium": "MAGNESIUM",
    "egfr": "eGFR (CKD-EPI)", "estimated gfr": "eGFR (CKD-EPI)",
    "egfr (ckd-epi)": "eGFR (CKD-EPI)", "egfr (ckd-epi 2021)": "eGFR (CKD-EPI)",

    # ── Liver ─────────────────────────────────────────────────────────────────
    "tp": "TOTAL PROTEIN", "tp-o": "TOTAL PROTEIN",
    "total protein": "TOTAL PROTEIN", "protein, total": "TOTAL PROTEIN",
    "alb": "ALBUMIN", "albumin": "ALBUMIN",
    "glob": "GLOBULIN", "globulin": "GLOBULIN", "globulin, total": "GLOBULIN",
    "a/g": "A/G RATIO", "a/g ratio": "A/G RATIO",
    "t-bil": "BILIRUBIN (TOTAL)", "tbil": "BILIRUBIN (TOTAL)",
    "tb": "BILIRUBIN (TOTAL)", "total bilirubin": "BILIRUBIN (TOTAL)",
    "bilirubin, total": "BILIRUBIN (TOTAL)",
    "d-bil": "BILIRUBIN (DIRECT)", "dbil": "BILIRUBIN (DIRECT)",
    "direct bil": "BILIRUBIN (DIRECT)", "direct bilirubin": "BILIRUBIN (DIRECT)",
    "ibil": "BILIRUBIN (INDIRECT)", "ind bill": "BILIRUBIN (INDIRECT)",
    "ind bil": "BILIRUBIN (INDIRECT)", "ind-bil": "BILIRUBIN (INDIRECT)",
    "indirect bilirubin": "BILIRUBIN (INDIRECT)",
    # Depending on how the analyser's PDF encodes the Greek gamma, the code can
    # come out of the text extractor as γ-GT, y-GT or g-GT — all three are here.
    "γ-gt": "GAMMA GT", "y-gt": "GAMMA GT", "g-gt": "GAMMA GT",
    "ggt": "GAMMA GT", "ggtp": "GAMMA GT", "gamma gt": "GAMMA GT",
    "gamma-gt": "GAMMA GT", "gamma glutamyl transferase": "GAMMA GT",
    "alt": "ALT (SGPT)", "sgpt": "ALT (SGPT)", "alt (sgpt)": "ALT (SGPT)",
    "alanine aminotrans": "ALT (SGPT)",
    "ast": "AST (SGOT)", "sgot": "AST (SGOT)", "ast (sgot)": "AST (SGOT)",
    "aspartate aminotrans": "AST (SGOT)",
    "ast/alt": "AST/ALT RATIO",
    "alp": "ALKALINE PHOSPHATASE",
    "alkaline phos.": "ALKALINE PHOSPHATASE",
    "alkaline phosphatase": "ALKALINE PHOSPHATASE",

    # ── Lipids ────────────────────────────────────────────────────────────────
    "tc": "TOTAL CHOLESTEROL", "chol": "TOTAL CHOLESTEROL",
    "total chol": "TOTAL CHOLESTEROL", "total chol.": "TOTAL CHOLESTEROL",
    "cholesterol": "TOTAL CHOLESTEROL",
    "cholesterol, total": "TOTAL CHOLESTEROL",
    "hdl": "HDL-CHOLESTEROL", "hdl-c": "HDL-CHOLESTEROL",
    "hdl-chol.": "HDL-CHOLESTEROL", "hdl cholesterol": "HDL-CHOLESTEROL",
    "ldl": "LDL-CHOLESTEROL", "ldl-c": "LDL-CHOLESTEROL",
    "ldl cholesterol": "LDL-CHOLESTEROL",
    "ldl chol calc (nih)": "LDL-CHOLESTEROL",
    "tg": "TRIGLYCERIDES", "trig": "TRIGLYCERIDES",
    "triglyceride": "TRIGLYCERIDES", "triglycerides": "TRIGLYCERIDES",
    "vldl": "VLDL-CHOLESTEROL", "vldl-c": "VLDL-CHOLESTEROL",
    "vldl cholesterol cal": "VLDL-CHOLESTEROL",
    "coronary r": "CORONARY RISK", "coronary risk": "CORONARY RISK",
    "chol/hdl": "CORONARY RISK",
    "ldl/hdl": "LDL/HDL RATIO",

    # ── Glycated haemoglobin — only the NGSP percentage is reported ───────────
    "a1c(ngsp)": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "a1c (ngsp)": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "ngsp": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "hba1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "hb a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "hemoglobin a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "haemoglobin a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "glycated": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "glycated haemoglobin (hba1c)": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "a1c(ifcc)": "__SKIP__", "a1c (ifcc)": "__SKIP__", "ifcc": "__SKIP__",
    "eag": "__SKIP__", "estimated average glucose": "__SKIP__",

    # ── Uric acid ─────────────────────────────────────────────────────────────
    "ua": "URIC ACID", "urate": "URIC ACID", "uric acid": "URIC ACID",
}


def resolve_test_name(raw_name):
    """
    Turns an analyser code into the name used in the report.

    Returns (display_name, panel). If the code is unknown the original name is
    returned unchanged and the panel is "OTHER INVESTIGATIONS", so nothing the
    analyser measured is ever silently lost.
    """
    key = (raw_name or "").strip().lower()

    # Some parsers hand us "CODE — FULL NAME"; the code alone is what we want.
    if " — " in key:
        key = key.split(" — ")[0].strip()

    # Trailing punctuation and stray spaces are common in analyser output
    key = key.strip(" .:")

    name = TEST_CODE_MAP.get(key)
    if name is None:
        # Maybe the parser already gave us a proper report name
        for candidate in TEST_INFO:
            if candidate.lower() == key:
                name = candidate
                break

    if name is None:
        return raw_name.strip(), OTHER
    if name == "__SKIP__":
        return "__SKIP__", ""
    return name, TEST_INFO[name]["panel"]


# ==============================================================================
# SECTION 7 — TIDYING UP THE PARSED RESULTS
#
# apply_test_library() is the bridge between the parser (which reads whatever
# the analyser printed) and the report (which wants the laboratory's own names,
# units and reference intervals). It runs once, before anything is drawn.
# ==============================================================================

def _to_number(text):
    """Turns '4.01' or '<0.20' into a number, or None if it isn't numeric."""
    if text is None:
        return None
    cleaned = re.sub(r"[^\d.\-]", "", str(text))
    if cleaned in ("", "-", ".", "-."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def compute_flag(name, value, sex=""):
    """
    Works out "High" / "Low" / "" for one result.
    Only consulted when SHOW_FLAGS is True.

    A test whose interval differs by sex carries low_f/high_f and low_m/high_m.
    When the patient's sex is not known, the widest span of the two is used, so
    a result is never flagged on a guess.
    """
    info = TEST_INFO.get(name)
    if not info:
        return ""
    number = _to_number(value)
    if number is None:
        return ""

    if "low_f" in info:
        first = str(sex or "").strip().upper()[:1]
        if first == "F":
            low, high = info.get("low_f"), info.get("high_f")
        elif first == "M":
            low, high = info.get("low_m"), info.get("high_m")
        else:
            low = min(info.get("low_f"), info.get("low_m"))
            high = max(info.get("high_f"), info.get("high_m"))
    else:
        low, high = info.get("low"), info.get("high")
    if high is not None and number > high:
        return "High"
    if low is not None and number < low:
        return "Low"
    return ""


def tidy_reference(text):
    """
    Smartens a reference interval as printed by the analyser so it matches the
    look of the rest of the report: "53.0-120.0" becomes "53.0 – 120.0", and
    the characters that upset the PDF markup are escaped.
    """
    text = str(text or "").strip()
    if not text:
        return ""

    # A hyphen sitting between two numbers is a range dash, not a minus sign
    text = re.sub(r"(?<=[\d.])\s*[-–—]\s*(?=[\d.])", " – ", text)

    # "<" and "&" would be read as markup inside a Paragraph
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return text


# ── Glycated haemoglobin: one row, and it has to be the NGSP percentage ───────
#
# One HbA1c run can print the same haemoglobin two or three ways: the NGSP
# percentage, the IFCC figure in mmol/mol, an estimated average glucose, and on
# some settings a second channel named HbA1c-2. The laboratory reports the NGSP
# percentage and nothing else, so the whole family is gathered up here and one
# row is chosen from it before anything is printed. The choice is made on the
# unit and the size of the figure, not on the name alone, because the naming
# differs from one analyser setting to the next.

HBA1C_NGSP_NAME = "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)"

HBA1C_WORDS = ("a1c", "glycated", "glycosylated", "ngsp", "ifcc",
               "estimated average glucose", "eag")

NGSP_LOW, NGSP_HIGH = 3.0, 20.0      # A percentage this size is the NGSP
IFCC_LOW, IFCC_HIGH = 20.0, 200.0    # A figure this size is the IFCC


def _hba1c_text(row):
    return f"{row.get('test', '')} {row.get('full_name', '')}".lower()


def is_hba1c_row(row):
    """True if this result is any of the glycated haemoglobin figures."""
    text = _hba1c_text(row)
    return any(word in text for word in HBA1C_WORDS)


def hba1c_kind(row):
    """Which figure this row carries: ngsp, ifcc, eag, or unknown."""
    text = _hba1c_text(row)
    unit = str(row.get("unit", "")).lower().replace(" ", "")

    if "eag" in text or "average glucose" in text:
        return "eag"
    if "ifcc" in text or "mmol/mol" in unit:
        return "ifcc"
    if "ngsp" in text or "%" in unit:
        return "ngsp"

    number = _to_number(row.get("result", ""))
    if number is not None:
        if NGSP_LOW <= number <= NGSP_HIGH:
            return "ngsp"
        if IFCC_LOW < number <= IFCC_HIGH:
            return "ifcc"
    return "unknown"


def choose_hba1c(results):
    """
    Cuts the glycated haemoglobin family down to the single NGSP percentage.

    Returns (rows, dropped). The kept row sits where the first of the family
    sat, so the report keeps its order.
    """
    family = [row for row in results if is_hba1c_row(row)]
    if not family:
        return list(results), []

    def preference(row):
        text = _hba1c_text(row)
        # A row that says NGSP outright is taken first; a second channel
        # (HbA1c-2, A1c-2) is a repeat of the same test and is never the one
        # reported, so it is taken last.
        return (0 if "ngsp" in text else 1,
                1 if re.search(r"-\s*\d(\b|$)", text) else 0)

    percentages = sorted([row for row in family if hba1c_kind(row) == "ngsp"],
                         key=preference)
    chosen = percentages[0] if percentages else None

    if chosen is None:
        # Only the IFCC figure was printed, so work the percentage out from it
        # with the master equation:  NGSP % = 0.09148 × IFCC + 2.152
        for row in family:
            if hba1c_kind(row) != "ifcc":
                continue
            number = _to_number(row.get("result", ""))
            if number is None:
                continue
            chosen = dict(row)
            chosen["result"] = f"{0.09148 * number + 2.152:.2f}"
            chosen["unit"] = "%"
            print("   HbA1c: the analyser printed only the IFCC figure "
                  f"({row.get('result', '')} mmol/mol), so the NGSP "
                  f"percentage was worked out from it ({chosen['result']} %).")
            break

    if chosen is not None:
        chosen = dict(chosen)
        chosen["test"] = HBA1C_NGSP_NAME
        chosen["full_name"] = ""
        chosen["unit"] = "%"

    kept, dropped, placed = [], [], False
    for row in results:
        if not is_hba1c_row(row):
            kept.append(row)
            continue
        if not placed and chosen is not None:
            kept.append(chosen)
            placed = True
        dropped.append(row)

    return kept, dropped


def apply_test_library(results, patient=None):
    """
    Takes the raw list of results from the parser and returns a clean list
    ready for printing:

      • analyser codes replaced with the laboratory's own test names
      • __SKIP__ rows (HbA1c IFCC, eAG) dropped
      • units and reference intervals filled in from TEST_LIBRARY
      • tests marked source_ref keep the interval the analyser printed
      • each row tagged with the section it belongs to
      • a flag computed when SHOW_FLAGS is on

    The analyser's own unit and reference interval are kept whenever
    TEST_LIBRARY has nothing to say about that test.
    """
    sex = (patient or {}).get("sex", "") if isinstance(patient, dict) else ""
    cleaned = []
    seen = set()

    # One glycated haemoglobin row, carrying the NGSP percentage
    results, _dropped = choose_hba1c(results)

    for raw in results:
        name, panel = resolve_test_name(raw.get("test", ""))
        if name == "__SKIP__":
            continue

        # Guard against an analyser printing the same test twice
        if name in seen:
            continue
        seen.add(name)

        row = dict(raw)
        row["test"] = name
        row["panel"] = panel

        source_ref = tidy_reference(raw.get("ref", ""))

        info = TEST_INFO.get(name)
        if info:
            row["unit"] = info["unit"]

            if info.get("source_ref") and source_ref:
                # Creatinine, ALP, uric acid: print what the analyser gave
                row["ref_lines"] = [source_ref]
            else:
                row["ref_lines"] = list(info["ref"])

            if info.get("note"):
                row["note"] = info["note"]
            if info.get("italic"):
                row["italic"] = True
        else:
            # Unknown test: show whatever the analyser printed
            row["ref_lines"] = [source_ref] if source_ref else []

        if SHOW_FLAGS and not row.get("flag"):
            row["flag"] = compute_flag(name, row.get("result", ""), sex)

        cleaned.append(row)

    return cleaned


def group_results_by_section(results):
    """
    Sorts the cleaned results into their sections and puts each section's rows
    into the order used by the laboratory's template.

    Returns a dictionary: {"RENAL FUNCTION Test": [row, row, …], …}
    """
    grouped = {section: [] for section in SECTION_ORDER}

    for row in results:
        panel = row.get("panel", OTHER)
        if panel not in grouped:
            panel = OTHER
        grouped[panel].append(row)

    for section in grouped:
        grouped[section].sort(
            key=lambda r: TEST_ORDER.get(r["test"], 9999)
        )

    return grouped


# ==============================================================================
# SECTION 8 — TEXT STYLES
#
# Like the Styles panel in Word: define each kind of text once, then apply it
# by name everywhere it is needed.
# ==============================================================================

def styles():
    return dict(
        # Red centred section title, e.g. "RENAL FUNCTION Test"
        section_title=ParagraphStyle(
            "SectionTitle", fontName=F_TITLE, fontSize=18, leading=24,
            textColor=RED, alignment=TA_CENTER, spaceAfter=0),

        # Column headings: INVESTIGATION | VALUE | UNIT | FLAG | REFERENCE …
        col_hdr=ParagraphStyle(
            "ColHdr", fontName=F_COLHDR, fontSize=11.5, leading=14,
            textColor=colors.black, alignment=TA_CENTER),

        # Test name — left aligned, regular weight
        test_name=ParagraphStyle(
            "TestName", fontName=F_BODY, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_LEFT),

        # Test name for HbA1c — italic, as in the template
        test_name_italic=ParagraphStyle(
            "TestNameItalic", fontName=F_BODY_ITALIC, fontSize=10, leading=13,
            textColor=colors.black, alignment=TA_LEFT),

        # The measured value — centred and bold
        value=ParagraphStyle(
            "Value", fontName=F_BODY_BOLD, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_CENTER),

        # Unit — centred, regular, slightly smaller
        unit=ParagraphStyle(
            "Unit", fontName=F_BODY, fontSize=10, leading=13,
            textColor=colors.black, alignment=TA_CENTER),

        # FLAG column — bold red when there is something to say
        flag=ParagraphStyle(
            "Flag", fontName=F_BODY_BOLD, fontSize=11, leading=14,
            textColor=RED, alignment=TA_CENTER),

        # Reference interval — bold, left aligned inside its column.
        # spaceAfter puts a gap between the Female and Male lines, and between
        # the kidney-function stages, exactly as the Word template does.
        ref=ParagraphStyle(
            "Ref", fontName=F_BODY_BOLD, fontSize=10, leading=14,
            spaceAfter=8, textColor=colors.black, alignment=TA_LEFT),

        # The eGFR explanatory paragraph — bold italic, justified
        note=ParagraphStyle(
            "Note", fontName=F_BODY_BI, fontSize=9.5, leading=15,
            textColor=colors.black, alignment=TA_JUSTIFY),

        # "- -END OF REPORT- -"
        end_of_report=ParagraphStyle(
            "EndOfReport", fontName=F_BODY_BI, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_CENTER),
    )


# ==============================================================================
# SECTION 9 — THE PAGE ITSELF
#
# LabReportDoc is a document that knows how to draw the laboratory header and
# the footer on every single page. Everything above the rule and below the
# results is painted here; the results themselves flow in between.
# ==============================================================================

class LabReportDoc(BaseDocTemplate):

    def __init__(self, filename, patient, **kwargs):
        super().__init__(filename, **kwargs)
        self.patient = patient or {}
        self.logo = get_logo_image()

        frame = Frame(
            self.leftMargin, self.bottomMargin,
            self.width, self.height, id="body",
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        )
        self.addPageTemplates([
            PageTemplate(id="All", frames=frame, onPage=self._on_page)
        ])

    # ── Drawn automatically at the start of every page ────────────────────────
    def _on_page(self, canvas, doc):
        canvas.saveState()
        self._draw_header(canvas)
        self._draw_footer(canvas)
        canvas.restoreState()

    def _draw_header(self, canvas):
        p = self.patient
        page_w, page_h = self.pagesize

        # ── Logo, top left ────────────────────────────────────────────────────
        if self.logo is not None:
            try:
                canvas.drawImage(
                    self.logo,
                    LOGO_X * inch,
                    page_h - (LOGO_TOP + LOGO_H) * inch,
                    width=LOGO_W * inch, height=LOGO_H * inch,
                    mask="auto", preserveAspectRatio=True, anchor="sw",
                )
            except Exception:
                pass   # A missing logo must never stop a report printing

        # ── Laboratory name, centred ──────────────────────────────────────────
        canvas.setFont(F_HEADER, 14)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(page_w / 2.0, page_h - 0.98 * inch,
                                 safe_text(LAB_NAME, markup=False))

        # ── Patient block ─────────────────────────────────────────────────────
        canvas.setFont(F_HEADER, 11)
        x1 = HDR_TEXT_X * inch
        x2 = HDR_COL2_X * inch

        def line(y_from_top, left_text, right_text=None):
            y = page_h - y_from_top * inch
            canvas.drawString(x1, y, safe_text(left_text, markup=False))
            if right_text:
                canvas.drawString(x2, y, safe_text(right_text, markup=False))

        line(1.34, f"PATH. NO: {p.get('path_no', '')}",
                   f"SPECIMEN DATE: {p.get('specimen_date', '')}")
        line(1.66, f"PATIENT NAME: {p.get('name', '')}",
                   f"SEX: {p.get('sex', '')}")
        line(2.08, f"AGE: {p.get('age', '')} YEARS")
        line(2.31, f"SPECIMEN: {p.get('specimen', '')}")

        # ── The rule that closes the header ───────────────────────────────────
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(1.0)
        y_rule = page_h - 2.41 * inch
        canvas.line(x1, y_rule, HDR_RULE_END * inch, y_rule)

    def _draw_footer(self, canvas):
        # 0.92 inch up from the bottom edge — measured off the Word template
        y = 0.92 * inch

        printed = self.patient.get("date_printed") or \
            datetime.today().strftime("%d-%b-%y")

        canvas.setFillColor(colors.black)
        canvas.setFont(F_HEADER, 11)
        canvas.drawString(HDR_TEXT_X * inch, y, "Date printed ")
        label_w = canvas.stringWidth("Date printed ", F_HEADER, 11)
        canvas.setFont(F_BODY, 11)
        canvas.drawString(HDR_TEXT_X * inch + label_w, y, printed)

        canvas.setFont(F_HEADER, 11)
        canvas.drawString(4.55 * inch, y, safe_text(SIGNATORY, markup=False))


# ==============================================================================
# SECTION 10 — BUILDING ONE SECTION
#
# Returns the list of drawing instructions for a single red-titled block:
# the title, the column headings, and one row per result.
# ==============================================================================

# Column widths as fractions of the usable page width, measured off the
# Word template so the columns land in the same place.
COL_FRACTIONS_5 = [0.1726, 0.1921, 0.1948, 0.1813, 0.2592]  # with FLAG
COL_FRACTIONS_4 = [0.1726, 0.1921, 0.1948, 0.4405]          # without FLAG


def _ref_paragraphs(lines, s):
    """Turns one or more reference lines into a stack of paragraphs."""
    return [Paragraph(safe_text(line), s["ref"])
            for line in lines if str(line).strip()]


SECTION_COMMENTS = [
    {"panel": LIPID, "test": "TRIGLYCERIDES", "above": 4.5, "text": TG_NOTE},
]


def comments_for_section(section_name, rows):
    """
    The comments this section has earned, given what was measured.

    Returns a list of paragraphs to print under the section's last row. A
    comment appears once, however many rows would have triggered it.
    """
    notes = []
    for rule in SECTION_COMMENTS:
        if rule["panel"] != section_name or rule["text"] in notes:
            continue
        for row in rows:
            if row.get("test") != rule["test"]:
                continue
            number = _to_number(row.get("result", ""))
            if number is None:
                continue
            if "above" in rule and number > rule["above"]:
                notes.append(rule["text"])
                break
            if "below" in rule and number < rule["below"]:
                notes.append(rule["text"])
                break
    return notes


def build_section(section_name, rows, W, s):
    """
    Builds one section of the report.

    section_name — e.g. "RENAL FUNCTION Test"
    rows         — the results belonging to it, already in template order
    W            — usable page width in points
    s            — the styles dictionary
    """
    config = SECTIONS.get(section_name, SECTIONS[OTHER])
    show_flag = config["show_flag"]

    # ── Where the columns sit ─────────────────────────────────────────────────
    # Most sections use the five positions measured off the template. A section
    # may override them with its own widths, and ask for a left indent — the
    # indent is simply an extra empty column at the front.
    indent = config.get("indent", 0) * inch
    if config.get("cols"):
        widths = [c * inch for c in config["cols"]]
        overflow = (indent + sum(widths)) - W
        if overflow > 0:                    # Never run off the edge of the page
            widths[-1] -= overflow
    else:
        fractions = COL_FRACTIONS_5 if show_flag else COL_FRACTIONS_4
        widths = [(W - indent) * f for f in fractions]

    if indent:
        widths = [indent] + widths

    def with_indent(cells):
        """Adds the empty spacer cell when the section is indented."""
        return ([""] + cells) if indent else cells

    block = []

    # ── The red centred title ─────────────────────────────────────────────────
    block.append(Paragraph(safe_text(section_name), s["section_title"]))
    block.append(Spacer(1, 20))

    # ── Column headings ───────────────────────────────────────────────────────
    heads = ["INVESTIGATION", "VALUE", "UNIT"]
    if show_flag:
        heads.append("FLAG")
    heads.append(config["ref_label"])

    hdr_table = Table(
        [with_indent([Paragraph(h, s["col_hdr"]) for h in heads])],
        colWidths=widths,
    )
    hdr_table.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "BOTTOM"),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
    ]))
    block.append(hdr_table)
    block.append(Spacer(1, 14))

    # ── One table row per result ──────────────────────────────────────────────
    for row in rows:
        name_style = s["test_name_italic"] if row.get("italic") else s["test_name"]

        cells = [
            Paragraph(safe_text(row["test"]), name_style),
            Paragraph(safe_text(row.get("result", "")), s["value"]),
            Paragraph(safe_text(row.get("unit", "")), s["unit"]),
        ]
        if show_flag:
            cells.append(Paragraph(safe_text(row.get("flag", "") or ""), s["flag"]))
        cells.append(_ref_paragraphs(row.get("ref_lines", []), s))

        row_table = Table([with_indent(cells)], colWidths=widths)
        row_table.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING",   (0, 0), (-1, -1), 2),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
        ]))
        block.append(row_table)

        # ── The eGFR explanatory paragraph, if this row carries one ───────────
        # It sits under the first columns while the reference column above it
        # continues with the CKD staging list — exactly as in the template.
        if row.get("note"):
            note_width = sum(widths[:-1])
            note_table = Table(
                [[Paragraph(safe_text(row["note"]), s["note"])]],
                colWidths=[note_width], hAlign="LEFT",
            )
            note_table.setStyle(TableStyle([
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING",    (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING",   (0, 0), (-1, -1), 12),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
            ]))
            block.append(note_table)

    # ── A comment for the whole section, where a result has earned one ────────
    # Set the same way as the eGFR paragraph, under the section's last row.
    for text in comments_for_section(section_name, rows):
        note_width = sum(widths[:-1])
        comment = Table([[Paragraph(safe_text(text), s["note"])]],
                        colWidths=[note_width], hAlign="LEFT")
        comment.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING",   (0, 0), (-1, -1), 12),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ]))
        block.append(comment)

    return block


# ==============================================================================
# SECTION 11 — BUILDING THE WHOLE REPORT
# ==============================================================================

def generate_report(patient, results, output_path):
    """
    Assembles and saves the finished PDF.

    patient     — dict with path_no, specimen_date, name, sex, age, specimen
    results     — the flat list of results from the parser
    output_path — where to save the PDF
    """
    doc = LabReportDoc(
        output_path, patient,
        pagesize=letter,
        leftMargin=PAGE_LEFT * inch,
        rightMargin=PAGE_RIGHT * inch,
        topMargin=PAGE_TOP * inch,
        bottomMargin=PAGE_BOTTOM * inch,
        title=f"Laboratory Report for {patient.get('name', '')}",
        author=LAB_NAME,
    )

    s = styles()
    W = doc.width
    story = []

    # The eGFR is worked out again here, from the age and the sex as they
    # finally stand, because both can have been typed over in the window since
    # the file was read. Doing it at this one point means every route into the
    # report — the window, the command line, a merged pair of files — gets the
    # same treatment.
    results = recompute_egfr(results, patient)

    cleaned = apply_test_library(results, patient)
    grouped = group_results_by_section(cleaned)

    # Which sections actually have something in them
    present = [name for name in SECTION_ORDER if grouped.get(name)]

    # ── Should OTHER INVESTIGATIONS be printed? ───────────────────────────────
    # It is left off a glycated haemoglobin report, and can be switched off
    # altogether in SECTION 2. Anything left off is still shown in the window.
    if OTHER in present:
        # HIDE_OTHER_WITH_HBA1C is about the single clean page a glycated
        # haemoglobin request is reported on. A merged report carrying
        # chemistry as well is not that page, and its OTHER INVESTIGATIONS
        # section is holding real chemistry results with nowhere else to go —
        # so there the section stays.
        hba1c_only = HBA1C in present and not [name for name in present
                                               if name not in (HBA1C, OTHER)]
        hide = (not SHOW_OTHER_SECTION
                or (HIDE_OTHER_WITH_HBA1C and hba1c_only))
        if hide:
            present.remove(OTHER)
            dropped = ", ".join(row["test"] for row in grouped[OTHER])
            print(f"   Not printed (OTHER INVESTIGATIONS): {dropped}")

    for position, section_name in enumerate(present):
        block = build_section(section_name, grouped[section_name], W, s)

        # The closing line belongs with the last section, not adrift on a page
        # of its own — so it goes inside that section's block.
        if SHOW_END_OF_REPORT and position == len(present) - 1:
            block.append(Spacer(1, 26))
            block.append(Paragraph("- -END OF REPORT- -", s["end_of_report"]))
        elif not SECTION_PER_PAGE:
            # Only needed when sections share a page. With one section to a
            # page the gap is what produces a blank sheet: a section that fills
            # its page exactly pushes the spacer onto the next one, and the
            # page break then pushes the following section past it.
            block.append(Spacer(1, 26))

        # Each section starts a fresh page, so that two of them are never found
        # sharing a sheet. The patient header is drawn on every page, so a page
        # handed over on its own still says whose results it is carrying.
        if SECTION_PER_PAGE and position > 0:
            story.append(PageBreak())

        # KeepTogether moves the whole section to the next page rather than
        # splitting it, unless it is too tall to fit on a page at all — in
        # which case ReportLab splits it as it must.
        story.append(KeepTogether(block))

    doc.build(story)
    print(f"Report saved: {output_path}")
    return output_path


# ==============================================================================
# SECTION 12 — DEMO DATA
# Fake but realistic data so you can see a finished report without an analyser.
# Reproduces the two Word templates the report was modelled on.
# ==============================================================================

DEMO_PATIENT = {
    "path_no":       "CHE901/26",
    "specimen_date": "27/04/2026",
    "name":          "ADWOA SAMPLE BOATENG",
    "sex":           "FEMALE",
    "age":           "46",
    "specimen":      "SERUM",
}

DEMO_RESULTS = [
    {"test": "CREATININE",           "result": "66.0"},
    {"test": "UREA",                 "result": "2.8"},
    {"test": "SODIUM",               "result": "141.4"},
    {"test": "POTASSIUM",            "result": "4.7"},
    {"test": "CHLORIDE",             "result": "103.1"},
    {"test": "eGFR (CKD-EPI)",       "result": "100"},
    {"test": "TOTAL PROTEIN",        "result": "74.6"},
    {"test": "ALBUMIN",              "result": "47.1"},
    {"test": "GLOBULIN",             "result": "27.5"},
    {"test": "BILIRUBIN (TOTAL)",    "result": "15.6"},
    {"test": "BILIRUBIN (DIRECT)",   "result": "4.9"},
    {"test": "BILIRUBIN (INDIRECT)", "result": "10.7"},
    {"test": "GAMMA GT",             "result": "14.8"},
    {"test": "ALT (SGPT)",           "result": "11.4"},
    {"test": "AST (SGOT)",           "result": "15.3"},
    {"test": "ALKALINE PHOSPHATASE", "result": "52.3"},
    {"test": "TOTAL CHOLESTEROL",    "result": "4.01"},
    {"test": "HDL-CHOLESTEROL",      "result": "1.62"},
    {"test": "LDL-CHOLESTEROL",      "result": "2.12"},
    {"test": "TRIGLYCERIDES",        "result": "0.59"},
    {"test": "CORONARY RISK",        "result": "2.5"},
]

DEMO_HBA1C_PATIENT = {
    "path_no":       "CHE902/26",
    "specimen_date": "15/07/2026",
    "name":          "AKUA EXAMPLE",
    "sex":           "FEMALE",
    "age":           "32",
    "specimen":      "WHOLE BLOOD",
}

DEMO_HBA1C_RESULTS = [
    {"test": "A1c(NGSP)", "result": "5.64"},
]


# The application mark, drawn into the title bar. Base64 PNG.
APP_ICON_B64 = (
    "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"
    "/n/tTg2OxPgxTwAAAABJRU5ErkJggg=="
)

# ==============================================================================
# HELPER — build_template_patient()
#
# The parser reads whatever the analyser printed at the top of its PDF, using
# its own field names (Patient, Sample ID, Gender, Test Date …). The report
# header needs six specific fields instead:
#
#     PATH. NO · SPECIMEN DATE · PATIENT NAME · SEX · AGE · SPECIMEN
#
# This function does that translation, and tidies the values on the way:
# names and specimen types are put in capitals, the age is reduced to a plain
# number, and the date is rewritten in day/month/year order.
# ==============================================================================

# The Mindray BS240 prints dates American-style (month/day/year).
# Set this to False if your analyser already prints day/month/year.
ANALYSER_DATES_ARE_US = True


def _reformat_date(text):
    """
    Rewrites a date as DD/MM/YYYY, which is what the template shows.
    Anything it cannot make sense of is passed straight through unchanged.
    """
    if not text:
        return ""

    raw = str(text).strip()

    # Keep only the date part — analysers often append a time
    match = re.search(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{2,4})", raw)
    if not match:
        return raw

    a, b, c = match.group(1), match.group(2), match.group(3)

    # ISO order: 2026-04-27
    if len(a) == 4:
        year, month, day = a, b, c
    else:
        year = c
        if ANALYSER_DATES_ARE_US:
            month, day = a, b
        else:
            day, month = a, b
        # A "month" above 12 can only be a day — the analyser swapped them
        try:
            if int(month) > 12 >= int(day):
                day, month = month, day
        except ValueError:
            pass

    if len(year) == 2:
        year = "20" + year

    try:
        return f"{int(day):02d}/{int(month):02d}/{year}"
    except ValueError:
        return raw


def _year_suffix(date_text):
    """Returns the two-digit year used to build a path number, e.g. '26'."""
    if PATH_NO_YEAR:
        return str(PATH_NO_YEAR)[-2:]
    match = re.search(r"(\d{4})\s*$", str(date_text or "").strip())
    if match:
        return match.group(1)[-2:]
    return datetime.today().strftime("%y")


def build_template_patient(parsed, results=None):
    """
    Turns the parser's patient dictionary into the six fields the report header
    prints. The original keys are kept as well, so nothing is lost.

    The results are passed in as well because one of the six fields depends on
    them: a glycated haemoglobin is run on whole blood, whatever sample type
    the analyser happens to have been told about.
    """
    patient = dict(parsed or {})

    # ── SPECIMEN DATE ─────────────────────────────────────────────────────────
    # Prefer the collection date; fall back to whatever date we do have.
    raw_date = (parsed.get("collected") or parsed.get("received")
                or parsed.get("reported") or "")
    if str(raw_date).strip().upper() in ("", "N/A"):
        raw_date = ""
    specimen_date = _reformat_date(raw_date)

    # ── PATH. NO ──────────────────────────────────────────────────────────────
    # The laboratory's path number looks like "CHE901/26". The analyser usually
    # prints the "CHE901" part as the Patient ID or the Sample ID.
    path_no = ""
    for key in ("req", "id"):
        candidate = str(parsed.get(key, "")).strip()
        if candidate and candidate.upper() != "N/A":
            path_no = candidate
            break
    if path_no and "/" not in path_no:
        path_no = f"{path_no}/{_year_suffix(specimen_date)}"

    # ── AGE — digits only ─────────────────────────────────────────────────────
    age = re.sub(r"[^\d]", "", str(parsed.get("age", "")))

    # ── SEX ───────────────────────────────────────────────────────────────────
    sex = str(parsed.get("gender", "")).strip().upper()
    if sex in ("N/A", "UNKNOWN"):
        sex = ""

    # ── SPECIMEN TYPE ─────────────────────────────────────────────────────────
    specimen = str(parsed.get("sample_type", "")).strip().upper()
    if specimen in ("N/A", ""):
        specimen = "SERUM"          # The laboratory's usual specimen

    # A glycated haemoglobin is always run on whole blood in EDTA, and the
    # analyser often still prints "serum", so it is corrected here. Only for a
    # run that is glycated haemoglobin and nothing else: where the same
    # printout also carries chemistry, the sample really was serum.
    if results and any(is_hba1c_row(row) for row in results):
        chemistry = [row for row in results
                     if not is_hba1c_row(row)
                     and resolve_test_name(row.get("test", ""))[1]
                     not in (OTHER, "")]
        if not chemistry:
            specimen = "WHOLE BLOOD"

    # ── PATIENT NAME ──────────────────────────────────────────────────────────
    name = str(parsed.get("name", "")).strip().upper()
    if name in ("UNKNOWN", "N/A"):
        name = ""

    patient.update({
        "path_no":       path_no,
        "specimen_date": specimen_date,
        "name":          name,
        "sex":           sex,
        "age":           age,
        "specimen":      specimen,
    })
    return patient


# ==============================================================================
# SECTION 12B — TWO FILES, ONE PATIENT
#
# The laboratory runs the chemistry and the glycated haemoglobin as separate
# jobs, so one patient often arrives as two analyser printouts. Everything
# needed to recognise that and print the two as a single report lives here.
#
# Until version 2.4 both files were reported separately, and because the file
# name was built from the path number and the patient's name alone, the two
# reports were given the same name and the second wrote silently over the
# first. One of the two reports simply was not there afterwards.
# ==============================================================================

def _name_key(text):
    """
    A patient's name reduced to something two printouts can be matched on.

    Punctuation, spacing and the order of the names are all thrown away, so
    "BOATENG, ADWOA SAMPLE" and "Adwoa Sample Boateng" come out the same.
    """
    words = re.sub(r"[^A-Za-z]+", " ", str(text or "")).upper().split()
    return " ".join(sorted(words))


def patient_merge_key(patient):
    """
    What decides that two files belong to the same patient: the name, and only
    the name.

    Deliberately so. Where the chemistry and the glycated haemoglobin were run
    as two separate jobs the two printouts carry different sample numbers, and
    often different dates as well, so neither can be used to hold them
    together. A file with no name on it is never merged with anything.
    """
    return _name_key((patient or {}).get("name", "")) or None


def _has_chemistry(results):
    """True if this file carries anything besides a glycated haemoglobin."""
    for row in results or []:
        if is_hba1c_row(row):
            continue
        name, _panel = resolve_test_name(row.get("test", ""))
        if name and name != "__SKIP__":
            return True
    return False


def _chemistry_first(members):
    """
    The files of a group, chemistry printouts ahead of glycated haemoglobin
    ones. sorted() is stable, so files otherwise keep the order they were added
    in. The chemistry printout leads because it is the one carrying the path
    number the laboratory files the request under.
    """
    return sorted(members, key=lambda m: 0 if _has_chemistry(m[1]) else 1)


def merge_patient_fields(members):
    """
    Builds the one set of header fields a merged report prints.

    `members` is a list of (patient, results) pairs, in the order the files
    were added to the window.
    """
    if len(members) == 1:
        return dict(members[0][0] or {})

    # Start from everything the parser found, first file winning, so that keys
    # outside the six printed fields are not lost on the way through.
    merged = {}
    for patient, _results in reversed(members):
        merged.update({key: value for key, value in (patient or {}).items()
                       if str(value).strip()})

    # ── The six printed fields: the first file that has one wins ──────────────
    for key, _label in PATIENT_FIELDS:
        merged[key] = ""
        for patient, _results in members:
            candidate = str((patient or {}).get(key, "")).strip()
            if candidate:
                merged[key] = candidate
                break

    ordered = _chemistry_first(members)

    # ── PATH. NO and SPECIMEN DATE come off the chemistry printout ────────────
    for key in ("path_no", "specimen_date"):
        for patient, _results in ordered:
            candidate = str((patient or {}).get(key, "")).strip()
            if candidate:
                merged[key] = candidate
                break

    # ── SPECIMEN ──────────────────────────────────────────────────────────────
    # Each file already carries the right sample type for what is on it:
    # build_template_patient() sets WHOLE BLOOD on a glycated haemoglobin run
    # and leaves the analyser's own sample type on a chemistry one. A merged
    # report is printed from two different samples, so it says so, rather than
    # claiming the whole report came off one of them.
    specimens = []
    for patient, _results in ordered:
        value = str((patient or {}).get("specimen", "")).strip().upper()
        if value and value not in specimens:
            specimens.append(value)
    if specimens:
        merged["specimen"] = MERGED_SPECIMEN_JOIN.join(specimens)

    return merged


def merge_results(members):
    """
    One flat list of results for a merged report.

    The lists are laid end to end and nothing else is done to them, because the
    two steps that follow already tidy up the joins: choose_hba1c() cuts the
    glycated haemoglobin family down to the single NGSP percentage however many
    files it arrived across, and apply_test_library() keeps only the first row
    for any given test. Chemistry goes first, so where a test was run on both
    printouts it is the chemistry figure that prints.
    """
    combined = []
    for _patient, results in _chemistry_first(members):
        combined.extend(results or [])
    return combined


def recompute_egfr(results, patient):
    """
    Throws away an eGFR LabScribe worked out earlier and works it out again
    from the age and the sex as they now stand.

    The figure depends on both, and both can be typed over in the window after
    the file was read — so it has to be worked out at the moment the report is
    written, not left as it was when the PDF was first opened. An eGFR the
    analyser itself printed is left alone.
    """
    kept = [row for row in (results or []) if not row.get("computed")]
    return calculate_egfr(kept, patient)


def merged_job(group):
    """
    Turns a group of files into the (patient, results) pair one report is
    built from. A group of one goes through unchanged, apart from the eGFR
    being worked out again.
    """
    members = [(job.get("patient") or {}, job.get("results") or [])
               for job in group]
    patient = merge_patient_fields(members)
    results = recompute_egfr(merge_results(members), patient)
    return patient, results


def group_jobs_by_patient(jobs):
    """
    Sorts the files waiting to be reported into one group per patient.

    Returns a list of lists, each holding the files that print as a single
    report, in the order they were added. A file with no patient name on it, or
    one marked "keep separate" in the window, is always a group of its own.
    """
    groups, position_of = [], {}

    for job in jobs:
        key = None
        if MERGE_SAME_PATIENT and not job.get("separate"):
            key = patient_merge_key(job.get("patient") or {})

        if key is not None and key in position_of:
            groups[position_of[key]].append(job)
            continue

        if key is not None:
            position_of[key] = len(groups)
        groups.append([job])

    return groups


def unique_output_path(out_dir, filename, taken):
    """
    Makes certain no two reports in one run are written to the same name.

    `taken` is a set carried across the run, holding the names already used. A
    name that has come up before is given a number rather than being allowed to
    write over the report already sitting there. Names left over from earlier
    runs are not counted, so generating the same work again still replaces it.
    """
    stem, extension = os.path.splitext(filename)
    candidate, counter = filename, 2

    while candidate.lower() in taken:
        candidate = f"{stem}_{counter}{extension}"
        counter += 1

    taken.add(candidate.lower())
    return os.path.join(out_dir, candidate)


# ==============================================================================
# SECTION 13 — READING THE ANALYSER'S PDF
#
# Everything from here to SECTION 14 has one job: opening the PDF the analyser
# produced and turning it into a plain list of results. None of it knows
# anything about how the finished report looks.
# ==============================================================================

# ==============================================================================
# SECTION 13A — extract_text_from_pdf() FUNCTION
#
# Opens a PDF exported by the Mindray BS240 analyser and extracts all text
# AND positional word data so we can reconstruct the column layout.
#
# WHY POSITIONAL EXTRACTION?
#   The BS240 produces PDFs in (at least) two different layouts:
#
#   Layout A — Old-style (Weija Gbawe format):
#     Uses real table borders → pdfplumber's extract_tables() works perfectly.
#     Columns: Test | Full Name | Concentration | Unit | Results | Remark | Reference
#
#   Layout B — New-style (Yaa Specimen / CHE900 format):
#     NO table borders — just text positioned at fixed x-coordinates.
#     extract_tables() returns 0 tables so the old approach gets nothing.
#     The columns are identified purely by their horizontal (x) position:
#       x ≈   7  → Test name (e.g. "T-bil", "ALT", "CORONARY R")
#       x ≈ 116  → Result value (e.g. "43.7", "0.8")
#       x ≈ 228  → Unit (e.g. "U/L", "mmol/L", "μmol/L")
#       x ≈ 303  → Flag (just "H" or "L" — only present if abnormal)
#       x ≈ 398  → Reference range (e.g. "0.0-41.0", "44.0-106.0")
#
# This function now returns a dict with THREE keys:
#   "header_text"   → full plain text string (used for patient field regex)
#   "table_rows"    → rows from extract_tables() — populated in Layout A only
#   "word_rows"     → rows built from positional words — populated for Layout B
#                     Each row: [test_name, result, unit, flag, ref_range]
# ==============================================================================

def extract_text_from_pdf(pdf_path):
    """
    Extracts header text, table rows (if any), and positional word rows
    from a Mindray BS240 PDF export.

    Returns:
      {
        "header_text": "Patient: YAA SPECIMEN ...",
        "table_rows":  [...],   # non-empty only when real table borders exist
        "word_rows":   [        # always populated from positional grouping
            ["ALT",  "43.7",  "U/L",    "H", "0.0-41.0"],
            ["AST",  "34.1",  "U/L",    "",  "0.0-40.0"],
            ["CREA", "123.1", "μmol/L", "H", "44.0-106.0"],
            ...
        ]
      }

    Requires pdfplumber — install with:  pip install pdfplumber
    """

    try:
        import pdfplumber
    except ImportError:
        print("❌ pdfplumber is not installed.")
        print("   Run:  pip install pdfplumber")
        raise

    header_text = ""   # Plain text for patient field regex
    table_rows  = []   # Rows from extract_tables() — only for bordered-table PDFs
    word_rows   = []   # Rows reconstructed from x-position grouping

    # Everything learned about the file on the way through, kept so that a file
    # which cannot be read can say why rather than just "could not be read".
    diagnostics = {
        "path": pdf_path,
        "pages": 0,
        "images": 0,
        "text_characters": 0,
        "table_rows": 0,
        "word_rows": 0,
        "near_misses": [],   # Rows with a test name but no number beside it
        "route": "",
        "results_found": 0,
        "unknown_codes": [],
        "notes": [],
    }

    with pdfplumber.open(pdf_path) as pdf:
        diagnostics["pages"] = len(pdf.pages)

        for page in pdf.pages:
            try:
                diagnostics["images"] += len(page.images or [])
            except Exception:
                pass

            # ── Plain text: used for patient header fields ─────────────────────
            plain = page.extract_text()
            if plain:
                header_text += plain + "\n"

            # ── Table extraction: works when table borders exist ───────────────
            # This succeeds for Layout A (old Weija Gbawe style).
            # For Layout B it returns an empty list — we handle that below.
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or all(c is None or str(c).strip() == "" for c in row):
                        continue
                    cleaned = [str(c).strip() if c is not None else "" for c in row]
                    table_rows.append(cleaned)

            # ── Positional word grouping: works for both layouts ───────────────
            # extract_words() returns every word with its x,y coordinates.
            # We group words that share the same vertical position (same row)
            # then assign each word to a column based on its x coordinate.
            #
            # COLUMN BOUNDARIES (from inspecting the actual PDF word positions):
            #   x < 100     → Column 0: Test name
            #   100 ≤ x < 200 → Column 1: Result value
            #   200 ≤ x < 280 → Column 2: Unit
            #   280 ≤ x < 360 → Column 3: Flag (H or L)
            #   360 ≤ x       → Column 4: Reference range
            #
            # These boundaries work for both the old and new BS240 PDF formats
            # because the result columns are always in the same relative positions.
            #
            words = page.extract_words()

            # ── Group the words into rows ─────────────────────────────────────
            # Words that look level on the page rarely have identical
            # coordinates: a letter with a descender (the γ of "γ-GT") or a
            # slightly smaller unit sits a point or two off its neighbours.
            #
            # So rather than rounding each word to a grid — which splits a row
            # whenever two words fall either side of a boundary — we sort the
            # words down the page and start a new row only when the gap from
            # the current one is bigger than ROW_TOLERANCE.
            #
            # Each word is measured at its vertical middle rather than its top,
            # which ignores differences in letter height altogether.
            ROW_TOLERANCE = 4.0     # points

            def middle(w):
                return (w["top"] + w["bottom"]) / 2.0

            rows = []                    # list of (middle, [words])
            for word in sorted(words, key=middle):
                centre = middle(word)
                if rows and abs(centre - rows[-1][0]) <= ROW_TOLERANCE:
                    rows[-1][1].append(word)
                else:
                    rows.append((centre, [word]))

            # Process each row in top-to-bottom order
            for _centre, words_in_row in rows:
                row_words = sorted(words_in_row, key=lambda w: w["x0"])

                # Assign each word to a column slot based on its x position
                # slots[0] = test name words, slots[1] = result, etc.
                slots = ["", "", "", "", ""]   # 5 slots for 5 columns

                for word in row_words:
                    x = word["x0"]
                    text = word["text"]

                    if x < 100:
                        # Column 0: Test name — words can be multi-word so we append
                        slots[0] = (slots[0] + " " + text).strip()
                    elif x < 200:
                        # Column 1: Result value
                        slots[1] = (slots[1] + " " + text).strip()
                    elif x < 280:
                        # Column 2: Unit (e.g. "U/L", "mmol/L", "μmol/L")
                        slots[2] = (slots[2] + " " + text).strip()
                    elif x < 360:
                        # Column 3: Flag — only "H" or "L" will appear here
                        slots[3] = (slots[3] + " " + text).strip()
                    else:
                        # Column 4: Reference range (e.g. "0.0-41.0")
                        slots[4] = (slots[4] + " " + text).strip()

                # Only keep rows that have something in both col 0 (test name)
                # AND col 1 (result value), and where col 1 contains a digit.
                # This filters out the page header, column label row, footer etc.
                if slots[0] and slots[1] and re.search(r"\d", slots[1]):
                    word_rows.append(slots)
                elif slots[0] and len(slots[0]) < 40 and any(slots[1:]):
                    # A row that looks like a test but had no number where one
                    # was expected. Kept so that a file laid out differently
                    # can show what LabScribe nearly recognised, instead of
                    # leaving nothing behind to look at.
                    if len(diagnostics["near_misses"]) < 12:
                        diagnostics["near_misses"].append(
                            " | ".join(slot for slot in slots if slot))

    diagnostics["text_characters"] = len(header_text.strip())
    diagnostics["table_rows"] = len(table_rows)
    diagnostics["word_rows"] = len(word_rows)

    return {
        "header_text": header_text,
        "table_rows":  table_rows,
        "word_rows":   word_rows,
        "diagnostics": diagnostics,
    }


# ==============================================================================
# SECTION 13B — parse_pdf_data() FUNCTION
#
# Takes the extracted data from extract_text_from_pdf() and turns it into
# two clean Python structures:
#   1. patient  — a dictionary with patient details (name, age, gender etc.)
#   2. results  — a list of test result dictionaries
#
# PARSER PRIORITY (chosen automatically):
#   1. parse_bs240_table()    — when bordered table rows were found (Layout A)
#   2. parse_bs240_words()    — when positional word rows were found (Layout B)
#   3. parse_generic_fallback() — last resort for unknown formats
# ==============================================================================

def parse_pdf_data(extracted):
    """
    Master parser. Accepts the dict returned by extract_text_from_pdf().
    Returns: (patient_dict, list_of_result_dicts)

    patient dict keys:
      name, age, gender, dob, id, req, physician,
      collected, received, reported, status, fasting,
      sample_type, blood_type, bed_no, priority, diagnosis

    result dict keys:
      test, full_name, result, unit, flag, remark, ref
    """


    # ── Unpack ────────────────────────────────────────────────────────────────
    if isinstance(extracted, dict):
        header_text = extracted.get("header_text", "")
        table_rows  = extracted.get("table_rows",  [])
        word_rows   = extracted.get("word_rows",   [])
        diagnostics = extracted.get("diagnostics") or {}
    else:
        # Backwards compatibility: plain string passed in
        header_text = extracted
        table_rows  = []
        word_rows   = []
        diagnostics = {}

    diagnostics.setdefault("notes", [])
    diagnostics.setdefault("unknown_codes", [])
    diagnostics["raw_text"] = header_text

    # ── Helper: find a regex match in header_text ─────────────────────────────
    def find(pattern, default="N/A"):
        m = re.search(pattern, header_text, re.IGNORECASE)
        return m.group(1).strip() if m else default

    # ==========================================================================
    # PART A — PATIENT INFORMATION
    #
    # This parser handles BOTH BS240 header formats:
    #
    # FORMAT 1 (Old / Weija Gbawe style):
    #   Patient: Enoch Sipitey che 320   Gender: Male   Age: 65Year   Sample ID: 26
    #   Sample: Serum    MRN.:    Zone:    Bed No.:
    #   Sent from:   Priority: No
    #   Test Date:  12/18/2023 3:36:41 PM
    #
    # FORMAT 2 (New / Yaa Specimen style):
    #   Patient: YAA SPECIMEN          Sample ID: CHE900
    #   Patient ID: CHE900               Sample Type: Serum
    #   Date of Birth:                   Bar Code: CHE900
    #   Age: 68 Year                     Collection Date: 06/23/2025
    #   Gender: Female                   Collection Time:
    #   Physician:                       Department:
    #   Diagnosis:                       Comment:
    #   Ordering Date/Time: 06/23/2025   Print Date/Time: 06/23/2025 04:18:58 PM
    # ==========================================================================

    patient = {

        # Patient name — stops at 2+ spaces OR known next-column label
        # Handles both "Patient:" and "Patient ID:" (we want the first one)
        "name": find(
            r"Patient\s*:\s*([A-Za-z][A-Za-z0-9 '\-\.]+?)(?:\s{2,}|Sample\s*ID|Patient\s*ID|\n)",
            "Unknown"
        ),

        # Age — digits followed by optional "Year"/"Yr"
        "age": find(r"Age\s*:\s*(\d+\s*(?:Year|Yr|Y)?)", "N/A"),

        # Gender
        "gender": find(r"Gender\s*:\s*(Male|Female|Other|\w+)", "N/A"),

        # Date of Birth — Format 2 has "Date of Birth:" field
        # Matches dates like 01/15/1980 or 1980-01-15
        "dob": find(
            r"Date\s*of\s*Birth\s*:\s*([\d]{1,2}[\/\-][\d]{1,2}[\/\-][\d]{2,4}|[\d]{4}[\/\-][\d]{2}[\/\-][\d]{2})",
            "N/A"
        ),

        # Sample ID — both formats use "Sample ID:"
        "id": find(r"Sample\s*ID\s*:\s*(\S+)", "N/A"),

        # Patient ID — Format 2 has a separate "Patient ID:" field
        # We use this as req/requisition number if MRN is absent
        "req": find(r"Patient\s*ID\s*:\s*(\S+)",
               find(r"MRN\.?\s*:\s*((?!Zone|Bed|Sample|Sender|\s)\S+)", "N/A")),

        # Sample type — Format 1: "Sample: Serum", Format 2: "Sample Type: Serum"
        "sample_type": find(r"Sample\s*(?:Type)?\s*:\s*(\w+)", "N/A"),

        # Physician — Format 2 has "Physician:" field
        "physician": find(r"Physician\s*:\s*([A-Za-z][A-Za-z ,\.]+?)(?:\s{2,}|\n|Department)", "N/A"),

        # Diagnosis — Format 2 has "Diagnosis:" field
        "diagnosis": find(r"Diagnosis\s*:\s*([^\n]+?)(?:\s{2,}|Comment\s*:|\s*$)", "N/A"),

        # Blood type — only if explicitly filled in
        "blood_type": find(r"Blood\s*(?:Type|Group)\s*:\s*([ABO][B]?[+\-]?(?:pos|neg|itive|ative)?)", "N/A"),

        # Bed number (old format only)
        "bed_no": find(r"Bed\s*No\.?\s*:\s*((?!Sender|Sent|\s)\S+)", "N/A"),

        # Priority (old format only)
        "priority": find(r"Priority\s*:\s*(\S+)", "N/A"),

        # Collection date —
        #   Format 1: "Test Date: 12/18/2023 3:36:41 PM"
        #   Format 2: "Collection Date: 06/23/2025" + "Ordering Date/Time: 06/23/2025"
        "collected": find(
            r"(?:Collection\s*Date|Ordering\s*Date(?:/Time)?|Test\s*Date)\s*[:/]\s*([\d/]+(?:\s+[\d:]+\s*(?:AM|PM)?)?)",
            "N/A"
        ),

        # Received date —
        #   Format 1: "Send Date: 12/18/2023 3:36:41 PM"
        #   Format 2: "Test Date/Time: 06/23/2025"
        "received": find(
            r"(?:Test\s*Date(?:/Time)?|Send\s*Date)\s*[:/]\s*([\d/]+(?:\s+[\d:]+\s*(?:AM|PM)?)?)",
            "N/A"
        ),

        # Reported / Print date
        "reported": find(
            r"Print\s*Date(?:/Time)?\s*[:/]\s*([\d/]+\s+[\d:]+\s*(?:AM|PM)?)",
            "N/A"
        ),

        # Fasting status — not present in either BS240 format
        "fasting": "N/A",

        # Always Final for a completed BS240 printout
        "status": "Final",
    }

    # ── Clean up age (remove "Year"/"Yr" suffix) ──────────────────────────────
    age_raw   = patient["age"]
    age_clean = re.sub(r"[^\d]", "", age_raw)
    patient["age"] = age_clean if age_clean else age_raw

    # ── Clean up name (strip trailing spaces) ─────────────────────────────────
    patient["name"] = patient["name"].strip()

    # ── Clean up physician — "N/A" if only whitespace was captured ───────────
    if patient["physician"].strip() in ("", "N/A"):
        patient["physician"] = "N/A"

    # ==========================================================================
    # PART B — CHOOSE RESULTS PARSER
    #
    # Priority:
    #   1. Bordered-table rows  (Layout A — old Weija Gbawe style)
    #   2. Positional word rows (Layout B — new Yaa Specimen / CHE900 style)
    #   3. Generic text fallback (unknown formats)
    # ==========================================================================

    # Count usable table rows (must have a number in the Concentration column)
    data_row_count = sum(
        1 for row in table_rows
        if len(row) >= 4 and re.search(r"\d", row[2] if len(row) > 2 else "")
    )

    # Count usable word rows (must have a number in the Result slot)
    word_row_count = sum(
        1 for row in word_rows
        if len(row) >= 2 and re.search(r"\d", row[1])
    )

    if data_row_count >= 1:
        print(f"   Using BS240 table parser  ({data_row_count} bordered-table rows)")
        diagnostics["route"] = "the bordered-table reader"
        results = parse_bs240_table(table_rows)

    elif word_row_count >= 1:
        print(f"   Using BS240 word parser   ({word_row_count} positional word rows)")
        diagnostics["route"] = "the word-position reader"
        results = parse_bs240_words(word_rows)

    else:
        print("   No structured data found, using the plain text parser instead")
        diagnostics["route"] = "the plain-text reader"
        results = parse_generic_fallback(header_text)

    # ==========================================================================
    # PART C — A NOTE ON NAMES AND REFERENCE INTERVALS
    #
    # The parser deliberately leaves the analyser's own short codes alone
    # ("T-bil", "γ-GT", "CREA"). They are translated into the laboratory's
    # test names — and given the laboratory's units and reference intervals —
    # further along, by apply_test_library() in SECTION 7.
    #
    # Keeping the two jobs apart means the reference intervals live in exactly
    # one place (TEST_LIBRARY), which is where you go to change them.
    # ==========================================================================

    # ==========================================================================
    # PART D — CALCULATE eGFR (CKD-EPI 2021 formula)
    #
    # eGFR (estimated Glomerular Filtration Rate) measures how well the kidneys
    # are filtering waste from the blood. It is NOT printed on the BS240 report
    # because the analyser only measures Creatinine — eGFR must be calculated.
    #
    # FORMULA: CKD-EPI 2021 (race-free version, the current clinical standard)
    #
    #   eGFR = 142 × min(Scr/κ, 1)^α × max(Scr/κ, 1)^(-1.200)
    #          × 0.9938^Age × (1.012 if Female)
    #
    #   Where:
    #     Scr = serum creatinine in mg/dL
    #     κ   = 0.7 (Female) or 0.9 (Male)
    #     α   = -0.241 (Female) or -0.302 (Male)
    #     Age = patient age in years
    #
    # UNIT CONVERSION:
    #   The BS240 reports creatinine in µmol/L.
    #   To convert to mg/dL: divide by 88.42
    #
    # REFERENCE RANGES (KDIGO staging):
    #   >=90         Normal or high (G1)
    #   60-89        Mildly decreased (G2)
    #   45-59        Mildly to moderately decreased (G3a)
    #   30-44        Moderately to severely decreased (G3b)
    #   15-29        Severely decreased (G4)
    #   <15          Kidney failure (G5)
    #   We report simply: >=60 as the lower acceptable bound.
    # ==========================================================================

    results = calculate_egfr(results, patient)

    # ==========================================================================
    # PART E — TRANSLATE THE HEADER FIELDS
    #
    # The analyser labels its header fields its own way (Patient, Sample ID,
    # Gender …). build_template_patient() turns them into the six lines the
    # report prints: PATH. NO, SPECIMEN DATE, PATIENT NAME, SEX, AGE, SPECIMEN.
    # ==========================================================================

    patient = build_template_patient(patient, results)

    # ── Which codes on this file LabScribe does not know ──────────────────────
    # An unrecognised code is not an error: the result still prints, under
    # OTHER INVESTIGATIONS. But it is worth seeing, because it is the one thing
    # that turns into a line in TEST_CODE_MAP and a test in its proper section.
    diagnostics["results_found"] = len(results)
    unknown = []
    for row in results:
        raw = str(row.get("test", "")).strip()
        if not raw or is_hba1c_row(row):
            continue
        name, panel = resolve_test_name(raw)
        if panel == OTHER and raw not in unknown:
            unknown.append(raw)
    diagnostics["unknown_codes"] = unknown

    return patient, results


# ==============================================================================
# HELPER — calculate_egfr()
#
# eGFR (estimated Glomerular Filtration Rate) says how well the kidneys are
# filtering waste out of the blood. The BS240 does not print it, because the
# analyser only measures creatinine — the rest is arithmetic, done here.
#
# FORMULA — CKD-EPI 2021, the race-free version that is the current standard:
#
#   eGFR = 142 × min(Scr/κ, 1)^α × max(Scr/κ, 1)^(-1.200)
#          × 0.9938^Age × (1.012 if female)
#
#   Scr = serum creatinine in mg/dL   (the BS240 reports µmol/L: divide by 88.42)
#   κ   = 0.7 female, 0.9 male
#   α   = -0.241 female, -0.302 male
#
# The result is inserted into the list directly after creatinine, so it prints
# at the foot of the renal section as the template shows.
# ==============================================================================

CREATININE_CODES = {
    "crea", "creat", "creatinine", "cre",
}


def calculate_egfr(results, patient):
    """
    Adds an eGFR row to the results, worked out from the creatinine result.

    Returns the list unchanged if creatinine is missing, if the age is unknown,
    or if the analyser already reported an eGFR of its own.
    """
    results = list(results)

    # ── Do we already have one? ───────────────────────────────────────────────
    for r in results:
        name, _ = resolve_test_name(r.get("test", ""))
        if name == "eGFR (CKD-EPI)":
            return results

    # ── Find creatinine ───────────────────────────────────────────────────────
    crea_index = None
    crea_value = None
    crea_unit = ""
    for i, r in enumerate(results):
        code = str(r.get("test", "")).strip().lower()
        if " — " in code:
            code = code.split(" — ")[0].strip()
        if code.strip(" .:") in CREATININE_CODES:
            crea_index = i
            crea_value = r.get("result", "")
            crea_unit = str(r.get("unit", ""))
            break

    if crea_index is None:
        return results

    scr = _to_number(crea_value)
    if scr is None or scr <= 0:
        return results

    # ── Age ───────────────────────────────────────────────────────────────────
    age_text = re.sub(r"[^\d]", "", str(patient.get("age", "")))
    if not age_text:
        return results
    age = int(age_text)
    if age <= 0:
        return results

    # ── Convert creatinine to mg/dL if it came in µmol/L ──────────────────────
    unit_lower = crea_unit.lower()
    looks_like_micromol = ("mol" in unit_lower) or (not unit_lower and scr > 30)
    if looks_like_micromol:
        scr = scr / 88.42

    # ── Sex-specific constants ────────────────────────────────────────────────
    sex = str(patient.get("gender") or patient.get("sex") or "").strip().lower()
    is_female = sex.startswith("f")

    if is_female:
        kappa, alpha, sex_factor = 0.7, -0.241, 1.012
    else:
        kappa, alpha, sex_factor = 0.9, -0.302, 1.0

    ratio = scr / kappa
    egfr = (142
            * (min(ratio, 1.0) ** alpha)
            * (max(ratio, 1.0) ** -1.200)
            * (0.9938 ** age)
            * sex_factor)

    # The laboratory reports eGFR as a whole number
    egfr_value = int(round(egfr))

    results.insert(crea_index + 1, {
        "test":   "eGFR (CKD-EPI)",
        "result": str(egfr_value),
        "unit":   "ml/min/1.73m<super>2</super>",
        "flag":   "",
        "ref":    "",
        # Marks this row as LabScribe's own arithmetic rather than something
        # the analyser printed. It lets the figure be thrown away and worked
        # out again when the age or the sex in the window is corrected — see
        # recompute_egfr().
        "computed": True,
    })

    return results


# ==============================================================================
# HELPER — parse_bs240_table()
#
# Used for Layout A PDFs that have real table borders.
# Columns arrive already separated by pdfplumber's extract_tables():
#   [0] Test code  [1] Full Name  [2] Concentration  [3] Unit
#   [4] Results (Normal/high/low)  [5] Remark (RFH/RFL)  [6] Reference
# ==============================================================================

def parse_bs240_table(table_rows):
    """
    Parses bordered-table rows (Layout A — old Weija Gbawe format).
    Returns a list of result dicts.
    """
    results = []

    HEADER_WORDS = {
        "test", "full name", "concentration", "unit", "results",
        "remark", "reference", "result", "fullname", "chemistry",
    }

    for row in table_rows:
        if len(row) < 4:
            continue

        test_code  = row[0] if len(row) > 0 else ""
        full_name  = row[1] if len(row) > 1 else ""
        conc       = row[2] if len(row) > 2 else ""
        unit       = row[3] if len(row) > 3 else ""
        result_txt = row[4] if len(row) > 4 else ""
        remark     = row[5] if len(row) > 5 else ""
        ref        = row[6] if len(row) > 6 else ""

        if test_code.lower().strip() in HEADER_WORDS:
            continue

        conc_clean = conc.strip()
        if not re.search(r"\d", conc_clean):
            continue

        result_lower = result_txt.strip().lower()
        if result_lower == "high" or remark.upper() in ("RFH",):
            flag = "High"
        elif result_lower == "low" or remark.upper() in ("RFL",):
            flag = "Low"
        else:
            flag = ""

        if full_name and full_name.strip() and full_name.strip() != test_code.strip():
            display_name = f"{test_code.strip()} — {full_name.strip()}"
        else:
            display_name = test_code.strip()

        results.append({
            "test":      display_name,
            "full_name": full_name.strip(),
            "result":    conc_clean,
            "unit":      unit.strip(),
            "flag":      flag,
            "remark":    remark.strip(),
            "ref":       ref.strip().replace("–", "-").replace("—", "-"),
        })

    return results


# ==============================================================================
# HELPER — parse_bs240_words()
#
# Used for Layout B PDFs — no table borders, columns identified by x-position.
# Each row in word_rows is already a 5-slot list built in extract_text_from_pdf:
#   [0] test_name  [1] result  [2] unit  [3] flag (H/L/"")  [4] ref_range
#
# This format comes from the "Yaa Specimen / CHE900" style PDF where:
#   - Flag column contains just "H" or "L" (not "Normal"/"high"/"low")
#   - Some rows have no unit and no ref range (e.g. AST/ALT ratio, CORONARY R)
#   - Test names can contain special chars: γ-GT, K+, Cl-, Na+, T-bil
# ==============================================================================

def parse_bs240_words(word_rows):
    """
    Parses positional word rows (Layout B — new CHE900 / Yaa Specimen format).
    Returns a list of result dicts.
    """
    results = []

    # Words that identify header rows or footer rows — skip these
    SKIP_NAMES = {
        "chemistry", "result", "unit", "flag", "ref", "range",
        "ordering", "operator", "reviewer", "page", "the",
        "print", "test", "collection", "patient", "gender",
        "age", "physician", "diagnosis", "sample", "date",
        "ordering date/time:", "print date/time:", "test date/time:",
        "send date/time:", "the result", "the results",
    }

    # Also skip any row where the test name contains "date" or "time" or "page"
    # These are footer/header lines, not test results
    DATE_TIME_SKIP = re.compile(r"(?:date|time|page|operator|reviewer|result\s+are)", re.IGNORECASE)

    for row in word_rows:
        if len(row) < 2:
            continue

        test_name  = row[0].strip()
        result_val = row[1].strip()
        unit       = row[2].strip() if len(row) > 2 else ""
        flag_raw   = row[3].strip() if len(row) > 3 else ""
        ref        = row[4].strip() if len(row) > 4 else ""

        # Skip header/footer rows — by exact name match
        if test_name.lower() in SKIP_NAMES:
            continue
        # Skip any row whose test name looks like a date/time label or footer line
        if DATE_TIME_SKIP.search(test_name):
            continue

        # Skip if result has no digit
        if not re.search(r"\d", result_val):
            continue

        # Normalise flag:
        #   Layout B uses "H" or "L" in the Flag column
        #   Layout A uses "Normal"/"high"/"low" in the Results column
        flag_upper = flag_raw.upper()
        if flag_upper in ("H", "HIGH", "RFH"):
            flag = "High"
        elif flag_upper in ("L", "LOW", "RFL"):
            flag = "Low"
        else:
            flag = ""

        results.append({
            "test":      test_name,
            "full_name": "",       # Layout B doesn't have a separate Full Name column
            "result":    result_val,
            "unit":      unit,
            "flag":      flag,
            "remark":    flag_raw,  # Store original H/L for reference
            "ref":       ref.replace("–", "-"),
        })

    return results


# ==============================================================================
# HELPER — parse_generic_fallback()
#
# Last-resort parser used when neither table extraction nor positional word
# grouping produces results. Scans every line of plain text with regex.
# Less accurate but works for unknown analyser formats and simple text PDFs.
# ==============================================================================

def parse_generic_fallback(text):
    """
    Regex line-scanner fallback. Returns result dicts in the same format
    as parse_bs240_table() and parse_bs240_words().
    """
    results = []

    SKIP_WORDS = {
        "test", "full name", "concentration", "unit", "results", "chemistry",
        "remark", "reference", "patient", "sample", "gender", "flag", "ref",
        "sender", "sent", "diagnosis", "test date", "send date", "print date",
        "ordering", "operator", "reviewer", "collection", "physician",
    }

    # Pattern: TEST_NAME   NUMBER   UNIT   (H|L|flag)   (ref_range)
    pattern = re.compile(
        r"^([A-Za-zγµ][A-Za-z0-9 /\-\(\)\.\+γµ]+?)\s+"
        r"([<>]?\d+\.?\d*)\s*"
        r"([µμa-zA-Z/%\.]+(?:/[a-zA-Z\.]+)?)?\s*"
        r"(Normal|high|low|H\b|L\b|RFH|RFL|\*)?\s*"
        r"([<>=]{0,2}[\d\.\-–]+(?:\s*[-–]\s*[\d\.]+)?)?",
        re.MULTILINE | re.IGNORECASE
    )

    for m in pattern.finditer(text):
        test_name  = m.group(1).strip()
        result_val = m.group(2).strip()
        unit       = m.group(3).strip() if m.group(3) else ""
        flag_word  = m.group(4).strip() if m.group(4) else ""
        ref        = m.group(5).strip() if m.group(5) else ""

        if test_name.lower() in SKIP_WORDS or len(test_name) < 2:
            continue

        flag_upper = flag_word.upper()
        if flag_upper in ("HIGH", "H", "RFH"):
            flag = "High"
        elif flag_upper in ("LOW", "L", "RFL"):
            flag = "Low"
        else:
            flag = ""

        results.append({
            "test":      test_name,
            "full_name": "",
            "result":    result_val,
            "unit":      unit,
            "flag":      flag,
            "remark":    flag_word,
            "ref":       ref.replace("–", "-"),
        })

    return results


# ==============================================================================
# SECTION 14 — main()  —  running the script from a terminal
#
# Double-clicking the file opens the window instead; this is here for batch
# work, for testing, and for anyone who prefers the command line.
# ==============================================================================

def read_pdf(input_pdf, overrides=None):
    """
    Reads one analyser PDF and returns (patient, results, diagnostics).

    The diagnostics are everything LabScribe noticed on the way through — how
    many pages, whether there was any text at all, which of the three readers
    ran, what it produced, which codes were not recognised. They are what makes
    the difference between "could not be read" and a file that says what is
    wrong with it.

    `overrides` is an optional dict of header fields typed by the user, which
    always win over whatever was parsed out of the PDF.
    """
    extracted = extract_text_from_pdf(input_pdf)
    parsed, results = parse_pdf_data(extracted)
    patient = build_template_patient(parsed, results)

    for key, value in (overrides or {}).items():
        if str(value).strip():
            patient[key] = str(value).strip()

    return patient, results, (extracted.get("diagnostics") or {})


def process_pdf(input_pdf, overrides=None):
    """The same thing without the diagnostics, for anything that only wants
    the two."""
    patient, results, _diagnostics = read_pdf(input_pdf, overrides)
    return patient, results


def explain_reading(diagnostics, error=""):
    """
    Turns the diagnostics into something a person can act on.

    Returns (headline, evidence, suggestions) — a single line saying what
    happened, the facts behind it, and what to try. Written for whoever is
    standing at the bench with a PDF that will not open, not for a programmer.
    """
    diagnostics = diagnostics or {}
    opened = "pages" in diagnostics
    pages = diagnostics.get("pages", 0)
    characters = diagnostics.get("text_characters", 0)
    images = diagnostics.get("images", 0)
    found = diagnostics.get("results_found", 0)
    lowered = str(error).lower()

    evidence, suggestions = [], []

    # ── The file itself would not open ────────────────────────────────────────
    if "password" in lowered or "encrypt" in lowered:
        return ("This PDF is locked with a password.",
                ["LabScribe cannot open it without the password."],
                ["Save an unlocked copy from whatever locked it, then add "
                 "that copy instead."])

    if not opened:
        return ("This file could not be opened at all.",
                [f"What went wrong: {error}" if error else
                 "It is damaged, or it is not really a PDF."],
                ["Export the report from the analyser again.",
                 "Check the file opens in a PDF reader before adding it here."])

    if pages == 0:
        return ("There are no pages in this PDF.",
                ["The file opened, but there is nothing inside it."],
                ["Export the report from the analyser again — the last export "
                 "probably did not finish."])

    evidence.append(f"{pages} page" + ("" if pages == 1 else "s") + " in the file.")

    # ── A scan rather than the analyser's own export ──────────────────────────
    if characters == 0:
        if images:
            return ("This is a picture of a report, not the analyser's printout.",
                    evidence + [f"No text at all — only {images} image"
                                + ("" if images == 1 else "s") + ".",
                                "That is what a scan or a photograph looks like."],
                    ["Export the report from the analyser again, or print it "
                     "to PDF.",
                     "A scanned paper copy cannot be read: the numbers are "
                     "pixels, not text."])
        return ("There is no text in this PDF.",
                evidence + ["Nothing could be extracted from any page."],
                ["Export the report from the analyser again.",
                 "Check the file opens and shows text in a PDF reader."])

    evidence.append(f"{characters} characters of text were read off it.")

    route = diagnostics.get("route")
    if route:
        evidence.append(f"LabScribe tried {route}.")

    table_rows = diagnostics.get("table_rows", 0)
    word_rows = diagnostics.get("word_rows", 0)
    if table_rows or word_rows:
        evidence.append(f"{table_rows} bordered-table rows and {word_rows} "
                        f"rows found by word position.")

    # ── Text there, but nothing that looks like a result ──────────────────────
    if found == 0:
        near = diagnostics.get("near_misses") or []
        if near:
            evidence.append("Rows that looked like tests but had no number "
                            "beside them:")
            evidence.extend(f"    {line}" for line in near[:6])
        return ("LabScribe found text on this file, but nothing it recognised "
                "as a result.",
                evidence,
                ["Check this is the analyser printout and not a request form, "
                 "a worksheet or a covering page.",
                 "If it is the printout, the columns are not where LabScribe "
                 "expects them. Use \"Show the text LabScribe found\" below "
                 "and send it on — the layout can be added.",
                 "Try another file from the same run to see whether it is this "
                 "one file or the whole layout."])

    # ── It read, and the caller still called it a failure ─────────────────────
    evidence.append(f"{found} results were read.")
    return ("LabScribe read this file, but something went wrong afterwards.",
            evidence + ([f"What went wrong: {error}"] if error else []),
            ["Use \"Show the text LabScribe found\" below and send it on."])


def default_output_name(patient, input_pdf=None):
    """Builds a sensible file name, e.g. 'CHE901_26_ADWOA_SAMPLE_BOATENG.pdf'."""
    path_no = re.sub(r"[^A-Za-z0-9]+", "_", patient.get("path_no", "")).strip("_")
    name = re.sub(r"[^A-Za-z0-9]+", "_", patient.get("name", "")).strip("_")
    stem = "_".join(part for part in (path_no, name) if part)
    if not stem:
        stem = os.path.splitext(os.path.basename(input_pdf or "report"))[0]
    return f"{stem}.pdf"


def main():
    parser = argparse.ArgumentParser(
        description="Report generator for Ngleshie Amanfrom Polyclinic Laboratory",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Sample renal / liver / lipid report:
    python lab_report.py --demo

  Sample glycated haemoglobin report:
    python lab_report.py --demo-hba1c

  A real analyser PDF:
    python lab_report.py "CHE901.pdf"

  Correcting the header by hand:
    python lab_report.py "CHE901.pdf" --path-no CHE901/26 --sex FEMALE
        --age 46 --specimen SERUM --specimen-date 27/04/2026
        --output "CHE901 ADWOA SAMPLE BOATENG.pdf"
        """,
    )

    parser.add_argument("input_pdf", nargs="*",
                        help="The PDF(s) exported from the analyser. Give two "
                             "and, if they carry the same patient, they are "
                             "reported together on one page.")

    # ── Header fields — each one overrides what was read from the PDF ─────────
    parser.add_argument("--path-no",       default=None, help="e.g. CHE901/26")
    parser.add_argument("--specimen-date", default=None, help="e.g. 27/04/2026")
    parser.add_argument("--name",          default=None, help="Patient name")
    parser.add_argument("--sex",           default=None, help="MALE or FEMALE")
    parser.add_argument("--age",           default=None, help="Age in years")
    parser.add_argument("--specimen",      default=None,
                        help="e.g. SERUM or WHOLE BLOOD")

    parser.add_argument("--lab-name",  default=None, help="Laboratory name in the header")
    parser.add_argument("--signatory", default=None, help="Footer signature line")
    parser.add_argument("--output",    default=None, help="Output PDF file name")
    parser.add_argument("--flags", action="store_true",
                        help="Work out High/Low and print it in the FLAG column")

    parser.add_argument("--demo",       action="store_true",
                        help="Sample renal / liver / lipid report")
    parser.add_argument("--demo-hba1c", action="store_true",
                        help="Sample glycated haemoglobin report")

    args = parser.parse_args()

    global LAB_NAME, SIGNATORY, SHOW_FLAGS
    if args.lab_name:
        LAB_NAME = args.lab_name
    if args.signatory:
        SIGNATORY = args.signatory
    if args.flags:
        SHOW_FLAGS = True

    # ── Demo modes ────────────────────────────────────────────────────────────
    if args.demo:
        out = args.output or "demo_report.pdf"
        generate_report(DEMO_PATIENT, DEMO_RESULTS, out)
        return

    if args.demo_hba1c:
        out = args.output or "demo_hba1c_report.pdf"
        generate_report(DEMO_HBA1C_PATIENT, DEMO_HBA1C_RESULTS, out)
        return

    # ── A real analyser PDF ───────────────────────────────────────────────────
    if args.input_pdf:
        overrides = {
            "path_no":       args.path_no,
            "specimen_date": args.specimen_date,
            "name":          (args.name or "").upper() or None,
            "sex":           (args.sex or "").upper() or None,
            "age":           args.age,
            "specimen":      (args.specimen or "").upper() or None,
        }
        overrides = {k: v for k, v in overrides.items() if v}

        jobs = []
        for input_pdf in args.input_pdf:
            print(f"Reading: {input_pdf}")
            try:
                patient, results = process_pdf(input_pdf, overrides)
            except FileNotFoundError:
                print(f"File not found: {input_pdf}")
                print("   Check the path. On Windows, put quotes around paths "
                      "with spaces.")
                continue
            except Exception as exc:
                print(f"Could not read the PDF: {exc}")
                continue

            if not results:
                print("No test results could be read from this PDF.")
                print("   The layout may not match what the parser expects.")
                print("   Try --demo first to confirm the rest of the script works.")
                continue

            print(f"Found {len(results)} results for "
                  f"{patient.get('name') or 'this patient'}")
            jobs.append({"path": input_pdf, "patient": patient,
                         "results": results})

        if not jobs:
            return

        # Files carrying the same patient are printed as one report.
        taken = set()
        for group in group_jobs_by_patient(jobs):
            patient, results = merged_job(group)

            if len(group) > 1:
                print(f"   {len(group)} files carry "
                      f"{patient.get('name') or 'the same patient'}, so they "
                      f"are being reported together.")

            if args.output and len(jobs) == 1:
                out = args.output
            else:
                out = unique_output_path(
                    os.path.dirname(os.path.abspath(group[0]["path"])),
                    default_output_name(patient, group[0]["path"]),
                    taken)

            generate_report(patient, results, out)
        return

    parser.print_help()


# ==============================================================================
# SECTION 15 - THE LABSCRIBE WINDOW
#
# The desktop application. One window, three areas:
#
#   Left    the list of analyser files you have added, each with its status
#   Right   the patient details and the results read from whichever file is
#           selected, so the reading can be checked before anything prints
#   Bottom  where the reports are saved, and the Generate button
#
# Files are read in the background the moment they are added, so the window
# never freezes and you can see straight away whether a file was understood.
# ==============================================================================

# ── HOW THE WINDOW LOOKS ──────────────────────────────────────────────────────
# Six styles to choose from, under Settings ▸ Appearance. They are not just
# different colours: the lettering changes with them, and the top of the window
# is either a solid coloured band or a quiet white strip with a thin rule.
#
# Whichever one is picked is remembered, so the program opens that way next
# time. Every colour the window uses is listed here, and nowhere else, so a new
# style is added simply by writing another block in this list.
#
# The keys, in plain words:
#   bg          the window behind everything
#   panel       the cards and lists that sit on it
#   line        hairline borders;   line_soft  fainter separators
#   ink         ordinary text;      ink_soft   labels;   ink_faint  hints
#   accent      the one strong colour: header band, buttons, headings
#   accent_dark the same colour pressed;   accent_tint  a selected row
#   accent_soft lettering that sits on top of the accent, second rank
#   on_accent   lettering that sits on top of the accent, first rank
#   danger      results outside the interval, and things that went wrong
#   warn        things worth a second look
#   field       inside a text box;  field_line  its border
#   drop        the "drop files here" panel;  row_alt  every other table row
#   off         a button that cannot be pressed yet
#   header      "solid" for a coloured band, "quiet" for white with a rule
#   fonts       first one of these the computer has is used

STYLES = [
    {
        "id": "green",
        "name": "Ghana Green",
        "note": "The Ghana Health Service green, so the window and the "
                "reports it prints look like they belong together.",
        "bg": "#eef1f4", "panel": "#ffffff",
        "line": "#d4dae0", "line_soft": "#e7ebef",
        "ink": "#1c2530", "ink_soft": "#5d6b7a", "ink_faint": "#7d8a98",
        "accent": "#0f7b3e", "accent_dark": "#0a5c2d",
        "accent_tint": "#e8f3ec", "accent_soft": "#bfe0cd",
        "on_accent": "#ffffff",
        "danger": "#c62828", "warn": "#b26a00",
        "field": "#ffffff", "field_line": "#c3ccd5",
        "drop": "#f7f9fa", "row_alt": "#fafbfc", "off": "#a9b6bf",
        "header": "solid",
        "fonts": ["Segoe UI", "Ubuntu", "DejaVu Sans", "Helvetica"],
    },
    {
        "id": "paper",
        "name": "Ink on Paper",
        "note": "Off white, black rules and a serif face, the way a bench "
                "worksheet reads. Red is kept for results out of range.",
        "bg": "#f2efe9", "panel": "#fbf9f5",
        "line": "#c9c3b6", "line_soft": "#e2ddd2",
        "ink": "#1a1a18", "ink_soft": "#57534a", "ink_faint": "#8b8478",
        "accent": "#23221f", "accent_dark": "#000000",
        "accent_tint": "#e8e3d8", "accent_soft": "#c8c2b4",
        "on_accent": "#faf8f3",
        "danger": "#a8261f", "warn": "#8a5a12",
        "field": "#ffffff", "field_line": "#bdb6a8",
        "drop": "#f0ece3", "row_alt": "#f5f2eb", "off": "#b3aca0",
        "header": "solid",
        "fonts": ["Georgia", "Cambria", "DejaVu Serif", "Times New Roman"],
    },
    {
        "id": "midnight",
        "name": "Midnight",
        "note": "A dark window for a dark bench, or for late runs when a "
                "white screen is too much.",
        "bg": "#161b22", "panel": "#1e242d",
        "line": "#2f3742", "line_soft": "#262d36",
        "ink": "#e6ebf1", "ink_soft": "#9fb0c0", "ink_faint": "#6f8093",
        "accent": "#2f9e6e", "accent_dark": "#25805a",
        "accent_tint": "#223a31", "accent_soft": "#b9e4d0",
        "on_accent": "#08110c",
        "danger": "#ff6b63", "warn": "#e0a04a",
        "field": "#141a21", "field_line": "#39424e",
        "drop": "#1a2029", "row_alt": "#222933", "off": "#3a444f",
        "header": "solid",
        "fonts": ["Segoe UI", "Ubuntu", "DejaVu Sans", "Helvetica"],
    },
    {
        "id": "oxford",
        "name": "Oxford Blue",
        "note": "Navy and cool grey. Sober, and easy to read for a long "
                "stretch at the machine.",
        "bg": "#eef1f6", "panel": "#ffffff",
        "line": "#ccd4e0", "line_soft": "#e3e8f0",
        "ink": "#141c2b", "ink_soft": "#4f5d75", "ink_faint": "#7b8aa2",
        "accent": "#173f74", "accent_dark": "#0f2c53",
        "accent_tint": "#e2eaf6", "accent_soft": "#a9c2e2",
        "on_accent": "#ffffff",
        "danger": "#b3261e", "warn": "#9a6100",
        "field": "#ffffff", "field_line": "#bcc6d6",
        "drop": "#f5f7fb", "row_alt": "#f8fafd", "off": "#a7b1c1",
        "header": "solid",
        "fonts": ["Segoe UI", "Ubuntu", "DejaVu Sans", "Helvetica"],
    },
    {
        "id": "clay",
        "name": "Warm Clay",
        "note": "Sand and burnt orange. Warmer than the rest without being "
                "loud about it.",
        "bg": "#f4f0ea", "panel": "#fffdfa",
        "line": "#dcd2c6", "line_soft": "#ece5db",
        "ink": "#241f1a", "ink_soft": "#6b5d4f", "ink_faint": "#968676",
        "accent": "#a4522a", "accent_dark": "#7f3d1e",
        "accent_tint": "#f6e7dc", "accent_soft": "#f0c8ae",
        "on_accent": "#fffaf6",
        "danger": "#b3261e", "warn": "#8a5a12",
        "field": "#ffffff", "field_line": "#cdbfb1",
        "drop": "#f7f2ec", "row_alt": "#faf7f2", "off": "#bbaea1",
        "header": "solid",
        "fonts": ["Segoe UI", "Ubuntu", "DejaVu Sans", "Helvetica"],
    },
    {
        "id": "graphite",
        "name": "Quiet Graphite",
        "note": "Almost no colour, and a white top strip instead of a band. "
                "The nearest thing here to a plain grey office program.",
        "bg": "#f4f5f6", "panel": "#ffffff",
        "line": "#d8dbde", "line_soft": "#eaecee",
        "ink": "#20242a", "ink_soft": "#5a616a", "ink_faint": "#858c95",
        "accent": "#3f4d5a", "accent_dark": "#2b363f",
        "accent_tint": "#e9edf1", "accent_soft": "#9aa6b2",
        "on_accent": "#ffffff",
        "danger": "#b3261e", "warn": "#8a5a12",
        "field": "#ffffff", "field_line": "#c6cbd1",
        "drop": "#f7f8f9", "row_alt": "#fafbfc", "off": "#aeb4ba",
        "header": "quiet",
        "fonts": ["Segoe UI", "Ubuntu", "DejaVu Sans", "Helvetica"],
    },
]

DEFAULT_STYLE = "green"

# UI is the style in force. Everything in the window reads its colours from
# here, so changing what is in it and redrawing is all a change of style takes.
UI = {}
UI_FONT = "Segoe UI"


# ── THE SPACING SCALE ─────────────────────────────────────────────────────────
# Every gap in the window is one of these six numbers. Before, each panel had
# been padded by whatever looked right at the time — 11 here, 14 there, 16
# somewhere else — and the eye reads that unevenness as carelessness even when
# it cannot say why. One scale, used everywhere, is what makes a window look
# deliberate.
SPACE = {
    "xs":  4,
    "sm":  8,
    "md":  12,
    "lg":  18,
    "xl":  26,
    "xxl": 36,
}


# ── THE TYPE SCALE ────────────────────────────────────────────────────────────
# Size and weight for every piece of lettering in the window, named by the job
# it does rather than by how big it is. Anything not on this list does not
# belong on the screen.
TYPE = {
    "display":   (15, "bold"),     # The program's name, top left
    "title":     (12, "bold"),     # Dialog headings, the headline on an error
    "subtitle":  (10, "bold"),     # The laboratory name in the header
    "label":     (8,  "bold"),     # The small capitals over a card
    "body":      (10, "normal"),   # Ordinary text, and what is typed in a box
    "body_bold": (10, "bold"),     # A button that means it
    "small":     (9,  "normal"),   # Hints, the status line, table rows
    "small_bold": (9, "bold"),
    "tiny":      (8,  "normal"),   # The credit line, and nothing else
    "table":     (9,  "normal"),
}


def ui_font(key, family=None):
    """One of the type scale's faces, ready to hand to a widget."""
    size, weight = TYPE.get(key, TYPE["body"])
    family = family or UI_FONT
    return (family, size) if weight == "normal" else (family, size, weight)


def style_by_id(style_id):
    for style in STYLES:
        if style["id"] == style_id:
            return style
    return STYLES[0]


def apply_style(style_id):
    """Loads one style into UI, ready for the window to be drawn from it."""
    style = style_by_id(style_id)
    UI.clear()
    UI.update(style)
    return style


def pick_font(candidates):
    """The first of these lettering faces the computer actually has."""
    try:
        import tkinter.font as tkfont
        have = {name.lower() for name in tkfont.families()}
        for name in candidates:
            if name.lower() in have:
                return name
    except Exception:
        pass
    return candidates[-1] if candidates else "Helvetica"


apply_style(DEFAULT_STYLE)


def app_directory():
    """
    The folder the application lives in.

    When LabScribe is running as a packaged .exe this is the folder holding
    the .exe; when it is being run as a .py file it is the folder holding the
    script. Used to find an optional lab_logo.png sitting beside it.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resource_directory():
    """
    Where files built into the program live.

    A one-file .exe unpacks what was bundled with it into a temporary folder
    and points sys._MEIPASS at it. That is where the icon is found when
    LabScribe is running as an .exe; when it is running as a .py file it is
    simply the folder the script is in.
    """
    bundled = getattr(sys, "_MEIPASS", None)
    if bundled:
        return bundled
    return os.path.dirname(os.path.abspath(__file__))


def settings_directory():
    """
    Where LabScribe keeps its saved settings.

    This is the user's own AppData folder, not the program folder, so the
    settings survive replacing the .exe with a newer version, and work even
    when the program sits somewhere the user cannot write to.
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "LabScribe")
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        return os.path.expanduser("~")
    return folder


def settings_path():
    return os.path.join(settings_directory(), "settings.txt")


def read_settings():
    """
    Whatever was saved last time, as a plain dictionary.

    Read at module level as well as inside the window, because the chosen
    style has to be known before the first widget is made.
    """
    saved = {}
    try:
        with open(settings_path(), encoding="utf-8") as handle:
            for line in handle:
                if "=" in line:
                    key, _, value = line.strip().partition("=")
                    saved[key] = value
    except Exception:
        pass
    return saved


def _build_gui_app():
    """
    Creates and returns the App class.

    It is built inside a function so the class body is only evaluated once
    tkinter is known to be present. That keeps the command line mode working
    on machines with no graphical libraries at all.
    """

    BaseClass = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

    class App(BaseClass):

        # ══════════════════════════════════════════════════════════════════════
        # SET-UP
        # ══════════════════════════════════════════════════════════════════════

        def __init__(self):
            super().__init__()

            saved = read_settings()

            # The style is settled first, before a single widget is drawn,
            # because every widget takes its colours from it as it is made.
            self.style_id = saved.get("style") or DEFAULT_STYLE
            apply_style(self.style_id)

            self.title(APP_NAME)
            self.geometry("1140x790")
            self.minsize(1000, 700)
            self.configure(bg=UI["bg"])
            self._set_window_icon()

            # ── What the window is holding ────────────────────────────────────
            # jobs is the heart of it: one entry per file added, carrying the
            # patient details and the results read from that file.
            #   {"path", "name", "status", "patient", "results", "error", "out"}
            self.jobs = []
            self.current = None          # Index of the file being shown, or None
            self.busy = False            # True while reading or generating

            self.lab_name_var  = tk.StringVar(value=LAB_NAME)
            self.signatory_var = tk.StringVar(value=SIGNATORY)
            self.output_var    = tk.StringVar()
            self.flags_var     = tk.BooleanVar(value=SHOW_FLAGS)
            self._drop_hover   = False   # The pointer is over the panel
            self._drop_dragging = False  # A file is being dragged over it
            # Reports land next to the analyser file they were read from,
            # unless a folder has been chosen by hand.
            self.beside_var    = tk.BooleanVar(value=True)
            # Ticked when the selected file is to be kept out of its patient's
            # group and reported on its own.
            self.separate_var  = tk.BooleanVar(value=False)
            self.status_var    = tk.StringVar(
                value="Add the PDFs the analyser printed and LabScribe will "
                      "read them.")

            # The six header fields, one text box each
            self.field_vars = {key: tk.StringVar() for key, _ in PATIENT_FIELDS}
            for var in self.field_vars.values():
                var.trace_add("write", self._on_field_edited)
            self._loading_fields = False     # Guards against edit-loops

            self._build_styles()
            self._build_menu()
            self._build_layout()
            self._load_settings()
            self._refresh_file_list()
            self._show_job(None)
            self._restore_geometry(saved.get("geometry", ""))
            self.protocol("WM_DELETE_WINDOW", self._on_close)

        def _on_close(self):
            """The size and position are kept, so the window opens where it
            was left rather than back in the middle every morning."""
            self._save_settings()
            self.destroy()

        def _restore_geometry(self, geometry):
            """
            Puts the window back where it was last time, provided that is still
            somewhere on the screen. A window restored onto a monitor that is
            no longer plugged in is worse than one opening in the middle.
            """
            match = re.match(r"^(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$",
                             str(geometry).strip())
            if match:
                width, height, x, y = (int(part) for part in match.groups())
                on_screen = (width >= 1000 and height >= 700
                             and -20 <= x <= self.winfo_screenwidth() - 200
                             and -20 <= y <= self.winfo_screenheight() - 200)
                if on_screen:
                    self.geometry(f"{width}x{height}+{x}+{y}")
                    return
            self._centre()

        def _set_window_icon(self):
            """
            Puts the LabScribe mark in the title bar and the taskbar.

            An icon sitting beside the program wins, so the mark can be
            changed without rebuilding — the same arrangement as lab_logo.png.
            Failing that, the one built into the program is used.
            """
            for folder in (app_directory(), resource_directory()):
                try:
                    icon_path = os.path.join(folder, "labscribe.ico")
                    if os.path.isfile(icon_path):
                        self.iconbitmap(default=icon_path)
                        return
                except Exception:
                    continue
            try:
                # Fall back to the logo embedded in this file
                self._icon_image = tk.PhotoImage(data=APP_ICON_B64)
                self.iconphoto(True, self._icon_image)
            except Exception:
                pass

        def _centre(self):
            self.update_idletasks()
            x = (self.winfo_screenwidth() - self.winfo_width()) // 2
            y = (self.winfo_screenheight() - self.winfo_height()) // 3
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        # ══════════════════════════════════════════════════════════════════════
        # APPEARANCE
        # ══════════════════════════════════════════════════════════════════════

        def _build_styles(self):
            global UI_FONT
            UI_FONT = pick_font(UI["fonts"])

            style = ttk.Style(self)
            try:
                style.theme_use("clam")      # The most restyleable built-in theme
            except tk.TclError:
                pass

            style.configure(".", background=UI["bg"], foreground=UI["ink"],
                            font=ui_font("body"))

            style.configure("Card.TFrame", background=UI["panel"])
            style.configure("Bg.TFrame", background=UI["bg"])

            # ── Lettering ─────────────────────────────────────────────────────
            # Every label in the window is one of these six, so nothing is
            # sized or coloured by hand at the point of use.
            for name, key, back, ink in (
                    ("CardTitle", "label",     "panel", "ink_faint"),
                    ("Field",     "small",     "panel", "ink_soft"),
                    ("Body",      "body",      "panel", "ink"),
                    ("BodyBold",  "body_bold", "panel", "ink"),
                    ("Hint",      "small",     "panel", "ink_faint"),
                    ("Status",    "small",     "bg",    "ink_soft"),
                    ("BarTitle",  "label",     "bg",    "ink_faint"),
                    ("Credit",    "tiny",      "bg",    "ink_faint")):
                style.configure(f"{name}.TLabel", background=UI[back],
                                foreground=UI[ink], font=ui_font(key))

            # ── Boxes you type in ─────────────────────────────────────────────
            # The border is the only thing that changes on focus: a ring in the
            # accent colour, drawn where the hairline was, so nothing moves.
            style.configure("Field.TEntry", fieldbackground=UI["field"],
                            background=UI["field"], foreground=UI["ink"],
                            bordercolor=UI["field_line"],
                            lightcolor=UI["field_line"],
                            darkcolor=UI["field_line"],
                            insertcolor=UI["ink"],
                            borderwidth=1, relief="solid",
                            padding=(SPACE["sm"], SPACE["sm"] - 1))
            style.map("Field.TEntry",
                      bordercolor=[("focus", UI["accent"]),
                                   ("hover", UI["ink_faint"])],
                      lightcolor=[("focus", UI["accent"])],
                      darkcolor=[("focus", UI["accent"])],
                      fieldbackground=[("disabled", UI["bg"])],
                      foreground=[("disabled", UI["ink_faint"])])

            # ── Buttons ───────────────────────────────────────────────────────
            # One button on the screen means it — the one that prints. It is
            # the only thing wearing the accent colour as a fill, which is what
            # makes it findable without hunting.
            style.configure("Primary.TButton", background=UI["accent"],
                            foreground=UI["on_accent"], font=ui_font("body_bold"),
                            borderwidth=0, focusthickness=0,
                            padding=(SPACE["xl"], SPACE["md"] - 1),
                            anchor="center")
            style.map("Primary.TButton",
                      background=[("disabled", UI["off"]),
                                  ("pressed", UI["accent_dark"]),
                                  ("active", UI["accent_dark"])],
                      foreground=[("disabled", UI["panel"])])

            style.configure("Ghost.TButton", background=UI["panel"],
                            foreground=UI["ink_soft"], font=ui_font("small"),
                            borderwidth=1, relief="solid",
                            bordercolor=UI["field_line"], focusthickness=0,
                            padding=(SPACE["md"], SPACE["sm"] - 1))
            style.map("Ghost.TButton",
                      background=[("active", UI["accent_tint"]),
                                  ("pressed", UI["accent_tint"])],
                      foreground=[("active", UI["ink"]),
                                  ("disabled", UI["ink_faint"])],
                      bordercolor=[("active", UI["accent"]),
                                   ("disabled", UI["line"])])

            style.configure("Link.TButton", background=UI["bg"],
                            foreground=UI["accent"], font=ui_font("small"),
                            borderwidth=0, focusthickness=0,
                            padding=(SPACE["xs"], SPACE["xs"]))
            style.map("Link.TButton",
                      background=[("active", UI["bg"])],
                      foreground=[("active", UI["accent_dark"])])

            # The tick box itself is left to the theme to draw — overriding
            # its indicator colours made it render as a crude cross rather
            # than a tick. Only the surround is restyled.
            style.configure("TCheckbutton", background=UI["panel"],
                            foreground=UI["ink_soft"], font=ui_font("small"),
                            focuscolor=UI["accent"])
            style.map("TCheckbutton",
                      background=[("active", UI["panel"])],
                      foreground=[("active", UI["ink"])])

            # The same tick box, sitting on the window background.
            style.configure("Bar.TCheckbutton", background=UI["bg"],
                            foreground=UI["ink_soft"], font=ui_font("small"))
            style.map("Bar.TCheckbutton",
                      background=[("active", UI["bg"])],
                      foreground=[("active", UI["ink"])])

            # The same tick box, sitting on the tinted merge strip.
            style.configure("Merge.TCheckbutton", background=UI["accent_tint"],
                            foreground=UI["ink_soft"], font=ui_font("small"))
            style.map("Merge.TCheckbutton",
                      background=[("active", UI["accent_tint"])],
                      foreground=[("active", UI["ink"])])

            # ── The two tables ────────────────────────────────────────────────
            # The file list and the results share one look, because they are
            # read the same way: down a column, one row at a time. Rows are
            # given room to breathe — the old 25 pixels ran them together.
            for name, height in (("Results", 29), ("Files", 26)):
                style.configure(f"{name}.Treeview", background=UI["panel"],
                                fieldbackground=UI["panel"],
                                foreground=UI["ink"], rowheight=height,
                                borderwidth=0, font=ui_font("table"))
                style.configure(f"{name}.Treeview.Heading",
                                background=UI["panel"],
                                foreground=UI["ink_faint"],
                                font=ui_font("label"), relief="flat",
                                borderwidth=0,
                                padding=(SPACE["sm"], SPACE["sm"]))
                style.map(f"{name}.Treeview.Heading",
                          background=[("active", UI["panel"])],
                          foreground=[("active", UI["ink_soft"])])
                style.map(f"{name}.Treeview",
                          background=[("selected", UI["accent_tint"])],
                          foreground=[("selected", UI["ink"])])

            # ── Scrollbars ────────────────────────────────────────────────────
            # Thin, in the hairline colour, with no arrow buttons. A chunky
            # grey scrollbar is the loudest thing in a quiet window.
            style.configure("Slim.Vertical.TScrollbar",
                            background=UI["field_line"], troughcolor=UI["panel"],
                            bordercolor=UI["panel"], arrowcolor=UI["panel"],
                            darkcolor=UI["field_line"],
                            lightcolor=UI["field_line"],
                            borderwidth=0, relief="flat", width=10,
                            arrowsize=0)
            style.map("Slim.Vertical.TScrollbar",
                      background=[("active", UI["ink_faint"]),
                                  ("pressed", UI["ink_soft"])])

            style.configure("Brand.Horizontal.TProgressbar",
                            background=UI["accent"], troughcolor=UI["line_soft"],
                            borderwidth=0, thickness=3)

        # ══════════════════════════════════════════════════════════════════════
        # MENU BAR
        # ══════════════════════════════════════════════════════════════════════

        def _build_menu(self):
            # Windows draws its own menu bar and ignores most of this; on
            # Linux it is what stops the bar looking pasted on.
            tint = {"bg": UI["panel"], "fg": UI["ink"],
                    "activebackground": UI["accent_tint"],
                    "activeforeground": UI["ink"],
                    "selectcolor": UI["accent"], "borderwidth": 0}

            menubar = tk.Menu(self, **tint)

            file_menu = tk.Menu(menubar, tearoff=0, **tint)
            file_menu.add_command(label="Add analyser PDFs...   Ctrl+O",
                                  command=self._browse_input)
            file_menu.add_command(label="Generate reports       Ctrl+G",
                                  command=self._generate)
            file_menu.add_separator()
            file_menu.add_command(label="Open the save folder",
                                  command=self._open_output_folder)
            file_menu.add_separator()
            file_menu.add_command(label="Close LabScribe", command=self.destroy)
            menubar.add_cascade(label="File", menu=file_menu)

            setup_menu = tk.Menu(menubar, tearoff=0, **tint)
            setup_menu.add_command(label="Laboratory name and signature...",
                                   command=self._edit_lab_details)
            setup_menu.add_checkbutton(
                label="Mark High and Low results in the FLAG column",
                variable=self.flags_var, command=self._save_settings)
            setup_menu.add_separator()
            setup_menu.add_command(label="Reference intervals in use...",
                                   command=self._show_intervals)
            setup_menu.add_separator()

            look_menu = tk.Menu(setup_menu, tearoff=0, **tint)
            self.style_var = tk.StringVar(value=self.style_id)
            for style in STYLES:
                look_menu.add_radiobutton(
                    label=style["name"], value=style["id"],
                    variable=self.style_var,
                    command=lambda sid=style["id"]: self._set_style(sid))
            look_menu.add_separator()
            look_menu.add_command(label="See them side by side...",
                                  command=self._choose_style)
            setup_menu.add_cascade(label="Appearance", menu=look_menu)
            menubar.add_cascade(label="Settings", menu=setup_menu)

            help_menu = tk.Menu(menubar, tearoff=0, **tint)
            help_menu.add_command(label=f"About {APP_NAME}", command=self._show_about)
            menubar.add_cascade(label="Help", menu=help_menu)

            self.config(menu=menubar)
            self.bind("<Control-o>", lambda _e: self._browse_input())
            self.bind("<Control-g>", lambda _e: self._generate())

        # ══════════════════════════════════════════════════════════════════════
        # LAYOUT
        # ══════════════════════════════════════════════════════════════════════

        def _build_layout(self):
            self._build_header()

            # The bar along the bottom is packed before the body, against the
            # bottom edge. Packed after, it was the part that got squeezed off
            # the window when a long results table made the middle grow — and
            # the part that must never disappear is the one with Generate on
            # it. Now the middle takes whatever is left over instead.
            self._build_action_bar()

            body = tk.Frame(self, bg=UI["bg"])
            body.pack(fill="both", expand=True,
                      padx=SPACE["lg"], pady=(SPACE["lg"], 0))
            body.columnconfigure(0, minsize=330, weight=0)
            body.columnconfigure(1, weight=1)
            body.rowconfigure(0, weight=1)

            self._build_file_column(body)
            self._build_detail_column(body)

            # Registered last, once every widget a file might be dropped on
            # actually exists.
            self._register_drop_targets()

        # ── The strip across the top ──────────────────────────────────────────
        # Deliberately quiet. It carries the program's name and the laboratory
        # it is reporting for, and neither of those is what anyone came to the
        # window to look at — so it is a band, not a banner.
        def _build_header(self):
            quiet = UI.get("header") == "quiet"

            back  = UI["panel"] if quiet else UI["accent"]
            title = UI["accent"] if quiet else UI["on_accent"]
            sub   = UI["ink_faint"] if quiet else UI["accent_soft"]

            bar = tk.Frame(self, bg=back)
            bar.pack(fill="x")

            inner = tk.Frame(bar, bg=back)
            inner.pack(fill="x", padx=SPACE["lg"], pady=SPACE["md"])

            left = tk.Frame(inner, bg=back)
            left.pack(side="left")
            tk.Label(left, text=APP_NAME, bg=back, fg=title,
                     font=ui_font("display")).pack(side="left")
            tk.Label(left, text=APP_TAGLINE, bg=back, fg=sub,
                     font=ui_font("small")).pack(side="left",
                                                 padx=(SPACE["md"], 0),
                                                 pady=(3, 0))

            # The laboratory name opens the box where it is changed. It is a
            # control, so it behaves like one: it lights under the pointer
            # instead of only turning the cursor into a hand.
            right = tk.Frame(inner, bg=back)
            right.pack(side="right")
            tk.Label(right, text="REPORTING AS", bg=back, fg=sub,
                     font=ui_font("label")).pack(anchor="e")
            self.lab_label = tk.Label(right, textvariable=self.lab_name_var,
                                      bg=back, fg=title,
                                      font=ui_font("subtitle"), cursor="hand2")
            self.lab_label.pack(anchor="e")
            self.lab_label.bind("<Button-1>", lambda _e: self._edit_lab_details())
            self.lab_label.bind("<Enter>",
                                lambda _e: self.lab_label.config(fg=sub))
            self.lab_label.bind("<Leave>",
                                lambda _e: self.lab_label.config(fg=title))

            rule = UI["accent"] if quiet else UI["accent_dark"]
            tk.Frame(self, bg=rule, height=2).pack(fill="x")

        # ── Left: the queue of analyser files ─────────────────────────────────
        def _build_file_column(self, parent):
            card = self._card(parent, row=0, column=0, padx=(0, SPACE["md"]))

            head = self._card_head(card, "ANALYSER FILES")
            self.count_label = ttk.Label(head, text="nothing added yet",
                                         style="Hint.TLabel")
            self.count_label.pack(side="right")

            # ── Drop zone ─────────────────────────────────────────────────────
            # Drawn on a canvas so the border can be dashed. A dashed outline
            # is what people already read as "put something here"; the flat
            # grey box it replaces read as a disabled panel.
            self.drop = tk.Canvas(card, bg=UI["panel"], height=86,
                                  highlightthickness=0, cursor="hand2")
            self.drop.pack(fill="x", padx=SPACE["lg"])
            self.drop.bind("<Configure>", lambda _e: self._draw_drop_zone())
            self.drop.bind("<Button-1>", lambda _e: self._browse_input())
            self.drop.bind("<Enter>", self._on_hover_enter)
            self.drop.bind("<Leave>", self._on_hover_leave)

            # ── The list itself ───────────────────────────────────────────────
            # A Treeview rather than a Listbox, so the rows can be given the
            # same height and the same slim scrollbar as the results table.
            # The two lists are read the same way and now look the same.
            list_wrap = tk.Frame(card, bg=UI["panel"])
            list_wrap.pack(fill="both", expand=True, padx=SPACE["lg"],
                           pady=(SPACE["md"], SPACE["sm"]))

            self.file_list = ttk.Treeview(list_wrap, show="tree",
                                          style="Files.Treeview",
                                          selectmode="browse", height=6)
            self.file_list.column("#0", width=280, stretch=True)
            self.file_list.pack(side="left", fill="both", expand=True)
            self.file_list.bind("<<TreeviewSelect>>", self._on_file_selected)
            self.file_list.bind("<Delete>", lambda _e: self._remove_selected())

            self.file_list.tag_configure("ready", foreground=UI["ink"])
            self.file_list.tag_configure("reading", foreground=UI["ink_faint"])
            self.file_list.tag_configure("error", foreground=UI["danger"])
            self.file_list.tag_configure("done", foreground=UI["accent"])

            bar = ttk.Scrollbar(list_wrap, orient="vertical",
                                style="Slim.Vertical.TScrollbar",
                                command=self.file_list.yview)
            bar.pack(side="right", fill="y")
            self.file_list.configure(yscrollcommand=bar.set)

            buttons = tk.Frame(card, bg=UI["panel"])
            buttons.pack(fill="x", padx=SPACE["lg"],
                         pady=(0, SPACE["lg"]))
            ttk.Button(buttons, text="Add...", style="Ghost.TButton",
                       command=self._browse_input).pack(side="left")
            self.remove_btn = ttk.Button(buttons, text="Remove",
                                         style="Ghost.TButton",
                                         command=self._remove_selected)
            self.remove_btn.pack(side="left", padx=SPACE["sm"])
            self.clear_btn = ttk.Button(buttons, text="Clear all",
                                        style="Ghost.TButton",
                                        command=self._clear_all)
            self.clear_btn.pack(side="left")

        def _register_drop_targets(self):
            """
            Makes the WHOLE window take a dropped file, not only the dashed
            panel.

            Dropping had to land inside one small rectangle to count, which is
            a fiddly thing to ask of anyone holding a file over a window — most
            of a miss looks exactly like drag-and-drop being broken. Now the
            window, the file list and the panel all accept a drop, and the
            panel simply lights up to say where the files will go.
            """
            if not DND_AVAILABLE:
                return

            targets = [self, self.drop]
            if getattr(self, "file_list", None) is not None:
                targets.append(self.file_list)

            for widget in targets:
                try:
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>", self._on_drop)
                    widget.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                    widget.dnd_bind("<<DragLeave>>", self._on_drag_leave)
                except Exception:
                    # One widget refusing to register is not a reason to lose
                    # the others.
                    continue

        def _draw_drop_zone(self):
            """
            Redraws the dashed panel files are dropped onto. Called whenever it
            changes size, and whenever something is dragged over it.
            """
            canvas = self.drop
            canvas.delete("all")
            width = canvas.winfo_width()
            height = int(canvas["height"])
            if width < 10:
                return

            lit = self._drop_hover or self._drop_dragging
            edge = UI["accent"] if lit else UI["field_line"]
            fill = UI["accent_tint"] if lit else UI["drop"]
            ink = UI["accent"] if lit else UI["ink_faint"]

            canvas.create_rectangle(2, 2, width - 3, height - 3,
                                    outline=edge, fill=fill,
                                    dash=(4, 3), width=2 if lit else 1)

            if self._drop_dragging:
                first, second = "Let go to add them", ""
            elif DND_AVAILABLE:
                first, second = ("Drop the analyser PDFs anywhere",
                                 "or click to go and find them")
            else:
                first, second = ("Click to add the analyser PDFs",
                                 "dragging is off - see Help, About")

            canvas.create_text(width // 2, height // 2 - (8 if second else 0),
                               text=first, fill=ink,
                               font=ui_font("small_bold"))
            if second:
                canvas.create_text(width // 2, height // 2 + 11, text=second,
                                   fill=UI["ink_faint"], font=ui_font("small"))

        # ── Right: patient details, then the results read from the file ───────
        def _build_detail_column(self, parent):
            column = tk.Frame(parent, bg=UI["bg"])
            column.grid(row=0, column=1, sticky="nsew")
            column.rowconfigure(1, weight=1)
            column.columnconfigure(0, weight=1)

            # ── Patient card ──────────────────────────────────────────────────
            card = self._card(column, row=0, column=0, pady=(0, SPACE["md"]))

            head = self._card_head(card, "PATIENT DETAILS")
            ttk.Label(head, text="these print at the top of the report",
                      style="Hint.TLabel").pack(side="left", padx=(SPACE["sm"], 0))

            grid = tk.Frame(card, bg=UI["panel"])
            grid.pack(fill="x", padx=SPACE["lg"], pady=(0, SPACE["md"]))
            for position in range(3):
                grid.columnconfigure(position, weight=1, uniform="field")

            # The label sits above its box rather than beside it. Six boxes in
            # two columns with the labels alongside left every box squeezed
            # into half the width it deserved; stacked, they get the room, and
            # the six read as one block instead of twelve scattered things.
            self.field_entries = {}
            for index, (key, label) in enumerate(PATIENT_FIELDS):
                row, col = divmod(index, 3)
                cell = tk.Frame(grid, bg=UI["panel"])
                cell.grid(row=row, column=col, sticky="ew",
                          padx=(0, 0 if col == 2 else SPACE["lg"]),
                          pady=(0, SPACE["md"]))
                cell.columnconfigure(0, weight=1)
                ttk.Label(cell, text=label.upper(),
                          style="CardTitle.TLabel").grid(row=0, column=0,
                                                         sticky="w",
                                                         pady=(0, SPACE["xs"]))
                entry = ttk.Entry(cell, textvariable=self.field_vars[key],
                                  style="Field.TEntry", font=ui_font("body"))
                entry.grid(row=1, column=0, sticky="ew")
                self.field_entries[key] = entry

            # ── The merge notice ──────────────────────────────────────────────
            # A tinted strip along the bottom of the patient card, shown only
            # when the selected file shares its patient with another. Tinted,
            # because it is saying that what is above it came from more than
            # one file, and that needs to be noticed rather than read.
            self.merge_bar = tk.Frame(card, bg=UI["accent_tint"])
            self.separate_check = ttk.Checkbutton(
                self.merge_bar, text="Report this file on its own",
                variable=self.separate_var, command=self._on_separate_toggled,
                style="Merge.TCheckbutton")
            self.separate_check.pack(side="right", padx=SPACE["md"],
                                     pady=SPACE["sm"])
            self.merge_note = tk.Label(self.merge_bar, text="",
                                       bg=UI["accent_tint"], fg=UI["ink_soft"],
                                       font=ui_font("small"), anchor="w",
                                       justify="left", wraplength=560)
            self.merge_note.pack(side="left", fill="x", expand=True,
                                 padx=(SPACE["md"], 0), pady=SPACE["sm"])

            # ── Results card ──────────────────────────────────────────────────
            card2 = self._card(column, row=1, column=0)

            head2 = self._card_head(card2, "WHAT WAS READ FROM THIS FILE")
            self.results_title = head2.winfo_children()[0]
            self.results_note = ttk.Label(head2, text="", style="Hint.TLabel")
            self.results_note.pack(side="right")

            # The panel shown instead of the table when a file will not read.
            self.trouble = tk.Frame(card2, bg=UI["panel"])

            # Shown in place of the table before anything has been added. It is
            # the screen the laboratory opens on every morning, so it says what
            # to do next rather than sitting blank.
            self.blank = tk.Frame(card2, bg=UI["panel"])
            self._build_blank_state(self.blank)

            table_wrap = tk.Frame(card2, bg=UI["panel"])
            self.table_wrap = table_wrap
            table_wrap.pack(fill="both", expand=True, padx=SPACE["lg"],
                            pady=(0, SPACE["lg"]))

            columns = ("section", "test", "result", "unit", "flag", "ref")
            self.table = ttk.Treeview(table_wrap, columns=columns,
                                      show="headings", style="Results.Treeview",
                                      selectmode="none", height=6)
            headings = [("section", "SECTION", 76, "w"),
                        ("test", "INVESTIGATION", 168, "w"),
                        ("result", "RESULT", 64, "e"),
                        ("unit", "UNIT", 62, "w"),
                        ("flag", "FLAG", 48, "center"),
                        ("ref", "REFERENCE", 130, "w")]
            for key, title, width, anchor in headings:
                self.table.heading(key, text=title, anchor=anchor)
                self.table.column(key, width=width, anchor=anchor,
                                  stretch=(key in ("test", "ref")))

            self.table.tag_configure("odd", background=UI["row_alt"])
            self.table.tag_configure("section", foreground=UI["ink_soft"])
            self.table.tag_configure("abnormal", foreground=UI["danger"])
            self.table.tag_configure("empty", foreground=UI["ink_faint"])

            self.table.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(table_wrap, orient="vertical",
                                   style="Slim.Vertical.TScrollbar",
                                   command=self.table.yview)
            scroll.pack(side="right", fill="y")
            self.table.configure(yscrollcommand=scroll.set)

            # Codes LabScribe did not recognise, named under the table rather
            # than left for someone to notice in OTHER INVESTIGATIONS.
            self.table_note = ttk.Label(card2, text="", style="Hint.TLabel",
                                        wraplength=620, justify="left")

            # A row lights under the pointer, so the eye can follow one test
            # across to its reference interval.
            self.table.tag_configure("hover", background=UI["accent_tint"])
            self.table.bind("<Motion>", self._hover_row)
            self.table.bind("<Leave>", lambda _e: self._hover_row(None))
            self._hovered = None

        # ── Bottom: where it saves, and the Generate button ───────────────────
        def _build_action_bar(self):
            # A rule, then the bar. The bottom of the window reads as a
            # separate place where something happens, rather than more of the
            # same panel.
            status = self.status_frame = tk.Frame(self, bg=UI["bg"])
            status.pack(side="bottom", fill="x", padx=SPACE["lg"],
                        pady=(SPACE["md"], SPACE["md"]))

            bar = tk.Frame(self, bg=UI["bg"])
            bar.pack(side="bottom", fill="x", padx=SPACE["lg"],
                     pady=(SPACE["md"], 0))

            tk.Frame(self, bg=UI["line"], height=1).pack(
                side="bottom", fill="x", pady=(SPACE["lg"], 0))

            left = tk.Frame(bar, bg=UI["bg"])
            left.pack(side="left", fill="x", expand=True)

            title = tk.Frame(left, bg=UI["bg"])
            title.pack(fill="x", pady=(0, SPACE["xs"]))
            ttk.Label(title, text="SAVE REPORTS TO",
                      style="BarTitle.TLabel").pack(side="left")

            # Ticked by default. Each report is written into the same folder as
            # the analyser file it was read from, so a day's work stays
            # together and nothing has to be chosen before pressing Generate.
            self.beside_check = ttk.Checkbutton(
                title, text="beside the analyser files",
                variable=self.beside_var, command=self._on_beside_toggled,
                style="Bar.TCheckbutton")
            self.beside_check.pack(side="left", padx=(SPACE["md"], 0))

            row = tk.Frame(left, bg=UI["bg"])
            row.pack(fill="x")
            self.output_entry = ttk.Entry(row, textvariable=self.output_var,
                                          style="Field.TEntry",
                                          font=ui_font("small"))
            self.output_entry.pack(side="left", fill="x", expand=True)
            self.output_entry.bind("<FocusOut>", lambda _e: self._save_settings())
            self.browse_btn = ttk.Button(row, text="Browse...",
                                         style="Ghost.TButton",
                                         command=self._browse_output)
            self.browse_btn.pack(side="left", padx=(SPACE["sm"], 0))

            right = tk.Frame(bar, bg=UI["bg"])
            right.pack(side="right", padx=(SPACE["xl"], 0))

            # A button that cannot be pressed says why, rather than sitting
            # there greyed out and leaving the reason to be guessed at.
            self.generate_hint = ttk.Label(right, text="", style="Status.TLabel",
                                           anchor="e")
            self.generate_hint.pack(side="top", fill="x", pady=(0, SPACE["xs"]))

            self.generate_btn = ttk.Button(right, text="Generate reports",
                                           style="Primary.TButton",
                                           command=self._generate)
            self.generate_btn.pack(side="top", fill="x")

            # Packed only while something is being written — an empty trough
            # sitting there the rest of the time is a progress bar that has
            # stopped meaning anything.
            self.progress = ttk.Progressbar(
                self, mode="determinate", style="Brand.Horizontal.TProgressbar")

            # A plain tk.Label rather than a ttk one, so that its background
            # can be tinted for a moment when the message changes — see
            # _set_status(). A status line nobody notices changing is a status
            # line that may as well not be there.
            self.status_label = tk.Label(status, textvariable=self.status_var,
                                         bg=UI["bg"], fg=UI["ink_soft"],
                                         font=ui_font("small"), anchor="w",
                                         padx=SPACE["sm"], pady=2)
            self.status_label.pack(side="left")
            self._status_flash = None

            ttk.Label(status, text=f"{APP_NAME} {APP_VERSION}  ·  {APP_PUBLISHER}",
                      style="Credit.TLabel").pack(side="right")
            self.open_link = ttk.Button(status, text="Open folder",
                                        style="Link.TButton",
                                        command=self._open_output_folder)

        def _build_blank_state(self, parent):
            """
            What fills the results card before a single file has been added.

            This is the screen the laboratory opens on every morning, and it
            used to be an empty table with "Nothing here yet." in it. Three
            numbered steps cost nothing and mean a new member of staff can be
            left alone with the program.
            """
            # Packed rather than placed, so the block can never be drawn
            # partly off the top of the card when the window is short.
            middle = tk.Frame(parent, bg=UI["panel"])
            middle.pack(anchor="w", padx=SPACE["sm"], pady=SPACE["sm"])

            tk.Label(middle, text="Nothing added yet", bg=UI["panel"],
                     fg=UI["ink"], font=ui_font("title")).pack(anchor="w")
            tk.Label(middle, text="Three steps, and the reports are written.",
                     bg=UI["panel"], fg=UI["ink_faint"],
                     font=ui_font("small")).pack(anchor="w",
                                                 pady=(SPACE["xs"], SPACE["lg"]))

            steps = [
                ("1", "Add the analyser PDFs",
                 "Drop them on the left, or press Add. Each one is read as it "
                 "arrives."),
                ("2", "Check what was read",
                 "The patient details and every result appear here. Type over "
                 "anything the analyser got wrong."),
                ("3", "Choose a folder and press Generate",
                 "One report per patient. Two files for the same patient print "
                 "together."),
            ]

            for number, title, detail in steps:
                row = tk.Frame(middle, bg=UI["panel"])
                row.pack(fill="x", anchor="w", pady=(0, SPACE["md"]))

                badge = tk.Label(row, text=number, bg=UI["accent_tint"],
                                 fg=UI["accent"], font=ui_font("small_bold"),
                                 width=3, pady=3)
                badge.pack(side="left", anchor="n")

                words = tk.Frame(row, bg=UI["panel"])
                words.pack(side="left", fill="x", expand=True,
                           padx=(SPACE["md"], 0))
                tk.Label(words, text=title, bg=UI["panel"], fg=UI["ink"],
                         font=ui_font("small_bold"), anchor="w").pack(anchor="w")
                tk.Label(words, text=detail, bg=UI["panel"],
                         fg=UI["ink_faint"], font=ui_font("small"),
                         anchor="w", justify="left",
                         wraplength=380).pack(anchor="w")

        def _show_panel(self, name):
            """
            Only ever one of the three in the results card: the table, the
            reading panel, or the three steps. Switching is done in one place
            so two of them can never be left showing at once.
            """
            if name != "trouble":
                self._release_wheel()

            for key, widget in (("table", self.table_wrap),
                                ("trouble", self.trouble),
                                ("blank", self.blank)):
                if key == name:
                    if not widget.winfo_manager():
                        widget.pack(fill="both", expand=True,
                                    padx=SPACE["lg"], pady=(0, SPACE["lg"]))
                elif widget.winfo_manager():
                    widget.pack_forget()

            if name != "table":
                self.table_note.pack_forget()

        def _card(self, parent, row, column, padx=0, pady=0):
            """A panel with a hairline border. Everything sits in one of these."""
            outer = tk.Frame(parent, bg=UI["line"])
            outer.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
            inner = tk.Frame(outer, bg=UI["panel"])
            inner.pack(fill="both", expand=True, padx=1, pady=1)
            return inner

        def _card_head(self, card, title):
            """
            The small capitals across the top of a card, with whatever else
            belongs up there packed in beside them. One helper, so that every
            card is titled at the same size, in the same colour, the same
            distance from the thing it is titling.
            """
            head = tk.Frame(card, bg=UI["panel"])
            head.pack(fill="x", padx=SPACE["lg"],
                      pady=(SPACE["lg"], SPACE["md"]))
            ttk.Label(head, text=title, style="CardTitle.TLabel").pack(side="left")
            return head

        # ══════════════════════════════════════════════════════════════════════
        # ADDING AND REMOVING FILES
        # ══════════════════════════════════════════════════════════════════════

        def _browse_input(self):
            if self.busy:
                return
            paths = filedialog.askopenfilenames(
                title="Choose the analyser PDFs",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
            self._add_paths(paths)

        def _on_drop(self, event):
            """
            What was let go of over the window.

            The list arrives as Tcl wrote it, which means any path with a space
            in it is wrapped in braces. Tcl's own splitter is asked to take it
            apart, because doing it by hand loses a file the moment two of them
            are dropped together and one has a space in its name.
            """
            try:
                paths = [str(part) for part in self.tk.splitlist(event.data)]
            except Exception:
                raw = str(event.data).strip()
                paths = [part.strip().strip("{}")
                         for part in raw.replace("} {", "\n").split("\n")
                         if part.strip()]

            self._add_paths(paths)
            self._on_drag_leave(event)

        def _on_hover_enter(self, _event=None):
            self._drop_hover = True
            self._draw_drop_zone()

        def _on_hover_leave(self, _event=None):
            self._drop_hover = False
            self._draw_drop_zone()

        def _on_drag_enter(self, _event=None):
            self._drop_dragging = True
            self._draw_drop_zone()

        def _on_drag_leave(self, _event=None):
            self._drop_dragging = False
            self._draw_drop_zone()

        def _add_paths(self, paths):
            """Queues each PDF and starts reading it in the background."""
            added = 0
            skipped = 0

            # A dropped folder means the PDFs inside it, which is how a day's
            # work usually arrives.
            opened = []
            for path in paths:
                path = os.path.abspath(str(path).strip())
                if os.path.isdir(path):
                    try:
                        opened.extend(
                            os.path.join(path, name)
                            for name in sorted(os.listdir(path))
                            if name.lower().endswith(".pdf"))
                    except Exception:
                        skipped += 1
                else:
                    opened.append(path)

            for path in opened:
                path = os.path.abspath(path)
                if not path.lower().endswith(".pdf") or not os.path.isfile(path):
                    skipped += 1
                    continue
                if any(job["path"] == path for job in self.jobs):
                    continue        # Already in the list
                self.jobs.append({
                    "path": path,
                    "name": os.path.basename(path),
                    "status": "reading",
                    "patient": {}, "results": [], "error": "", "out": "",
                    "separate": False, "diagnostics": {},
                })
                added += 1

            if not added:
                if skipped:
                    self._set_status(f"Left out {skipped} file(s). LabScribe reads "
                                     "PDFs only.", "warn")
                return

            if not self.output_var.get().strip():
                self.output_var.set(os.path.dirname(self.jobs[0]["path"]))

            self._refresh_file_list()
            if self.current is None:
                self._select_row(len(self.jobs) - added)

            self._set_status(f"Reading {added} file(s)...")
            threading.Thread(target=self._read_worker,
                             args=([j["path"] for j in self.jobs
                                    if j["status"] == "reading"],),
                             daemon=True).start()

        def _read_worker(self, paths):
            """Runs off the main thread: opens each PDF and parses it."""
            for path in paths:
                diagnostics = {}
                try:
                    patient, results, diagnostics = read_pdf(path)
                    if not results:
                        raise ValueError("nothing on this file was recognised "
                                         "as a result")
                    payload = {"status": "ready", "patient": patient,
                               "results": results, "error": "",
                               "diagnostics": diagnostics}
                except Exception as exc:
                    payload = {"status": "error", "patient": {}, "results": [],
                               "error": str(exc), "diagnostics": diagnostics}
                self.after(0, self._file_read, path, payload)

            self.after(0, self._reading_finished)

        def _file_read(self, path, payload):
            """Back on the main thread: store what was read and redraw."""
            for index, job in enumerate(self.jobs):
                if job["path"] == path:
                    job.update(payload)
                    self._refresh_file_list()
                    if self.current == index:
                        self._show_job(index)
                    break

        def _reading_finished(self):
            ready = sum(1 for j in self.jobs if j["status"] == "ready")
            failed = sum(1 for j in self.jobs if j["status"] == "error")

            if failed and ready:
                self._set_status(f"{ready} file(s) ready. {failed} could not be read.",
                                 "warn")
            elif failed:
                self._set_status(f"{failed} file(s) could not be read. Click one "
                                 "to see why.", "danger")
            else:
                self._set_status(f"{ready} file(s) ready. Check the details, then "
                                 "press Generate.")

            # Show the first file automatically so there is something to look at
            if self.current is None and self.jobs:
                self._select_row(0)

        def _remove_selected(self):
            if self.busy:
                return
            index = self._selected_index()
            if index is None:
                return
            del self.jobs[index]
            self.current = None
            self._refresh_file_list()
            if self.jobs:
                self._select_row(min(index, len(self.jobs) - 1))
            else:
                self._show_job(None)
                self._set_status("Add the PDFs the analyser printed and LabScribe "
                             "will read them.")

        def _clear_all(self):
            if self.busy:
                return
            self.jobs.clear()
            self.current = None
            self._refresh_file_list()
            self._show_job(None)
            self._set_status("Add one or more analyser PDFs to begin.")

        # ══════════════════════════════════════════════════════════════════════
        # SHOWING WHAT WAS READ
        # ══════════════════════════════════════════════════════════════════════

        MARKERS = {"reading": "·", "ready": "✓", "error": "!", "done": "✓"}

        # ══════════════════════════════════════════════════════════════════════
        # WHICH FILES BELONG TOGETHER
        # ══════════════════════════════════════════════════════════════════════

        def _ready_jobs(self):
            return [job for job in self.jobs
                    if job["status"] in ("ready", "done")]

        def _ready_groups(self):
            """The files that will print, sorted into one group per patient."""
            return group_jobs_by_patient(self._ready_jobs())

        def _group_of(self, job):
            """
            The files that will print on the same report as this one, itself
            included. A file on its own comes back as a list of one.
            """
            if job is None or job["status"] not in ("ready", "done"):
                return [job] if job else []
            for group in self._ready_groups():
                if any(member is job for member in group):
                    return group
            return [job]

        def _merge_labels(self):
            """
            A short tag for each file saying which report it is going onto —
            "1 of 2" and so on — so the joining is visible in the file list.
            Files printing on their own get no tag.
            """
            labels = {}
            for group in self._ready_groups():
                if len(group) < 2:
                    continue
                for position, job in enumerate(group, 1):
                    labels[id(job)] = f"one report, {position} of {len(group)}"
            return labels

        def _refresh_file_list(self):
            """Redraws the file list, keeping whatever was selected selected."""
            selected = self._selected_index()
            self.file_list.delete(*self.file_list.get_children())

            merge_labels = self._merge_labels()

            for index, job in enumerate(self.jobs):
                marker = self.MARKERS.get(job["status"], "·")
                suffix = {"reading": "reading", "error": "could not read",
                          "done": "report saved"}.get(job["status"], "")
                extra = merge_labels.get(id(job), "")
                tail = "   ·   ".join(part for part in (suffix, extra) if part)
                text = f" {marker}  {job['name']}"
                if tail:
                    text += f"   ·   {tail}"
                self.file_list.insert("", "end", iid=str(index), text=text,
                                      tags=(job["status"],))

            if selected is not None and selected < len(self.jobs):
                self.file_list.selection_set(str(selected))

            total = len(self.jobs)
            ready = len(self._ready_jobs())
            reports = len(self._ready_groups())
            text = "nothing added yet" if not total else f"{ready} of {total} ready"
            if reports and reports != ready:
                text += f" · {reports} report" + ("" if reports == 1 else "s")
            self.count_label.config(text=text)

            # Buttons that would do nothing are turned off rather than left
            # live to do nothing when pressed.
            state = "disabled" if not self.jobs or self.busy else "normal"
            self.clear_btn.config(state=state)
            self.remove_btn.config(
                state="normal" if (self.jobs and not self.busy
                                   and self._selected_index() is not None)
                else "disabled")

            if bool(self.beside_var.get()):
                self._apply_beside_state()

            self._update_generate_button()

        def _selected_index(self):
            """Which row of the file list is selected, as a number, or None."""
            chosen = self.file_list.selection()
            if not chosen:
                return None
            try:
                return int(chosen[0])
            except (TypeError, ValueError):
                return None

        def _select_row(self, index):
            if index is None or index >= len(self.jobs):
                return
            self.file_list.selection_set(str(index))
            self.file_list.see(str(index))
            self._show_job(index)

        def _on_file_selected(self, _event):
            index = self._selected_index()
            if index is not None:
                self.remove_btn.config(state="disabled" if self.busy else "normal")
                self._show_job(index)

        def _show_job(self, index):
            """Fills the patient fields and the results table from one file."""
            self.current = index

            self._loading_fields = True
            try:
                if index is None or index >= len(self.jobs):
                    for var in self.field_vars.values():
                        var.set("")
                    self._set_fields_enabled(False)
                    self.merge_bar.pack_forget()
                    self.results_title.config(text="GETTING STARTED")
                    self.results_note.config(text="")
                    self._show_panel("blank")
                    return

                job = self.jobs[index]
                group = self._group_of(job)
                self._show_merge_notice(job, group)

                # For a file that is joined to another, the boxes show the
                # details the merged report will actually print, not just what
                # was on this one file.
                patient = (merge_patient_fields(
                                [(member.get("patient") or {},
                                  member.get("results") or [])
                                 for member in group])
                           if len(group) > 1 else (job.get("patient") or {}))

                for key, _label in PATIENT_FIELDS:
                    self.field_vars[key].set(patient.get(key, ""))
                self._set_fields_enabled(job["status"] in ("ready", "done"))

                if job["status"] == "error":
                    self.results_title.config(text="WHY THIS FILE WOULD NOT READ")
                    self.results_note.config(text="")
                    self._show_trouble(job)
                elif job["status"] == "reading":
                    self.results_title.config(text="WHAT WAS READ FROM THIS FILE")
                    self._show_panel("table")
                    self._fill_table([], "Reading it now...")
                    self.results_note.config(text="")
                    self.table_note.pack_forget()
                else:
                    self._show_panel("table")
                    if len(group) > 1:
                        self.results_title.config(
                            text="WHAT WILL PRINT ON THIS REPORT")
                        _patient, results = merged_job(group)
                    else:
                        self.results_title.config(
                            text="WHAT WAS READ FROM THIS FILE")
                        results = recompute_egfr(job["results"], patient)

                    rows = apply_test_library(results, patient)
                    self._fill_table(rows, "")
                    abnormal = sum(1 for r in rows if r.get("flag"))
                    note = f"{len(rows)} results"
                    if len(group) > 1:
                        note += f" · from {len(group)} files"
                    if abnormal:
                        note += f" · {abnormal} flagged"
                    self.results_note.config(text=note)
                    self._show_unknown_codes(group)
            finally:
                self._loading_fields = False

        def _show_unknown_codes(self, group):
            """
            Names the codes on these files that LabScribe does not know.

            Not an error — the results still print, under OTHER INVESTIGATIONS.
            But an unrecognised code is the one thing that turns into a line in
            TEST_CODE_MAP and a test in its proper section, and it was
            previously left for someone to spot on the finished report.
            """
            codes = []
            for job in group:
                for code in (job.get("diagnostics") or {}).get("unknown_codes", []):
                    if code not in codes:
                        codes.append(code)

            if not codes:
                self.table_note.pack_forget()
                return

            shown = ", ".join(codes[:6])
            if len(codes) > 6:
                shown += f" and {len(codes) - 6} more"
            word = "code" if len(codes) == 1 else "codes"
            self.table_note.config(
                text=f"{len(codes)} unrecognised {word}: {shown} — printed "
                     f"under OTHER INVESTIGATIONS.")
            # Packed against the bottom of the card, and the table re-packed
            # after it, so the note is given its space first and a long
            # results table cannot push it off the edge.
            self.table_note.pack(side="bottom", fill="x", padx=SPACE["lg"],
                                 pady=(0, SPACE["md"]))
            if self.table_wrap.winfo_manager():
                self.table_wrap.pack_forget()
                self.table_wrap.pack(fill="both", expand=True,
                                     padx=SPACE["lg"], pady=(0, SPACE["lg"]))

        def _show_merge_notice(self, job, group):
            """
            Says, above the results, which other file this one is being joined
            to — and offers the way out if the two only happen to share a name.
            """
            self.separate_var.set(bool(job.get("separate")))

            if len(group) > 1:
                others = ", ".join(member["name"] for member in group
                                   if member is not job)
                self.merge_note.config(
                    text=f"Same patient as {others} — printed as one report.")
                self.merge_bar.pack(fill="x", side="bottom")
            elif job.get("separate"):
                self.merge_note.config(
                    text="Kept out of its patient's group: this file prints on "
                         "its own.")
                self.merge_bar.pack(fill="x", side="bottom")
            else:
                self.merge_bar.pack_forget()

        def _on_separate_toggled(self):
            """The "Report this file on its own" tick box."""
            if self._loading_fields or self.current is None:
                return
            if self.current >= len(self.jobs):
                return
            self.jobs[self.current]["separate"] = bool(self.separate_var.get())
            self._refresh_file_list()
            self._show_job(self.current)

        def _fill_table(self, rows, message):
            self.table.delete(*self.table.get_children())

            if message:
                self.table.insert("", "end", tags=("empty",),
                                  values=("", message, "", "", "", ""))
                return

            if not rows:
                self.table.insert("", "end", tags=("empty",),
                                  values=("", "Nothing here yet.", "", "", "", ""))
                return

            grouped = group_results_by_section(rows)
            line = 0
            for section in SECTION_ORDER:
                for position, row in enumerate(grouped.get(section, [])):
                    tags = ["odd"] if line % 2 else []
                    if row.get("flag"):
                        tags.append("abnormal")
                    short = SECTIONS.get(section, {}).get("short", section)
                    self.table.insert("", "end", tags=tuple(tags), values=(
                        short if position == 0 else "",
                        _plain(row["test"]),
                        row.get("result", ""),
                        _plain(row.get("unit", "")),
                        row.get("flag", ""),
                        _plain(" / ".join(row.get("ref_lines", []))),
                    ))
                    line += 1

        def _hover_row(self, event):
            """Lights the row under the pointer, and only that row."""
            row = self.table.identify_row(event.y) if event else ""
            if row == self._hovered:
                return

            for item, wanted in ((self._hovered, False), (row, True)):
                if not item or not self.table.exists(item):
                    continue
                tags = [t for t in self.table.item(item, "tags") if t != "hover"]
                if wanted:
                    tags.append("hover")
                self.table.item(item, tags=tuple(tags))

            self._hovered = row or None

        # ══════════════════════════════════════════════════════════════════════
        # WHEN A FILE WILL NOT READ
        # ══════════════════════════════════════════════════════════════════════

        def _show_trouble(self, job):
            """
            Puts the reading panel where the results table normally sits, and
            fills it with what actually happened to this file.
            """
            for child in list(self.trouble.winfo_children()):
                child.destroy()

            headline, evidence, suggestions = explain_reading(
                job.get("diagnostics"), job.get("error", ""))

            # The two buttons sit against the bottom of the panel and stay
            # there. Everything above them scrolls, because how much there is
            # to say depends on the file, and the one thing that must never be
            # pushed out of sight is the way to send the file on.
            buttons = tk.Frame(self.trouble, bg=UI["panel"])
            buttons.pack(side="bottom", fill="x", pady=(12, 0))
            ttk.Button(buttons, text="Show the text LabScribe found",
                       style="Ghost.TButton",
                       command=lambda: self._show_raw_text(job)).pack(side="left")
            ttk.Button(buttons, text="Copy this for support",
                       style="Ghost.TButton",
                       command=lambda: self._copy_trouble(job)).pack(
                           side="left", padx=8)

            canvas = tk.Canvas(self.trouble, bg=UI["panel"], highlightthickness=0)
            bar = ttk.Scrollbar(self.trouble, orient="vertical",
                                style="Slim.Vertical.TScrollbar",
                                command=canvas.yview)
            canvas.configure(yscrollcommand=bar.set)
            canvas.pack(side="left", fill="both", expand=True)

            body = tk.Frame(canvas, bg=UI["panel"])
            window = canvas.create_window((0, 0), window=body, anchor="nw")

            def fit(_event=None):
                canvas.configure(scrollregion=canvas.bbox("all"))
                canvas.itemconfig(window, width=canvas.winfo_width())
                # The scrollbar appears only when there is something to scroll.
                needed = body.winfo_reqheight() > canvas.winfo_height()
                if needed and not bar.winfo_manager():
                    bar.pack(side="right", fill="y")
                elif not needed and bar.winfo_manager():
                    bar.pack_forget()

            body.bind("<Configure>", fit)
            canvas.bind("<Configure>", fit)

            # The wheel is claimed only while the pointer is over this panel,
            # and given back on the way out, so it never scrolls a canvas that
            # has since been taken down.
            def wheel(event):
                canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

            canvas.bind("<Enter>", lambda _e: self.bind_all("<MouseWheel>", wheel))
            canvas.bind("<Leave>", lambda _e: self._release_wheel())

            wrap = 560

            tk.Label(body, text=headline, bg=UI["panel"], fg=UI["danger"],
                     font=ui_font("title"), anchor="w", justify="left",
                     wraplength=wrap).pack(fill="x", pady=(2, 10))

            if evidence:
                tk.Label(body, text="WHAT LABSCRIBE FOUND", bg=UI["panel"],
                         fg=UI["ink_faint"], font=ui_font("label"),
                         anchor="w").pack(fill="x")
                for line in evidence:
                    tk.Label(body, text=line, bg=UI["panel"],
                             fg=UI["ink_soft"], font=ui_font("small"), anchor="w",
                             justify="left", wraplength=wrap).pack(
                                 fill="x", pady=(2, 0))

            if suggestions:
                tk.Label(body, text="WHAT TO TRY", bg=UI["panel"],
                         fg=UI["ink_faint"], font=ui_font("label"),
                         anchor="w").pack(fill="x", pady=(14, 0))
                for line in suggestions:
                    tk.Label(body, text="·  " + line, bg=UI["panel"],
                             fg=UI["ink"], font=ui_font("small"), anchor="w",
                             justify="left", wraplength=wrap).pack(
                                 fill="x", pady=(3, 0))

            self._show_panel("trouble")

        def _release_wheel(self):
            try:
                self.unbind_all("<MouseWheel>")
            except Exception:
                pass

        def _hide_trouble(self):
            self._show_panel("table")

        def _trouble_report(self, job):
            """Everything about one unreadable file, as plain text."""
            diagnostics = job.get("diagnostics") or {}
            headline, evidence, suggestions = explain_reading(
                diagnostics, job.get("error", ""))

            lines = [f"{APP_NAME} {APP_VERSION} — could not read a file",
                     f"File: {job['name']}",
                     "",
                     headline,
                     ""]
            lines += evidence
            lines += ["", "What to try:"] + [f"  · {s}" for s in suggestions]
            lines += ["", "Raw error:", f"  {job.get('error', '')}",
                      "", "The text found on the file:", "",
                      (diagnostics.get("raw_text") or "(none)")[:8000]]
            return "\n".join(lines)

        def _copy_trouble(self, job):
            try:
                self.clipboard_clear()
                self.clipboard_append(self._trouble_report(job))
                self._set_status("Copied. Paste it into an email or a message "
                                 "and send it on.", "ok")
            except Exception:
                self._set_status("That could not be copied to the clipboard.",
                                 "warn")

        def _show_raw_text(self, job):
            """The text LabScribe pulled off the file, exactly as it found it."""
            diagnostics = job.get("diagnostics") or {}
            raw = (diagnostics.get("raw_text") or "").strip()

            dialog = self._dialog(f"The text found on {job['name']}", 760)
            dialog.resizable(True, True)

            body = tk.Frame(dialog, bg=UI["bg"])
            body.pack(fill="both", expand=True, padx=16, pady=(0, 8))

            if raw:
                box = tk.Text(body, wrap="none", height=22, width=92,
                              bg=UI["panel"], fg=UI["ink"], borderwidth=0,
                              highlightthickness=1,
                              highlightbackground=UI["line"],
                              font=("Consolas", 9), padx=10, pady=8)
                box.insert("1.0", raw)
                box.config(state="disabled")
                box.pack(side="left", fill="both", expand=True)
                bar = ttk.Scrollbar(body, orient="vertical",
                                    style="Slim.Vertical.TScrollbar",
                                    command=box.yview)
                bar.pack(side="right", fill="y")
                box.configure(yscrollcommand=bar.set)
            else:
                tk.Label(body, text="There was no text on this file at all.\n"
                                    "That is what a scan looks like.",
                         bg=UI["bg"], fg=UI["ink_soft"], font=ui_font("body"),
                         justify="left").pack(anchor="w", pady=20)

            feet = tk.Frame(dialog, bg=UI["bg"])
            feet.pack(fill="x", padx=16, pady=(0, 14))
            ttk.Button(feet, text="Copy it all", style="Ghost.TButton",
                       command=lambda: self._copy_trouble(job)).pack(side="left")
            ttk.Button(feet, text="Close", style="Ghost.TButton",
                       command=dialog.destroy).pack(side="right")

            self._show_dialog(dialog)

        def _set_fields_enabled(self, enabled):
            state = "normal" if enabled else "disabled"
            for entry in self.field_entries.values():
                entry.config(state=state)

        def _on_field_edited(self, *_args):
            """
            Writes an edited box straight back into the file it belongs to —
            and into every file joined to it, since they share one report and
            one set of printed details. Keeping them in step also means a
            corrected name does not split a pair apart halfway through typing.
            """
            if self._loading_fields or self.current is None:
                return
            if self.current >= len(self.jobs):
                return
            job = self.jobs[self.current]
            if job["status"] not in ("ready", "done"):
                return

            for member in self._group_of(job):
                for key, _label in PATIENT_FIELDS:
                    value = self.field_vars[key].get().strip()
                    member["patient"][key] = (value.upper()
                                              if key in ("name", "sex", "specimen")
                                              else value)

            # A corrected name can join a file to another, or part it from one.
            self._refresh_file_list()

        # ══════════════════════════════════════════════════════════════════════
        # GENERATING THE REPORTS
        # ══════════════════════════════════════════════════════════════════════

        def _update_generate_button(self):
            # One report per patient, not per file — two files carrying the
            # same patient print together, so the button counts groups.
            reports = len(self._ready_groups())
            if self.busy:
                return
            self.generate_btn.config(
                state="normal" if reports else "disabled",
                text="Generate report" if reports == 1
                     else f"Generate {reports} reports" if reports
                     else "Generate reports")
            self.generate_hint.config(text="" if reports
                                      else self._why_not_ready())

        def _why_not_ready(self):
            """Why the Generate button cannot be pressed, in one short line."""
            if not self.jobs:
                return "Add a PDF from the analyser to begin"

            reading = sum(1 for j in self.jobs if j["status"] == "reading")
            failed = sum(1 for j in self.jobs if j["status"] == "error")

            if reading:
                return (f"still reading {reading} file"
                        + ("" if reading == 1 else "s"))
            if failed == len(self.jobs):
                word = "file" if failed == 1 else "files"
                return f"none of the {failed} {word} could be read"
            return "nothing ready to print yet"

        def _generate(self):
            if self.busy:
                return

            ready = [j for j in self.jobs if j["status"] in ("ready", "done")]
            if not ready:
                self._set_status("Nothing to generate yet. Add a PDF from the "
                                 "analyser first.", "warn")
                return

            lab_name = self.lab_name_var.get().strip()
            if not lab_name:
                self._set_status("The laboratory name has gone blank. Please put "
                                 "it back.", "warn")
                self._edit_lab_details()
                return

            # Where the reports go. Beside the file they were read from by
            # default: that is the folder the work for that day already lives
            # in, and it means nothing has to be chosen before pressing
            # Generate. A folder typed in by hand overrides it for everything.
            beside = bool(self.beside_var.get())
            out_dir = "" if beside else self.output_var.get().strip()

            if not beside:
                if not out_dir:
                    self._set_status("Choose a folder for the reports, or tick "
                                     "\"Save beside the analyser files\".",
                                     "warn")
                    return
                try:
                    os.makedirs(out_dir, exist_ok=True)
                except Exception as exc:
                    self._set_status(f"That folder will not take the reports: "
                                     f"{exc}", "danger")
                    return
                self.output_var.set(out_dir)

            # Settings that apply to every report in this run
            global LAB_NAME, SIGNATORY, SHOW_FLAGS
            LAB_NAME = lab_name
            SIGNATORY = self.signatory_var.get().strip() or SIGNATORY
            SHOW_FLAGS = bool(self.flags_var.get())

            # Files carrying the same patient print as one report, so it is the
            # groups that are counted and written, not the files.
            groups = group_jobs_by_patient(ready)

            self._set_busy(True, f"Writing {len(groups)} report(s)...")
            self.progress.config(maximum=len(groups), value=0)
            self.open_link.pack_forget()

            threading.Thread(target=self._generate_worker,
                             args=(groups, out_dir, beside), daemon=True).start()

        def _generate_worker(self, groups, out_dir, beside=False):
            written, errors = [], []

            # Every name written in this run, so that two reports can never be
            # given the same one. Before version 2.4 they could, and the second
            # wrote silently over the first. The names are kept per folder,
            # because two reports in different folders may share a name quite
            # legitimately.
            taken = {}
            folders = []

            for position, group in enumerate(groups, 1):
                names = ", ".join(job["name"] for job in group)
                try:
                    # Beside the file it came from, or in the one folder that
                    # was chosen for the whole run.
                    folder = (os.path.dirname(os.path.abspath(group[0]["path"]))
                              if beside else out_dir)
                    os.makedirs(folder, exist_ok=True)
                    if folder not in folders:
                        folders.append(folder)

                    patient, results = merged_job(group)
                    out_path = unique_output_path(
                        folder,
                        default_output_name(patient, group[0]["path"]),
                        taken.setdefault(os.path.normcase(folder), set()))
                    generate_report(patient, results, out_path)
                    for job in group:
                        job["out"] = out_path
                        job["status"] = "done"
                    written.append(out_path)
                except Exception as exc:
                    errors.append(f"{names}: {exc}")

                self.after(0, self.progress.config, {"value": position})

            # What to show as "the folder" afterwards: the one they all went
            # to, or the first of several.
            landed = folders[0] if folders else out_dir
            self.after(0, self._generation_finished, written, errors, landed,
                       len(folders))

        def _generation_finished(self, written, errors, out_dir, folders=1):
            self._last_output_dir = out_dir
            self._set_busy(False)
            self.progress.config(value=0)
            self._refresh_file_list()
            self._save_settings()

            if not written:
                self._set_status("Not one report could be written.", "danger")
                if errors:
                    messagebox.showerror("Nothing generated", "\n".join(errors))
                return

            if errors:
                self._set_status(
                    f"{len(written)} report(s) saved. {len(errors)} did not.",
                    "warn")
                messagebox.showwarning(
                    "Some reports failed",
                    f"{len(written)} saved.\n\nThese did not:\n" + "\n".join(errors))
            else:
                word = "report" if len(written) == 1 else "reports"
                where = (f"{out_dir} and {folders - 1} other folder"
                         + ("" if folders == 2 else "s")
                         if folders > 1 else out_dir)
                self._set_status(f"{len(written)} {word} saved to {where}", "ok")

            self.open_link.pack(side="left", padx=(10, 0))

            if len(written) == 1:
                self._open_path(written[0])

        def _set_busy(self, busy, message=None):
            self.busy = busy
            self.generate_btn.config(state="disabled" if busy else "normal",
                                     text="Working on it..." if busy
                                     else "Generate reports")
            for button in (self.remove_btn, self.clear_btn):
                button.config(state="disabled" if busy else "normal")

            # The progress bar is only on the screen while something is moving
            # along it. Left there empty the rest of the time, it stops being
            # read as progress at all.
            if busy:
                self.generate_hint.config(text="")
                self.progress.pack(side="bottom", fill="x", padx=SPACE["lg"],
                                   pady=(SPACE["sm"], 0),
                                   before=self.status_frame)
            else:
                self.progress.pack_forget()

            if message:
                self._set_status(message)
            if not busy:
                self._update_generate_button()
                self._refresh_file_list()

        # ══════════════════════════════════════════════════════════════════════
        # SMALL DIALOGS
        # ══════════════════════════════════════════════════════════════════════

        def _edit_lab_details(self):
            """The laboratory name and the signature line at the foot of a report."""
            dialog = self._dialog("Laboratory details", 460)

            body = tk.Frame(dialog, bg=UI["panel"], padx=18, pady=16)
            body.pack(fill="both", padx=16, pady=(0, 10))
            body.columnconfigure(1, weight=1)

            rows = [("Laboratory name", self.lab_name_var,
                     "Prints centred at the top of every page"),
                    ("Signed by", self.signatory_var,
                     "Prints at the bottom right of every page")]

            for index, (label, var, hint) in enumerate(rows):
                ttk.Label(body, text=label, style="Field.TLabel").grid(
                    row=index * 2, column=0, sticky="w", pady=(0 if not index else 12, 2))
                entry = ttk.Entry(body, textvariable=var, style="Field.TEntry",
                                  font=ui_font("body"), width=44)
                entry.grid(row=index * 2, column=1, sticky="ew",
                           padx=(12, 0), pady=(0 if not index else 12, 2))
                ttk.Label(body, text=hint, style="Hint.TLabel").grid(
                    row=index * 2 + 1, column=1, sticky="w", padx=(12, 0))

            buttons = tk.Frame(dialog, bg=UI["bg"])
            buttons.pack(fill="x", padx=16, pady=(0, 14))
            ttk.Button(buttons, text="Done", style="Primary.TButton",
                       command=lambda: (self._save_settings(),
                                        dialog.destroy())).pack(side="right")
            dialog.bind("<Return>", lambda _e: (self._save_settings(),
                                                dialog.destroy()))
            self._show_dialog(dialog)

        def _show_intervals(self):
            """Lists every reference interval the reports are printed against."""
            dialog = self._dialog("Reference intervals", 640)

            body = tk.Frame(dialog, bg=UI["panel"], padx=16, pady=14)
            body.pack(fill="both", expand=True, padx=16, pady=(0, 10))

            ttk.Label(body, wraplength=580, justify="left", style="Hint.TLabel",
                      text="Every report is printed against these. Where it "
                           "says \u201cfrom the analyser\u201d, LabScribe prints "
                           "whatever interval the analyser gave for that "
                           "patient and only falls back to the one shown "
                           "here if the analyser gave none."
                      ).pack(anchor="w", pady=(0, 10))

            table_wrap = tk.Frame(body, bg=UI["panel"])
            table_wrap.pack(fill="both", expand=True)

            tree = ttk.Treeview(table_wrap, style="Results.Treeview",
                                columns=("test", "unit", "ref", "source"),
                                show="headings", height=16, selectmode="none")
            for key, title, width, anchor in [
                    ("test", "INVESTIGATION", 210, "w"),
                    ("unit", "UNIT", 80, "w"),
                    ("ref", "INTERVAL", 210, "w"),
                    ("source", "", 100, "w")]:
                tree.heading(key, text=title, anchor=anchor)
                tree.column(key, width=width, anchor=anchor,
                            stretch=(key == "ref"))
            tree.tag_configure("odd", background=UI["row_alt"])
            tree.tag_configure("head", foreground=UI["accent"])

            line = 0
            for section in SECTION_ORDER:
                tests = [t for t in TEST_LIBRARY if t["panel"] == section]
                if not tests:
                    continue
                tree.insert("", "end", tags=("head",),
                            values=(SECTIONS.get(section, {}).get("short", section),
                                    "", "", ""))
                for test in tests:
                    tags = ("odd",) if line % 2 else ()
                    tree.insert("", "end", tags=tags, values=(
                        "   " + _plain(test["name"])[:40],
                        _plain(test["unit"]),
                        _plain(" / ".join(test["ref"]))[:46],
                        "from the analyser" if test.get("source_ref") else ""))
                    line += 1

            tree.pack(side="left", fill="both", expand=True)
            scroll = ttk.Scrollbar(table_wrap, orient="vertical",
                                   style="Slim.Vertical.TScrollbar",
                                   command=tree.yview)
            scroll.pack(side="right", fill="y")
            tree.configure(yscrollcommand=scroll.set)

            ttk.Label(body, wraplength=580, justify="left", style="Hint.TLabel",
                      text="To change one, edit the TEST_LIBRARY list in the "
                           "LabScribe source file and build the program "
                           "again."
                      ).pack(anchor="w", pady=(10, 0))

            buttons = tk.Frame(dialog, bg=UI["bg"])
            buttons.pack(fill="x", padx=16, pady=(0, 14))
            ttk.Button(buttons, text="Close", style="Primary.TButton",
                       command=dialog.destroy).pack(side="right")
            self._show_dialog(dialog)

        # ── Changing how the window looks ─────────────────────────────────────

        def _set_style(self, style_id):
            """Switches style, remembers it, and redraws the window."""
            if self.busy:
                self._set_status("Hold on, LabScribe is still working.", "warn")
                return
            self.style_id = style_id
            apply_style(style_id)
            self._save_settings()
            # A moment's wait, so the menu this was chosen from has closed
            # itself before the window it belongs to is taken apart.
            self.after(60, self._finish_style_change)

        def _finish_style_change(self):
            self._redraw()
            self._set_status(
                f"Now showing in {style_by_id(self.style_id)['name']}.", "ok")

        def _redraw(self):
            """
            Takes the window apart and builds it again in the style now in
            force. The files already added and everything read from them are
            untouched, because none of that lives in the widgets.
            """
            for child in list(self.winfo_children()):
                try:
                    child.destroy()
                except Exception:
                    pass

            self.configure(bg=UI["bg"])
            self._build_styles()
            self._build_menu()
            self._build_layout()
            self._refresh_file_list()

            index = self.current
            if index is not None and index < len(self.jobs):
                self._select_row(index)
            else:
                self._show_job(None)

        def _choose_style(self):
            """The six styles laid out together, so one can be compared."""
            dialog = self._dialog("Appearance", 620)
            chosen = tk.StringVar(value=self.style_id)
            picked = {"id": None}

            body = tk.Frame(dialog, bg=UI["panel"], padx=16, pady=12)
            body.pack(fill="both", expand=True, padx=16, pady=(0, 10))

            ttk.Label(body, wraplength=560, justify="left", style="Hint.TLabel",
                      text="Pick whichever one you would rather look at all "
                           "day. It changes the window only. Nothing on the "
                           "printed report moves.").pack(anchor="w",
                                                         pady=(0, 10))

            for style in STYLES:
                row = tk.Frame(body, bg=UI["panel"])
                row.pack(fill="x", pady=3)

                button = tk.Radiobutton(
                    row, text=style["name"], value=style["id"],
                    variable=chosen, bg=UI["panel"], fg=UI["ink"],
                    activebackground=UI["panel"], selectcolor=UI["field"],
                    highlightthickness=0, anchor="w", width=16,
                    font=(pick_font(style["fonts"]), 10, "bold"))
                button.pack(side="left")

                swatch = tk.Frame(row, bg=UI["line"])
                swatch.pack(side="left", padx=(0, 12))
                for key in ("accent", "bg", "panel", "ink", "danger"):
                    tk.Frame(swatch, bg=style[key], width=18, height=18
                             ).pack(side="left", padx=1, pady=1)

                tk.Label(row, text=style["note"], bg=UI["panel"],
                         fg=UI["ink_soft"], font=ui_font("tiny"),
                         justify="left", wraplength=330, anchor="w"
                         ).pack(side="left", fill="x", expand=True)

            def use_it():
                picked["id"] = chosen.get()
                dialog.destroy()

            buttons = tk.Frame(dialog, bg=UI["bg"])
            buttons.pack(fill="x", padx=16, pady=(0, 14))
            ttk.Button(buttons, text="Use this style", style="Primary.TButton",
                       command=use_it).pack(side="right")
            ttk.Button(buttons, text="Leave it as it is", style="Ghost.TButton",
                       command=dialog.destroy).pack(side="right", padx=(0, 8))

            self._show_dialog(dialog)

            if picked["id"] and picked["id"] != self.style_id:
                self._set_style(picked["id"])

        def _show_about(self):
            dialog = self._dialog("About", 470)
            body = tk.Frame(dialog, bg=UI["panel"], padx=SPACE["lg"],
                            pady=SPACE["lg"])
            body.pack(fill="both", expand=True, padx=SPACE["lg"],
                      pady=(0, SPACE["sm"]))

            tk.Label(body, text=APP_NAME, bg=UI["panel"], fg=UI["accent"],
                     font=(UI_FONT, 18, "bold")).pack(anchor="w")
            tk.Label(body, text=f"{APP_TAGLINE}\nVersion {APP_VERSION}",
                     bg=UI["panel"], fg=UI["ink_soft"], font=ui_font("small"),
                     justify="left").pack(anchor="w", pady=(2, 12))

            ttk.Label(body, wraplength=410, justify="left", style="Body.TLabel",
                      text="LabScribe reads the PDF the chemistry analyser "
                           "prints and turns it into a finished patient "
                           "report: renal, liver, lipid, glycated haemoglobin "
                           "and uric acid, with the eGFR worked out from the "
                           "creatinine. It was written for the bench, so what "
                           "was read off the analyser is on the screen to be "
                           "checked before anything reaches paper."
                      ).pack(anchor="w")

            # Whether dragging works, and when it does not, why. Without this
            # a packaged build that lost its drag-and-drop library just looks
            # like a program that has stopped accepting dropped files.
            tk.Frame(body, bg=UI["line_soft"], height=1).pack(fill="x",
                                                              pady=(14, 12))
            if DND_AVAILABLE:
                ttk.Label(body, style="Hint.TLabel",
                          text="Dragging files onto the window: working."
                          ).pack(anchor="w")
            else:
                tk.Label(body, text="Dragging files onto the window is off",
                         bg=UI["panel"], fg=UI["warn"],
                         font=ui_font("small_bold")).pack(anchor="w")
                ttk.Label(body, style="Hint.TLabel", wraplength=410,
                          justify="left",
                          text="Use Add... instead. To switch dragging back "
                               "on, the tkinterdnd2 package has to be "
                               "installed and, in a packaged build, bundled "
                               "with the program.").pack(anchor="w",
                                                         pady=(1, 0))
                if DND_PROBLEM:
                    ttk.Label(body, style="Hint.TLabel", wraplength=410,
                              justify="left",
                              text=f"What it reported: {DND_PROBLEM}").pack(
                                  anchor="w", pady=(4, 0))

            tk.Frame(body, bg=UI["line_soft"], height=1).pack(fill="x",
                                                              pady=(14, 12))

            tk.Label(body, text=f"Built by {APP_PUBLISHER}", bg=UI["panel"],
                     fg=UI["ink"], font=ui_font("body_bold")
                     ).pack(anchor="w")
            ttk.Label(body, style="Hint.TLabel",
                      text="Questions, faults and requests are welcome."
                      ).pack(anchor="w", pady=(1, 8))

            contact = tk.Frame(body, bg=UI["panel"])
            contact.pack(anchor="w", fill="x")
            contact.columnconfigure(1, weight=1)

            for index, (label, value) in enumerate(
                    [("Email", APP_EMAIL), ("Telephone", APP_PHONE)]):
                tk.Label(contact, text=label, bg=UI["panel"],
                         fg=UI["ink_soft"], font=ui_font("small"), anchor="w",
                         width=10).grid(row=index, column=0, sticky="w",
                                        pady=2)
                # A read-only box rather than a label, so the address and the
                # number can be selected and copied.
                box = tk.Entry(contact, bg=UI["panel"], fg=UI["accent"],
                               font=ui_font("small"), relief="flat",
                               borderwidth=0, readonlybackground=UI["panel"],
                               highlightthickness=0)
                box.insert(0, value)
                box.config(state="readonly")
                box.grid(row=index, column=1, sticky="ew", pady=2)

            buttons = tk.Frame(dialog, bg=UI["bg"])
            buttons.pack(fill="x", padx=16, pady=(0, 14))
            ttk.Button(buttons, text="Copy email", style="Ghost.TButton",
                       command=lambda: self._copy(APP_EMAIL)).pack(side="left")
            ttk.Button(buttons, text="Close", style="Primary.TButton",
                       command=dialog.destroy).pack(side="right")
            self._show_dialog(dialog)

        def _copy(self, text):
            """Puts a piece of text on the clipboard."""
            try:
                self.clipboard_clear()
                self.clipboard_append(text)
                self._set_status(f"Copied {text}", "ok")
            except Exception:
                pass

        def _dialog(self, title, width):
            dialog = tk.Toplevel(self)
            dialog.title(title)
            dialog.configure(bg=UI["bg"])
            dialog.transient(self)
            dialog.resizable(False, False)
            tk.Frame(dialog, bg=UI["accent"], height=4, width=width).pack(fill="x")
            tk.Label(dialog, text=title, bg=UI["bg"], fg=UI["ink"],
                     font=ui_font("title"), anchor="w",
                     padx=SPACE["lg"], pady=SPACE["sm"]).pack(fill="x")
            return dialog

        def _show_dialog(self, dialog):
            dialog.update_idletasks()
            width, height = dialog.winfo_width(), dialog.winfo_height()

            x = self.winfo_rootx() + (self.winfo_width() - width) // 2
            y = self.winfo_rooty() + (self.winfo_height() - height) // 3

            # Kept on the screen whatever its height. A tall box placed at a
            # fixed offset down the window ended with its Close button below
            # the bottom of the screen, which is how a dialog traps someone.
            x = max(0, min(x, self.winfo_screenwidth() - width))
            y = max(0, min(y, self.winfo_screenheight() - height - 40))

            dialog.geometry(f"+{x}+{y}")
            dialog.grab_set()
            self.wait_window(dialog)

        # ══════════════════════════════════════════════════════════════════════
        # ODDS AND ENDS
        # ══════════════════════════════════════════════════════════════════════

        def _on_beside_toggled(self):
            """
            The folder box is only live when a folder is actually being used.
            Left enabled while reports were going somewhere else entirely, it
            would be saying something untrue.
            """
            self._apply_beside_state()
            self._save_settings()

        def _apply_beside_state(self):
            beside = bool(self.beside_var.get())
            state = "disabled" if beside else "normal"
            for widget in (getattr(self, "output_entry", None),
                           getattr(self, "browse_btn", None)):
                if widget is not None:
                    widget.config(state=state)

            if beside and getattr(self, "output_entry", None) is not None:
                folder = ""
                ready = self._ready_jobs()
                if ready:
                    folder = os.path.dirname(os.path.abspath(ready[0]["path"]))
                self.output_var.set(folder or self.output_var.get())

        def _browse_output(self):
            folder = filedialog.askdirectory(title="Where should the reports go?")
            if folder:
                self.output_var.set(folder)
                self._save_settings()

        def _open_output_folder(self):
            # Where the last run actually landed, which with "save beside the
            # analyser files" is not necessarily what is in the box.
            folder = (getattr(self, "_last_output_dir", "")
                      or self.output_var.get().strip())
            if folder and os.path.isdir(folder):
                self._open_path(folder)

        def _open_path(self, path):
            """Opens a file or folder in whatever the computer normally uses."""
            try:
                if sys.platform == "win32":
                    os.startfile(path)                       # noqa: S606
                elif sys.platform == "darwin":
                    import subprocess
                    subprocess.Popen(["open", path])
                else:
                    import subprocess
                    subprocess.Popen(["xdg-open", path])
            except Exception:
                pass

        def _set_status(self, message, tone="info"):
            colours = {"info": UI["ink_soft"], "ok": UI["accent"],
                       "warn": UI["warn"], "danger": UI["danger"]}
            changed = message != self.status_var.get()
            self.status_var.set(message)
            self.status_label.config(fg=colours.get(tone, UI["ink_soft"]))

            if not changed:
                return

            # One tint, held briefly, then gone. Not an animation — there is no
            # compositor behind a tkinter window and a hand-rolled fade only
            # flickers. This is the smallest thing that makes a change to the
            # line impossible to miss, and it is the only thing that moves.
            if self._status_flash is not None:
                try:
                    self.after_cancel(self._status_flash)
                except Exception:
                    pass
            self.status_label.config(bg=UI["accent_tint"])
            self._status_flash = self.after(
                700, lambda: self.status_label.config(bg=UI["bg"]))

        # ── Settings, kept in the user's AppData folder ───────────────────────

        def _settings_path(self):
            return settings_path()

        def _save_settings(self):
            try:
                with open(self._settings_path(), "w", encoding="utf-8") as handle:
                    handle.write(f"lab_name={self.lab_name_var.get().strip()}\n")
                    handle.write(f"signatory={self.signatory_var.get().strip()}\n")
                    handle.write(f"output_dir={self.output_var.get().strip()}\n")
                    handle.write(f"flags={int(bool(self.flags_var.get()))}\n")
                    handle.write(f"style={self.style_id}\n")
                    handle.write(f"beside={int(bool(self.beside_var.get()))}\n")
                    handle.write(f"geometry={self.winfo_geometry()}\n")
            except Exception:
                pass

        def _load_settings(self):
            try:
                saved = read_settings()
                if saved.get("lab_name"):
                    self.lab_name_var.set(saved["lab_name"])
                if saved.get("signatory"):
                    self.signatory_var.set(saved["signatory"])
                if saved.get("output_dir"):
                    self.output_var.set(saved["output_dir"])
                if saved.get("flags") in ("0", "1"):
                    self.flags_var.set(saved["flags"] == "1")
                if saved.get("beside") in ("0", "1"):
                    self.beside_var.set(saved["beside"] == "1")
                self._apply_beside_state()
            except Exception:
                pass

    return App


# ==============================================================================
# SECTION 16 — WHAT HAPPENS WHEN YOU RUN THE FILE
#
#   Double-click LabScribe            →  the window opens
#   LabScribe --demo                  →  a sample report, no analyser needed
#   LabScribe "file.pdf"              →  a real report, straight to a PDF
# ==============================================================================

if __name__ == "__main__":

    if len(sys.argv) > 1:
        # Something was typed after the file name — run in command-line mode
        main()

    else:
        # Nothing typed — open the window
        if not TK_AVAILABLE:
            print("tkinter is not available on this computer, so the window")
            print("cannot open. The command line still works, for example:")
            print("    LabScribe.exe --demo")
            sys.exit(0)

        # Check the two libraries the script cannot do without
        missing = []
        try:
            import reportlab           # noqa: F401  (imported only as a test)
        except ImportError:
            missing.append("reportlab")
        try:
            import pdfplumber          # noqa: F401
        except ImportError:
            missing.append("pdfplumber")

        if missing:
            message = ("Please install the missing package(s) first:\n\n"
                       "    pip install " + " ".join(missing) + "\n\n"
                       "Then run the script again.")
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Missing Packages", message)
                root.destroy()
            except Exception:
                print(message)
            sys.exit(1)

        if not DND_AVAILABLE:
            print("Drag-and-drop is off. To switch it on: pip install tkinterdnd2")

        AppClass = _build_gui_app()
        app = AppClass()
        app.mainloop()
