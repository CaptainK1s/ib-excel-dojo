#!/usr/bin/env python3
"""Build IB Excel Dojo — practice workbook for investment-banking Excel."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, Protection
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice

wb = Workbook()

# ── palette ────────────────────────────────────────────────────────────────
NAVY = "1F4E79"
NAVY2 = "2E75B6"
STEEL = "5B9BD5"
PALE = "D6E3F0"
SECTION = "D6DCE4"
YELLOW = "FFF2CC"
YELLOW_HOT = "FFE699"
GREEN_BG = "C6EFCE"
RED_BG = "FFC7CE"
AMBER_BG = "FFEB9C"
WHITE = "FFFFFF"
OFFWHITE = "F8F9FA"
GRAY_LINE = "BDD7EE"
INPUT_YELLOW = "FFFF99"
LIGHT_GRAY = "F2F2F2"
ORANGE = "F4B183"
PURPLE_BG = "E2D5F1"
DARK = "1A1A1A"

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
bottom_d = Border(bottom=Side(style="thin", color="1F4E79"))
bottom_double = Border(bottom=Side(style="double", color="000000"))
bottom_single = Border(bottom=Side(style="thin", color="000000"))
top_bottom = Border(
    top=Side(style="thin", color="000000"),
    bottom=Side(style="double", color="000000"),
)

font_title = Font(name="Arial", size=18, bold=True, color=WHITE)
font_h1 = Font(name="Arial", size=14, bold=True, color=WHITE)
font_h2 = Font(name="Arial", size=11, bold=True, color=NAVY)
font_h3 = Font(name="Arial", size=10, bold=True, color=NAVY)
font_body = Font(name="Arial", size=10, color="000000")
font_small = Font(name="Arial", size=9, color="595959")
font_input = Font(name="Arial", size=10, color="0000FF")
font_formula = Font(name="Arial", size=10, color="000000")
font_link = Font(name="Arial", size=10, color="008000")
font_white_b = Font(name="Arial", size=10, bold=True, color=WHITE)
font_navy_b = Font(name="Arial", size=10, bold=True, color=NAVY)
font_label = Font(name="Arial", size=10, color="000000")
font_section = Font(name="Arial", size=10, bold=True, color="000000")
font_hint = Font(name="Arial", size=9, italic=True, color="808080")
font_kbd = Font(name="Consolas", size=9, bold=True, color="1F4E79")
font_red = Font(name="Arial", size=10, color="C00000")
font_ok = Font(name="Arial", size=10, bold=True, color="006600")

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_navy2 = PatternFill("solid", fgColor=NAVY2)
fill_pale = PatternFill("solid", fgColor=PALE)
fill_section = PatternFill("solid", fgColor=SECTION)
fill_yellow = PatternFill("solid", fgColor=YELLOW)
fill_input = PatternFill("solid", fgColor=INPUT_YELLOW)
fill_green = PatternFill("solid", fgColor=GREEN_BG)
fill_red = PatternFill("solid", fgColor=RED_BG)
fill_amber = PatternFill("solid", fgColor=AMBER_BG)
fill_off = PatternFill("solid", fgColor=OFFWHITE)
fill_gray = PatternFill("solid", fgColor=LIGHT_GRAY)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_orange = PatternFill("solid", fgColor=ORANGE)
fill_purple = PatternFill("solid", fgColor=PURPLE_BG)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

FMT_MM = '$#,##0;($#,##0);"-"'
FMT_MM1 = '$#,##0.0;($#,##0.0);"-"'
FMT_NUM = '#,##0;(#,##0);"-"'
FMT_NUM1 = '#,##0.0;(#,##0.0);"-"'
FMT_PCT = "0.0%"
FMT_PCT2 = "0.00%"
FMT_X = '0.0x;(0.0x);"-"'
FMT_BPS = "0"

def apply_range_font(ws, cells, font):
    for addr in cells:
        ws[addr].font = font

def width(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w

def fill_row(ws, row, start, end, fill):
    for col in range(start, end + 1):
        ws.cell(row, col).fill = fill

def font_row(ws, row, start, end, font):
    for col in range(start, end + 1):
        ws.cell(row, col).font = font

def align_row(ws, row, start, end, al):
    for col in range(start, end + 1):
        ws.cell(row, col).alignment = al

def header_bar(ws, row, start, end, text, fill=fill_navy, font=font_h1):
    ws.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    c = ws.cell(row, start, text)
    c.fill = fill
    c.font = font
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    fill_row(ws, row, start, end, fill)
    ws.row_dimensions[row].height = 24

def section_bar(ws, row, start, end, text):
    ws.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    c = ws.cell(row, start, text)
    c.fill = fill_section
    c.font = font_h3
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    fill_row(ws, row, start, end, fill_section)
    ws.row_dimensions[row].height = 18

def banner(ws, title, subtitle, cols=10):
    header_bar(ws, 1, 1, cols, title)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    c = ws.cell(2, 1, subtitle)
    c.fill = fill_navy2
    c.font = Font(name="Arial", size=10, italic=True, color=WHITE)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    fill_row(ws, 2, 1, cols, fill_navy2)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "IB Excel Dojo"
    ws.oddFooter.right.text = "Page &P of &N"

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 00 Start
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.active
ws.title = "00_Start"
banner(ws, "IB EXCEL DOJO", "Practice workbook — model like a first-year investment banking analyst", 8)
width(ws, {"A": 3, "B": 38, "C": 28, "D": 28, "E": 28, "F": 28, "G": 22, "H": 18})

ws.merge_cells("B4:H4")
ws["B4"] = (
    "This file is a gym, not a lecture. Work left-to-right across the tabs. "
    "Blue font + yellow fill = you type. Black = formula. Green = link from another sheet."
)
ws["B4"].font = font_body
ws["B4"].alignment = left

# How to use
section_bar(ws, 6, 2, 8, "HOW TO USE THIS FILE")
steps = [
    ("1", "01_Shortcuts", "Print it or keep it open. Drill the boxed 'Week 1' shortcuts until they are muscle memory. No mouse."),
    ("2", "02_Format_Gym", "A deliberately ugly tab. Rebuild it to bank standard using only keyboard shortcuts. Check against the style rules."),
    ("3", "03_Formula_Dojo", "20 live drills: INDEX/MATCH, XLOOKUP-style lookups, SUMIFS, CAGR, IRR, flags, error traps."),
    ("4", "04_3Statement", "Mini 3-statement model for Apex Industrial. Change yellow assumptions and watch IS / BS / CFS and checks move."),
    ("5", "05_Sensitivity", "Two-way WACC × g table and operating-case toggle. This is how decks get 'sensitized' at 2 a.m."),
    ("6", "06_Comps", "Trading comps math: equity value → enterprise value → multiples → football-field stats."),
    ("7", "07_Checks", "The page a VP actually looks at. Every model needs a check dashboard that is impossible to miss."),
    ("8", "08_Plan", "14-day practice plan. 25–40 minutes a day beats a 6-hour cram Sunday night."),
]
ws["B7"] = "#"
ws["C7"] = "TAB"
ws["D7"] = "WHAT YOU DO"
ws.merge_cells("D7:H7")
for col in range(2, 9):
    ws.cell(7, col).font = font_white_b
    ws.cell(7, col).fill = fill_navy2
    ws.cell(7, col).alignment = center
ws["D7"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

for i, (n, tab, desc) in enumerate(steps, start=8):
    ws.cell(i, 2, n).font = font_navy_b
    ws.cell(i, 2).alignment = center
    ws.cell(i, 2).fill = fill_pale
    ws.cell(i, 3, tab).font = font_navy_b
    ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=8)
    ws.cell(i, 4, desc).font = font_body
    ws.row_dimensions[i].height = 28
    if i % 2 == 0:
        fill_row(ws, i, 3, 8, fill_off)

# Color legend
section_bar(ws, 17, 2, 8, "COLOR LANGUAGE (MEMORIZE THIS — EVERY BANK USES A VERSION OF IT)")
legend = [
    (18, "Blue font  RGB 0,0,255", "Hardcoded inputs. Historicals you typed. Growth rates. Tax rate. Anything a human decided.", font_input, fill_input),
    (19, "Black font  RGB 0,0,0", "Formulas. If you can write it as a formula, it is black. Never hardcode a calculated number.", font_formula, fill_white),
    (20, "Green font  RGB 0,128,0", "Cross-sheet links (Sheet!A1). Tells a reviewer 'this number lives somewhere else.'", font_link, fill_white),
    (21, "Yellow fill", "Key assumptions the user is supposed to toggle. Blue font + yellow fill = the control panel.", font_input, fill_input),
    (22, "Red / green check cells", "Model integrity. BS must balance. CFS ending cash must equal BS cash. Never ship without these.", font_formula, fill_green),
]
ws["B18"].font = font_input
for row, label, desc, fnt, fl in legend:
    ws.cell(row, 2, label).font = fnt
    ws.cell(row, 2).fill = fl
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)
    ws.cell(row, 3, desc).font = font_body
    ws.row_dimensions[row].height = 22

# House rules
section_bar(ws, 24, 2, 8, "HOUSE RULES OF A LIVE DEAL MODEL")
rules = [
    "One formula, copied across. The FY26 formula must drag to FY30 with F4 anchors — never rewrite year by year.",
    "Assumptions live in one place. Formulas reference cells. Nobody should hunt through 40 tabs for the tax rate.",
    "Units in the header, not the cell. 'Revenue ($mm)' — then the number is 510, not 510,000,000.",
    "Negatives in parentheses. Custom format  $#,##0;($#,##0);\"-\"  so zeros show as dashes.",
    "No merged cells in the number grid. Use Center Across Selection (Format Cells → Alignment) if you must span a title.",
    "Error checks on their own tab AND a visible 'BALANCES / BREAKS' flag at the top of every output sheet.",
    "Historical = blue. Forecast = black formulas. The color change is the fact / forecast boundary.",
    "Never leave a #DIV/0! or #REF! in a file that someone else will open. Wrap risky divisions in IF or IFERROR.",
    "Comments (Shift+F2) on any hardcoded assumption that came from a 10-K, CIM, or research note. Source it.",
    "Save versions. Apex_v3_pre-IC.xlsx. The file name is part of the control system.",
]
for i, r in enumerate(rules, start=25):
    ws.cell(i, 2, "▸").font = font_navy_b
    ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
    ws.cell(i, 3, r).font = font_body
    ws.row_dimensions[i].height = 20

section_bar(ws, 36, 2, 8, "THE MINDSET")
ws.merge_cells("B37:H39")
ws["B37"] = (
    "Investment bankers are not faster at Excel because they know obscure functions. "
    "They are faster because they never touch the mouse, they never retype a formula that can be copied, "
    "and they build every file so a VP can audit it in 90 seconds: blue is input, black is math, "
    "green is a link, yellow is what we are toggling on this call. "
    "That is the whole sport. The rest is reps."
)
ws["B37"].font = Font(name="Arial", size=10, italic=True, color="333333")
ws["B37"].alignment = left_top
ws.row_dimensions[37].height = 20
ws.row_dimensions[38].height = 20
ws.row_dimensions[39].height = 20

ws["B41"] = "Built for practice. Numbers are hypothetical. Not advice, not a live model."
ws["B41"].font = font_hint
ws.merge_cells("B41:H41")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 01 Shortcuts
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("01_Shortcuts")
banner(ws, "01  ·  SHORTCUT BIBLE", "Windows-first (what banks actually use). Mac notes in the last column. Drill the starred rows first.", 8)
width(ws, {"A": 3, "B": 28, "C": 26, "D": 22, "E": 38, "F": 16, "G": 22, "H": 28})

headers = ["GROUP", "ACTION", "WINDOWS", "WHY BANKERS CARE", "TIER", "MAC", "DRILL"]
for i, h in enumerate(headers, start=2):
    c = ws.cell(4, i, h)
    c.font = font_white_b
    c.fill = fill_navy2
    c.alignment = center
ws.row_dimensions[4].height = 20
ws.auto_filter.ref = "B4:H80"
ws.freeze_panes = "A5"

# (group, action, win, why, tier, mac, drill)
shorts = [
    ("NAV", "Jump to edge of block", "Ctrl + Arrow", "Cross a 400-row IS in one keystroke", "★ WEEK 1", "Ctrl + Arrow", "Stand in C20, jump to last number"),
    ("NAV", "Select to edge of block", "Ctrl + Shift + Arrow", "Highlight an entire forecast row to copy", "★ WEEK 1", "Ctrl + Shift + Arrow", "Select FY26–FY30 without the mouse"),
    ("NAV", "Jump to cell / Go To", "F5  or  Ctrl + G", "Type a cell name, land on it", "★ WEEK 1", "F5 / Ctrl + G", "F5 → Checks!C6"),
    ("NAV", "Move between sheets", "Ctrl + PgUp / PgDn", "Tab-hop a 12-sheet model", "★ WEEK 1", "Fn + Ctrl + Up/Down (varies)", "Loop 00 → 08 and back"),
    ("NAV", "Select whole row / col", "Shift + Space  /  Ctrl + Space", "Insert, delete, or format a line item", "WEEK 1", "Shift + Space / Ctrl + Space", "Select row 40, insert one row"),
    ("NAV", "Freeze panes", "Alt W F F", "Keep years visible while you scroll", "WEEK 1", "Layout or Alt sequence", "Freeze below year headers"),
    ("EDIT", "Edit in cell", "F2", "See the formula without the mouse", "★ WEEK 1", "F2 / Ctrl + U", "F2 a formula, read the colors"),
    ("EDIT", "Toggle $ absolute ref", "F4", "Lock the tax-rate cell as you copy right", "★ WEEK 1", "Cmd + T  (Excel Mac)", "Write =B10*(1+$B$5) using only F4"),
    ("EDIT", "Fill down / fill right", "Ctrl + D  /  Ctrl + R", "Replicate a formula across years in 1 beat", "★ WEEK 1", "Ctrl + D / Ctrl + R", "Write FY26, Ctrl+R across FY27–30"),
    ("EDIT", "AutoSum", "Alt + =", "Totals without touching the ribbon", "★ WEEK 1", "Cmd + Shift + T", "Total an opex block"),
    ("EDIT", "Paste Special", "Ctrl + Alt + V   (or Alt E S)", "Values, formulas, formats, add, subtract, skip blanks", "★ WEEK 1", "Ctrl + Cmd + V", "Paste Values over a live link"),
    ("EDIT", "Paste Values only", "Alt E S V  Enter", "Kill a formula, keep the number", "★ WEEK 1", "Ctrl + Cmd + V then V", "Snapshot a case"),
    ("EDIT", "Repeat last action", "F4  (not in edit mode)", "Apply the same border 40 times", "WEEK 1", "Cmd + Y", "Border one cell, F4 down the column"),
    ("EDIT", "Insert comment / note", "Shift + F2", "Source the growth rate ('10-K p.44')", "WEEK 1", "Shift + F2", "Source one yellow cell"),
    ("EDIT", "Undo / Redo", "Ctrl + Z  /  Ctrl + Y", "You will need this. A lot.", "WEEK 1", "Cmd + Z / Cmd + Y", "Break something, undo it"),
    ("FORMAT", "Format Cells dialog", "Ctrl + 1", "Custom formats live here. Home of $#,##0;( );\"-\"", "★ WEEK 1", "Cmd + 1", "Build the IB number format"),
    ("FORMAT", "Number w/ thousands", "Ctrl + Shift + !", "Quick comma format", "WEEK 1", "Ctrl + Shift + !", "Format a raw block"),
    ("FORMAT", "Percent", "Ctrl + Shift + %", "Growth rates, margins, tax", "WEEK 1", "Ctrl + Shift + %", "Format the assumption block"),
    ("FORMAT", "Currency", "Ctrl + Shift + $", "Then customize decimals with Alt H 9 / 0", "WEEK 1", "Ctrl + Shift + $", "Apply, then drop decimals"),
    ("FORMAT", "Blue input font", "Alt H F C  → pick blue", "Or set a QAT button for RGB 0,0,255", "★ WEEK 2", "Home → Font Color", "Color every hardcoded cell blue"),
    ("FORMAT", "Fill color (yellow)", "Alt H H", "Mark assumption cells", "WEEK 2", "Home → Fill", "Yellow the driver block"),
    ("FORMAT", "Borders menu", "Alt H B", "Top/bottom for totals, outline for tables", "WEEK 2", "Home → Borders", "Total line: top thin + bottom double"),
    ("FORMAT", "Autofit column", "Alt H O I", "Stop #### from showing", "WEEK 2", "Double-click divider (or ribbon)", "Autofit labels column"),
    ("FORMAT", "Increase / decrease decimal", "Alt H 0  /  Alt H 9", "Multiples at 0.0x, rates at 0.0%", "WEEK 2", "Home → decimal buttons", "Set EV/EBITDA to one decimal"),
    ("AUDIT", "Show all formulas", "Ctrl + `  (backtick)", "VP audit mode. Whole sheet becomes formulas.", "★ WEEK 2", "Ctrl + `", "Toggle formulas on 04_3Statement"),
    ("AUDIT", "Jump to precedents", "Ctrl + [", "Land on the cells this formula uses", "★ WEEK 2", "Ctrl + [", "From NI, jump back to EBT"),
    ("AUDIT", "Jump to dependents", "Ctrl + ]", "Who uses this cell?", "WEEK 2", "Ctrl + ]", "From tax rate, find every user"),
    ("AUDIT", "Trace precedents arrows", "Alt M P", "Visual arrows for messy formulas", "WEEK 2", "Formulas → Trace", "Trace WACC"),
    ("AUDIT", "Go To Special → Constants", "F5 → Alt S → O → X", "Select every hardcoded number to paint blue", "★ WEEK 2", "F5 → Special → Constants", "Blue the whole IS historicals"),
    ("AUDIT", "Go To Special → Formulas", "F5 → Alt S → F → X", "Select every formula to paint black", "★ WEEK 2", "F5 → Special → Formulas", "Black the forecast block"),
    ("AUDIT", "Find '!' cross-sheet links", "Ctrl + F  then  !", "Then font-color them green", "WEEK 2", "Ctrl + F", "Green every Sheet! reference"),
    ("AUDIT", "Find & replace formats", "Ctrl + H → Options → Format", "Mass-paint after Go To Special", "WEEK 3", "Ctrl + H", "Replace font color across a block"),
    ("STRUCT", "Insert row / column", "Alt I R  /  Alt I C   or  Ctrl + +", "Keep structure clean, don't overwrite", "WEEK 2", "Ctrl + Shift + +", "Insert a 'memo' line under EBITDA"),
    ("STRUCT", "Delete row / column", "Alt H D R  /  Alt H D C   or  Ctrl + -", "Remove a dead line item", "WEEK 2", "Ctrl + -", "Delete an extra spacer row"),
    ("STRUCT", "Group rows (outline)", "Alt A G G   or  Shift + Alt + →", "Collapse monthly detail under annual", "WEEK 3", "Data → Group", "Group monthly stub under FY"),
    ("STRUCT", "Hide / unhide rows", "Ctrl + 9  /  Ctrl + Shift + 9", "Park supporting calcs", "WEEK 3", "Ctrl + 9", "Hide the working-capital rollforward"),
    ("STRUCT", "Name a range", "Ctrl + F3  or  Name Box", "TaxRate instead of Assumptions!C12", "WEEK 3", "Ctrl + F3", "Name the yellow tax-rate cell"),
    ("STRUCT", "Create names from labels", "Ctrl + Shift + F3", "Turn a driver block into named cells in one shot", "WEEK 3", "Ctrl + Shift + F3", "Name the whole assumption column"),
    ("CALC", "Calculate now (all)", "F9", "Force after iterative / data-table work", "WEEK 2", "F9", "Change a driver, F9"),
    ("CALC", "Calculate active sheet", "Shift + F9", "Faster on huge workbooks", "WEEK 3", "Shift + F9", "Calc only 05_Sensitivity"),
    ("CALC", "Evaluate formula", "Alt M V", "Step through a nested IF", "WEEK 3", "Formulas → Evaluate", "Evaluate the cash plug"),
    ("CALC", "Insert function dialog", "Shift + F3", "Browse args when you forget INDEX syntax", "WEEK 2", "Shift + F3", "Build INDEX/MATCH via dialog once"),
    ("DATA", "Start a Data Table", "Alt A W T", "Two-way sensitivity. Set calc to 'Automatic except data tables'.", "WEEK 3", "Data → What-If", "Rebuild the WACC × g table by hand"),
    ("DATA", "Goal Seek", "Alt A W G", "What growth gets you to $50mm NI?", "WEEK 3", "Data → What-If → Goal Seek", "Goal-seek FY28 NI to 50"),
    ("PRINT", "Page setup / print preview", "Ctrl + P  then  Alt P S", "Fit-to-width 1, landscape, print titles", "WEEK 3", "Cmd + P", "Make 04_3Statement print on 1 page wide"),
    ("PRINT", "Set print area", "Alt P R S", "Don't send the scratch columns to the MD", "WEEK 3", "Page Layout → Print Area", "Print area = output only"),
]

tier_fill = {
    "★ WEEK 1": fill_orange,
    "WEEK 1": fill_yellow,
    "★ WEEK 2": fill_pale,
    "WEEK 2": fill_off,
    "WEEK 3": fill_gray,
}

for i, row in enumerate(shorts, start=5):
    for j, val in enumerate(row, start=2):
        cell = ws.cell(i, j, val)
        cell.alignment = left if j != 6 else center
        cell.font = font_kbd if j in (4, 7) else font_body
        if j == 6:
            cell.fill = tier_fill.get(val, fill_white)
            cell.font = Font(name="Arial", size=8, bold=True, color=NAVY)
            cell.alignment = center
        if i % 2 == 0 and j != 6:
            cell.fill = fill_off
    ws.row_dimensions[i].height = 18

# QAT note
r = 5 + len(shorts) + 1
section_bar(ws, r, 2, 8, "QUICK ACCESS TOOLBAR — THE SECRET WEAPON")
ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 3, end_column=8)
ws.cell(r + 1, 2, (
    "File → Options → Quick Access Toolbar. Add: Paste Values, Paste Formats, Border menu, Font Color, "
    "Fill Color, Increase/Decrease Decimal, Merge & Center (so you can avoid it), and Name Manager. "
    "They become Alt + 1, Alt + 2, Alt + 3… Bankers customize this on day one of a new machine. "
    "Also: File → Options → Formulas → Enable iterative calculation (max 100, change 0.001) for interest circs; "
    "set workbook calc to Automatic except data tables so huge sens tables don't freeze Excel on every keystroke."
))
ws.cell(r + 1, 2).alignment = left_top
ws.cell(r + 1, 2).font = font_body

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 02 Format Gym
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("02_Format_Gym")
banner(ws, "02  ·  FORMAT GYM", "This tab is intentionally ugly. Your job: make the LEFT block look like the RIGHT 'after' block using only shortcuts.", 12)
width(ws, {c: 14 for c in "ABCDEFGHIJKL"})
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 28
ws.column_dimensions["H"].width = 3
ws.column_dimensions["I"].width = 28

ws.merge_cells("B4:F4")
ws["B4"] = "BEFORE  —  messy extract (fix this side)"
ws["B4"].font = Font(name="Arial", size=11, bold=True, color="C00000")
ws.merge_cells("I4:L4")
ws["I4"] = "AFTER  —  bank standard (target look)"
ws["I4"].font = Font(name="Arial", size=11, bold=True, color="006600")

# BEFORE block — ugly
ugly_years = [2023, 2024, 2025]
ws["B6"] = "apex industrial  income statement"
ws["B6"].font = Font(name="Comic Sans MS", size=14, bold=True, color="FF00FF")
ws["C6"] = 2023
ws["D6"] = 2024
ws["E6"] = 2025
ws["F6"] = "notes"
for col, val in [(3, 420000000), (4, 460000000), (5, 510000000)]:
    pass
ws["B7"] = "rev"
ws["C7"] = 420000000
ws["D7"] = 460000000
ws["E7"] = 510000000
ws["B8"] = "cogs"
ws["C8"] = -260400000
ws["D8"] = -282900000
ws["E8"] = -311100000
ws["B9"] = "gp"
ws["C9"] = 159600000
ws["D9"] = 177100000
ws["E9"] = 198900000
ws["B10"] = "opex"
ws["C10"] = -75600000
ws["D10"] = -81880000
ws["E10"] = -89250000
ws["B11"] = "ebitda"
ws["C11"] = 84000000
ws["D11"] = 95220000
ws["E11"] = 109650000
ws["B12"] = "da"
ws["C12"] = -18000000
ws["D12"] = -19000000
ws["E12"] = -21000000
ws["B13"] = "ebit"
ws["C13"] = 66000000
ws["D13"] = 76220000
ws["E13"] = 88650000
ws["B14"] = "int"
ws["C14"] = -8000000
ws["D14"] = -7500000
ws["E14"] = -7000000
ws["B15"] = "ebt"
ws["C15"] = 58000000
ws["D15"] = 68720000
ws["E15"] = 81650000
ws["B16"] = "tax 25%"
ws["C16"] = -14500000
ws["D16"] = -17180000
ws["E16"] = -20412500
ws["B17"] = "net income"
ws["C17"] = 43500000
ws["D17"] = 51540000
ws["E17"] = 61237500
for row in range(6, 18):
    for col in range(2, 7):
        ws.cell(row, col).font = Font(name="Comic Sans MS", size=9, color="333333")
        ws.cell(row, col).number_format = "0.00"
        if col > 2:
            ws.cell(row, col).alignment = Alignment(horizontal="left")

ws.merge_cells("B19:F26")
ws["B19"] = (
    "MISSION — do this with the keyboard, not the mouse:\n"
    "1. Autofit / set column widths.  2. Arial 10 everywhere.  3. Title case line items, indent subtotals.\n"
    "4. Year headers centered, formatted as text '2023' not '2,023'.  5. Convert units to $mm (divide by 1,000,000).\n"
    "6. IB number format  $#,##0;($#,##0);\"-\"  — negatives in parens, zeros as dashes.\n"
    "7. Blue font on every hardcoded number. Black font on every formula you write (GP, EBITDA, EBIT, EBT, Tax, NI).\n"
    "8. Yellow fill on the tax-rate assumption (pull 25% out into its own input cell).\n"
    "9. Top border + double bottom border on Net Income. Thin bottom border on GP / EBITDA / EBIT.\n"
    "10. Header bar navy + white. Units in the title: 'Income Statement ($mm)'. No merged cells in the number grid."
)
ws["B19"].alignment = left_top
ws["B19"].font = font_body
ws["B19"].fill = fill_amber

# AFTER block — correct look
ws["I6"] = "Income Statement ($mm)"
ws["I6"].font = font_white_b
ws["I6"].fill = fill_navy
ws["J6"] = "2023"
ws["K6"] = "2024"
ws["L6"] = "2025"
for col in range(9, 13):
    ws.cell(6, col).font = font_white_b
    ws.cell(6, col).fill = fill_navy
    ws.cell(6, col).alignment = center

after_rows = [
    (7, "Revenue", 420, 460, 510, True),
    (8, "  COGS", -260, -283, -311, True),
    (9, "Gross Profit", 160, 177, 199, False),
    (10, "  Operating expenses", -76, -82, -89, True),
    (11, "EBITDA", 84, 95, 110, False),
    (12, "  D&A", -18, -19, -21, True),
    (13, "EBIT", 66, 76, 89, False),
    (14, "  Interest expense", -8, -8, -7, True),
    (15, "EBT", 58, 69, 82, False),
    (16, "  Tax", -15, -17, -20, False),
    (17, "Net Income", 44, 52, 61, False),
]
# Use exact $mm rounded consistently with later model
after_exact = [
    (7, "Revenue", 420.0, 460.0, 510.0, True),
    (8, "  COGS", -260.4, -282.9, -311.1, True),
    (9, "Gross Profit", 159.6, 177.1, 198.9, False),
    (10, "  Operating expenses", -75.6, -81.9, -89.3, True),
    (11, "EBITDA", 84.0, 95.2, 109.7, False),
    (12, "  D&A", -18.0, -19.0, -21.0, True),
    (13, "EBIT", 66.0, 76.2, 88.7, False),
    (14, "  Interest expense", -8.0, -7.5, -7.0, True),
    (15, "EBT", 58.0, 68.7, 81.7, False),
    (16, "  Tax", -14.5, -17.2, -20.4, False),
    (17, "Net Income", 43.5, 51.5, 61.2, False),
]
for row, label, a, b, c, inp in after_exact:
    ws.cell(row, 9, label).font = font_section if label.strip() in (
        "Revenue", "Gross Profit", "EBITDA", "EBIT", "EBT", "Net Income"
    ) else font_body
    for col, val in [(10, a), (11, b), (12, c)]:
        cell = ws.cell(row, col, val)
        cell.number_format = FMT_MM1
        cell.alignment = center
        cell.font = font_input if inp else font_formula
    if label.strip() in ("Gross Profit", "EBITDA", "EBIT"):
        for col in range(9, 13):
            ws.cell(row, col).border = bottom_single
    if label.strip() == "Net Income":
        for col in range(9, 13):
            ws.cell(row, col).border = top_bottom
            ws.cell(row, col).font = Font(name="Arial", size=10, bold=True, color="0000FF" if inp else "000000")

ws["I19"] = "Tax rate (input)"
ws["I19"].font = font_body
ws["J19"] = 0.25
ws["J19"].font = font_input
ws["J19"].fill = fill_input
ws["J19"].number_format = FMT_PCT
ws["J19"].alignment = center
ws["K19"] = "← yellow + blue = assumption"
ws["K19"].font = font_hint
ws.merge_cells("K19:L19")

ws.merge_cells("I21:L26")
ws["I21"] = (
    "STYLE NOTES\n"
    "• Totals are bold; contra-accounts indented two spaces.\n"
    "• Historicals stay blue because they were typed from the 10-K.\n"
    "• Once you forecast a year, that column turns black (formulas).\n"
    "• Tax is a formula: =−EBT × tax_rate. That is why it is black even in history if you choose to formula-link it.\n"
    "• Some groups blue ALL history including calculated subtotals. House style varies. Be consistent inside one file."
)
ws["I21"].alignment = left_top
ws["I21"].font = font_small

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 03 Formula Dojo
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("03_Formula_Dojo")
banner(ws, "03  ·  FORMULA DOJO", "Type formulas in the yellow cells. Hidden answer key sits to the right of column N — unhide or scroll if you get stuck.", 14)
width(ws, {"A": 3, "B": 26, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 16, "I": 18, "J": 16, "K": 14, "L": 14, "M": 14, "N": 3})
ws.column_dimensions["O"].width = 18
ws.column_dimensions["P"].width = 42

# --- Source data block
section_bar(ws, 4, 2, 8, "SOURCE DATA  —  do not overwrite  (used by the drills below)")
headers = ["Company", "Ticker", "Sector", "EV ($mm)", "EBITDA ($mm)", "Net Income", "FY25 Rev", "FY24 Rev"]
for i, h in enumerate(headers, start=2):
    c = ws.cell(5, i, h)
    c.font = font_white_b
    c.fill = fill_navy2
    c.alignment = center

comps_data = [
    ("Apex Industrial", "APX", "Industrials", 980, 110, 61.2, 510, 460),
    ("Northline Mfg", "NLM", "Industrials", 1420, 155, 88.0, 720, 680),
    ("Cedar Forge", "CFR", "Materials", 760, 92, 41.5, 390, 370),
    ("Helix Components", "HLX", "Industrials", 2105, 240, 130.4, 980, 910),
    ("Pinnacle Tools", "PNT", "Industrials", 540, 61, 22.0, 250, 240),
    ("Riverbend Co", "RVB", "Materials", 1180, 150, 70.1, 640, 600),
]
for i, row in enumerate(comps_data, start=6):
    for j, val in enumerate(row, start=2):
        cell = ws.cell(i, j, val)
        cell.font = font_input if j >= 5 else font_body
        cell.alignment = center if j > 2 else left
        if j >= 5:
            cell.number_format = FMT_MM1
            cell.fill = fill_off

# Cash flow stream for IRR/NPV
ws["B13"] = "Project cash flows ($mm)"
ws["B13"].font = font_h3
ws["C13"] = "Year 0"
ws["D13"] = "Year 1"
ws["E13"] = "Year 2"
ws["F13"] = "Year 3"
ws["G13"] = "Year 4"
ws["H13"] = "Year 5"
for col in range(3, 9):
    ws.cell(13, col).font = font_navy_b
    ws.cell(13, col).alignment = center
ws["B14"] = "CF"
ws["C14"] = -250
ws["D14"] = 40
ws["E14"] = 55
ws["F14"] = 70
ws["G14"] = 80
ws["H14"] = 120
for col in range(3, 9):
    ws.cell(14, col).font = font_input
    ws.cell(14, col).fill = fill_input
    ws.cell(14, col).number_format = FMT_MM
    ws.cell(14, col).alignment = center

ws["B15"] = "Discount rate"
ws["C15"] = 0.10
ws["C15"].font = font_input
ws["C15"].fill = fill_input
ws["C15"].number_format = FMT_PCT

# Drill table
section_bar(ws, 17, 2, 12, "DRILLS  —  yellow cells are yours. Don't hardcode the answer; the point is the formula.")
cols = ["#", "SKILL", "PROMPT", "", "", "YOUR FORMULA", "RESULT", "CHECK", "HINT"]
# We'll lay out: B=# C=skill D:G=prompt merged, H=your input, I=result (can be same as H), J=check vs answer, K:L hint

ws["B18"] = "#"
ws["C18"] = "SKILL"
ws["D18"] = "PROMPT"
ws.merge_cells("D18:G18")
ws["H18"] = "YOUR ANSWER"
ws["I18"] = "CHECK"
ws["J18"] = "STATUS"
ws["K18"] = "HINT"
ws.merge_cells("K18:L18")
for col in range(2, 13):
    ws.cell(18, col).font = font_white_b
    ws.cell(18, col).fill = fill_navy2
    ws.cell(18, col).alignment = center

drills = [
    # row, num, skill, prompt, answer_formula, hint, number_format
    (19, "01", "INDEX/MATCH",
     "EBITDA of ticker HLX",
     '=INDEX($E$6:$E$11,MATCH("HLX",$C$6:$C$11,0))',
     "INDEX(return_col, MATCH(lookup, lookup_col, 0))", FMT_MM1),
    (20, "02", "INDEX/MATCH",
     "Sector of Pinnacle Tools",
     '=INDEX($D$6:$D$11,MATCH("Pinnacle Tools",$B$6:$B$11,0))',
     "Exact match — last arg of MATCH is 0", "@"),
    (21, "03", "XLOOKUP",
     "EV of Cedar Forge (XLOOKUP or INDEX/MATCH)",
     '=INDEX($E$6:$E$11,MATCH("Cedar Forge",$B$6:$B$11,0))',
     "XLOOKUP(lookup, lookup_col, return_col) if your Excel has it", FMT_MM1),
    (22, "04", "SUMIFS",
     "Total EV of Industrials only",
     '=SUMIFS($E$6:$E$11,$D$6:$D$11,"Industrials")',
     "SUMIFS(sum_range, criteria_range, criteria)", FMT_MM1),
    (23, "05", "COUNTIFS",
     "How many Industrials comps?",
     '=COUNTIFS($D$6:$D$11,"Industrials")',
     "COUNTIFS(range, criteria)", FMT_NUM),
    (24, "06", "AVERAGEIFS",
     "Average EBITDA of Industrials",
     '=AVERAGEIFS($F$6:$F$11,$D$6:$D$11,"Industrials")',
     "AVERAGEIFS(avg_range, criteria_range, criteria)", FMT_MM1),
    (25, "07", "Multiple",
     "APX EV/EBITDA",
     '=INDEX($E$6:$E$11,MATCH("APX",$C$6:$C$11,0))/INDEX($F$6:$F$11,MATCH("APX",$C$6:$C$11,0))',
     "EV ÷ EBITDA. Format as 0.0x", FMT_X),
    (26, "08", "CAGR",
     "APX FY24–FY25 revenue CAGR (1 year)",
     '=INDEX($H$6:$H$11,MATCH("APX",$C$6:$C$11,0))/INDEX($I$6:$I$11,MATCH("APX",$C$6:$C$11,0))-1',
     "End/Start − 1 for one year; for n years (End/Start)^(1/n)−1", FMT_PCT),
    (27, "09", "IF + flag",
     "1 if Helix EV > 1500, else 0",
     '=IF(INDEX($E$6:$E$11,MATCH("HLX",$C$6:$C$11,0))>1500,1,0)',
     "IF(test, 1, 0) — flags drive SUMIFS later", FMT_NUM),
    (28, "10", "IFERROR",
     "EV / 0 wrapped so it does not blow up  (use 0 as denom test)",
     '=IFERROR(980/0,"-")',
     "IFERROR(value, value_if_error). Never ship #DIV/0!", "@"),
    (29, "11", "MAX / MIN",
     "Highest EV in the set",
     "=MAX($E$6:$E$11)",
     "MAX of the EV column", FMT_MM1),
    (30, "12", "LARGE",
     "2nd-largest EBITDA",
     "=LARGE($F$6:$F$11,2)",
     "LARGE(range, k)", FMT_MM1),
    (31, "13", "MEDIAN",
     "Median EV/EBITDA of the whole set",
     "=MEDIAN($E$6:$E$11/$F$6:$F$11)",
     "MEDIAN of EV/EBITDA. Array-safe in modern Excel; otherwise helper col.", FMT_X),
    (32, "14", "NPV",
     "NPV of the project CFs at the discount rate in C15",
     "=C14+NPV(C15,D14:H14)",
     "Excel NPV assumes first value is t=1. Add CF0 separately.", FMT_MM1),
    (33, "15", "IRR",
     "IRR of the project CFs (C14:H14)",
     "=IRR(C14:H14)",
     "IRR(range). Needs a sign change in the series.", FMT_PCT),
    (34, "16", "XIRR-ready",
     "Same IRR via RATE approximation check: does IRR > 10%? 1/0",
     '=IF(IRR(C14:H14)>$C$15,1,0)',
     "Flag whether project clears the hurdle", FMT_NUM),
    (35, "17", "SUMPRODUCT",
     "EV-weighted average EBITDA margin  (EBITDA/Rev * EV) / Σ EV",
     "=SUMPRODUCT($F$6:$F$11/$H$6:$H$11,$E$6:$E$11)/SUM($E$6:$E$11)",
     "SUMPRODUCT is the IB workhorse when SUMIFS is not enough", FMT_PCT),
    (36, "18", "TEXT / label",
     "Build 'APX trades at X.Xx EBITDA'",
     '="APX trades at "&TEXT(INDEX($E$6:$E$11,MATCH("APX",$C$6:$C$11,0))/INDEX($F$6:$F$11,MATCH("APX",$C$6:$C$11,0)),"0.0")&"x EBITDA"',
     "TEXT(number, \"0.0\") to control the multiple", "@"),
    (37, "19", "ABS + check",
     "Abs difference between NLM EV and APX EV",
     '=ABS(INDEX($E$6:$E$11,MATCH("NLM",$C$6:$C$11,0))-INDEX($E$6:$E$11,MATCH("APX",$C$6:$C$11,0)))',
     "ABS for check cells so −0.0 and +0.0 both pass", FMT_MM1),
    (38, "20", "CHALLENGE",
     "P/E of Riverbend: EV is not equity. Use NI. Assume net debt = 180. P/E = (EV−net debt)/NI",
     '=(INDEX($E$6:$E$11,MATCH("RVB",$C$6:$C$11,0))-180)/INDEX($G$6:$G$11,MATCH("RVB",$C$6:$C$11,0))',
     "Equity value = EV − net debt. P/E = equity / NI", FMT_X),
]

for row, num, skill, prompt, ans, hint, fmt in drills:
    ws.cell(row, 2, num).font = font_navy_b
    ws.cell(row, 2).alignment = center
    ws.cell(row, 3, skill).font = Font(name="Arial", size=9, bold=True, color=NAVY)
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=7)
    ws.cell(row, 4, prompt).font = font_body
    ws.cell(row, 4).alignment = left
    # user answer
    ws.cell(row, 8, None)
    ws.cell(row, 8).fill = fill_input
    ws.cell(row, 8).font = font_input
    ws.cell(row, 8).number_format = fmt
    ws.cell(row, 8).alignment = center
    ws.cell(row, 8).border = thin
    # check formula compares user H to answer in O
    # STATUS
    ws.cell(row, 9, f'=IF(H{row}="","—",IF(OR(ABS(N(H{row})-N(O{row}))<0.05,H{row}=O{row}),"OK","TRY AGAIN"))')
    ws.cell(row, 9).alignment = center
    ws.cell(row, 9).font = font_formula
    ws.merge_cells(start_row=row, start_column=11, end_row=row, end_column=12)
    ws.cell(row, 11, hint).font = font_hint
    ws.row_dimensions[row].height = 22
    # hidden answers in O
    ws.cell(row, 15, ans)
    ws.cell(row, 15).font = Font(name="Arial", size=8, color="808080")
    ws.cell(row, 15).number_format = fmt
    ws.cell(row, 16, "ANSWER KEY (formula)")
    ws.cell(row, 16).font = font_hint

ws["O18"] = "ANSWER"
ws["P18"] = "FORMULA"
ws["O18"].font = font_white_b
ws["O18"].fill = fill_navy2
ws["P18"].font = font_white_b
ws["P18"].fill = fill_navy2

# conditional formatting on status
ws.conditional_formatting.add(
    "I19:I38",
    FormulaRule(formula=['I19="OK"'], fill=fill_green, font=Font(name="Arial", size=9, bold=True, color="006600")),
)
ws.conditional_formatting.add(
    "I19:I38",
    FormulaRule(formula=['I19="TRY AGAIN"'], fill=fill_red, font=Font(name="Arial", size=9, bold=True, color="9C0006")),
)

ws.merge_cells("B40:L42")
ws["B40"] = (
    "HOW BANKERS ACTUALLY LOOK THINGS UP: INDEX/MATCH (or XLOOKUP) beats VLOOKUP because you can return a column to the LEFT "
    "of the lookup column and you don't break the formula when someone inserts a column. Always MATCH(...,0) for exact. "
    "Anchor both ranges with F4 before you copy the formula down. For comps, put tickers in a helper column and never type a name twice."
)
ws["B40"].font = font_body
ws["B40"].alignment = left_top

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 04 3-statement
# Layout:
# Col B labels
# C assumptions / notes
# D 2023A  E 2024A  F 2025A  G 2026E  H 2027E  I 2028E
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("04_3Statement")
banner(ws, "04  ·  APEX INDUSTRIAL  —  MINI 3-STATEMENT MODEL ($mm)", "Yellow + blue = toggle me. Historicals are blue hardcodes. Forecasts are black formulas. Green = cross-sheet.", 10)

width(ws, {"A": 3, "B": 34, "C": 16, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 28})
ws.freeze_panes = "D6"
ws.sheet_view.showGridLines = False

# Years
ws["B5"] = "Line item"
ws["C5"] = "Driver / note"
years = [("D", "2023A"), ("E", "2024A"), ("F", "2025A"), ("G", "2026E"), ("H", "2027E"), ("I", "2028E")]
for col, y in years:
    ws[f"{col}5"] = y
    ws[f"{col}5"].font = font_white_b
    ws[f"{col}5"].fill = fill_navy
    ws[f"{col}5"].alignment = center
ws["B5"].font = font_white_b
ws["B5"].fill = fill_navy
ws["C5"].font = font_white_b
ws["C5"].fill = fill_navy
ws["J5"] = "Commentary"
ws["J5"].font = font_white_b
ws["J5"].fill = fill_navy
ws.row_dimensions[5].height = 20

# Status flag
ws["B4"] = "MODEL STATUS"
ws["B4"].font = font_navy_b
ws["C4"] = '=IF(ABS(Checks!C6)<0.05,"BALANCES","BREAKS")'
ws["C4"].font = Font(name="Arial", size=12, bold=True, color="006600")
ws["C4"].alignment = center
ws["D4"] = "← live link to Checks tab (green font once you look at the formula)"
ws["D4"].font = font_hint
ws.merge_cells("D4:I4")
# make C4 a cross-sheet link style
ws["C4"].font = Font(name="Arial", size=12, bold=True, color="008000")

# ── ASSUMPTIONS ──
section_bar(ws, 6, 2, 10, "ASSUMPTIONS  (the only place you type forecasts)")

ws["B7"] = "Revenue growth"
ws["C7"] = "YoY"
ws["D7"] = "—"
ws["E7"] = "=D20/C20-1"  # wait years are D=2023 E=2024, revenue will be row 20
# I'll set rows carefully.

# Let me lock row map:
# 6 section assumptions
# 7 rev growth
# 8 COGS %
# 9 opex %
# 10 D&A % sales
# 11 tax rate
# 12 NWC % sales
# 13 capex % sales
# 14 interest rate on beg debt
# 15 debt paydown $
# 16 dividend payout
# 17 blank
# 18 section IS
# 19 header?
# 20 Revenue
# ...

# Actually historical growth can be formulas from revenue.
# Forecast growth is input.

# Assumptions rows 7-16
# Historical columns D-F: growth computed or "—", forecast G-I: inputs

ws["B7"] = "Revenue growth"
ws["C7"] = "YoY"
ws["D7"] = "n.a."
# E7 and F7 will be formulas vs revenue once revenue exists — but that creates a circular look from assumptions to IS.
# Cleaner: historical growth is just displayed from IS, forecast growth is input.
# Put forecast inputs only in G-I for growth.

ws["E7"] = 0.0952  # 460/420-1 placeholder, will replace with formula referencing revenue
# Better approach: all assumption G-I are inputs. D-F historical drivers computed below the IS or hardcoded.

# I'll hardcode forecast drivers as inputs (blue/yellow) and compute historical % as formulas from IS (black).

drivers = [
    # row, label, note, hist_d, hist_e, hist_f, f1, f2, f3, fmt, comment
    (7, "Revenue growth", "YoY", None, None, None, 0.08, 0.07, 0.06, FMT_PCT,
     "Source: management CIM base case, 8/2026. Street is 7.5% long-run."),
    (8, "COGS % sales", "incl. freight", 0.62, 0.615, 0.61, 0.605, 0.60, 0.598, FMT_PCT,
     "Gross margin expansion from mix shift to aftermarket."),
    (9, "OpEx % sales", "SG&A + R&D", 0.18, 0.178, 0.175, 0.172, 0.17, 0.168, FMT_PCT,
     "Operating leverage. Do not cut below 16% without a cost-out program."),
    (10, "D&A % sales", "book", None, None, None, 0.042, 0.042, 0.041, FMT_PCT,
     "Holds near historical ~4% of sales."),
    (11, "Tax rate", "book", 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, FMT_PCT,
     "Source: FY25 10-K rate recon, ~statutory 25%."),
    (12, "NWC % sales", "ex-cash", 0.114, 0.117, 0.118, 0.12, 0.12, 0.12, FMT_PCT,
     "AR+Inv−AP. Held at 12% in the forecast."),
    (13, "Capex % sales", "cash", None, None, None, 0.052, 0.050, 0.048, FMT_PCT,
     "Slightly above D&A to fund the capacity add."),
    (14, "Interest rate", "on beg. debt", None, None, None, 0.06, 0.06, 0.06, FMT_PCT,
     "Blended coupon. Circularity: interest depends on debt, debt on cash. See Checks."),
    (15, "Debt paydown", "$mm", 10, 10, 10, 10, 10, 10, FMT_MM,
     "Amortization schedule per credit agreement."),
    (16, "Dividend payout", "% of NI", 0.30, 0.30, 0.30, 0.30, 0.30, 0.30, FMT_PCT,
     "Board policy. Residual cash builds on BS."),
]

for row, label, note, d, e, f, g, h, i, fmt, cmt in drivers:
    ws.cell(row, 2, label).font = font_body
    ws.cell(row, 3, note).font = font_hint
    vals = [d, e, f, g, h, i]
    for col, val in zip(range(4, 10), vals):
        cell = ws.cell(row, col, val if val is not None else "—")
        cell.alignment = center
        if val is None:
            cell.font = font_hint
        elif col >= 7 or row in (8, 9, 11, 12, 15, 16):
            # forecast always input; some hist also input
            if isinstance(val, (int, float)):
                cell.font = font_input
                cell.fill = fill_input
                cell.number_format = fmt
            else:
                cell.font = font_hint
        else:
            if isinstance(val, (int, float)):
                cell.font = font_input
                cell.fill = fill_input
                cell.number_format = fmt
    ws.cell(row, 10, cmt).font = font_hint
    ws.cell(row, 4).comment = Comment(cmt, "Dojo") if cmt else None

# Historical D&A % and capex % and interest will be shown as formulas after IS/BS exist.
# For D10:F10 leave as em-dash for now; optional later.

# ── INCOME STATEMENT ──
section_bar(ws, 18, 2, 10, "INCOME STATEMENT")

# Row 19 spacer labels already in row 5
# 20 Revenue
# 21 COGS
# 22 GP
# 23 OpEx
# 24 EBITDA
# 25 D&A
# 26 EBIT
# 27 Interest
# 28 EBT
# 29 Tax
# 30 NI

def is_label(row, text, indent=False, bold=False):
    ws.cell(row, 2, ("  " + text) if indent else text)
    ws.cell(row, 2).font = Font(name="Arial", size=10, bold=bold, color="000000")

is_label(20, "Revenue", bold=True)
ws["D20"] = 420
ws["E20"] = 460
ws["F20"] = 510
ws["G20"] = "=F20*(1+G7)"
ws["H20"] = "=G20*(1+H7)"
ws["I20"] = "=H20*(1+I7)"

is_label(21, "COGS", indent=True)
ws["D21"] = "=-D20*D8"
ws["E21"] = "=-E20*E8"
ws["F21"] = "=-F20*F8"
ws["G21"] = "=-G20*G8"
ws["H21"] = "=-H20*H8"
ws["I21"] = "=-I20*I8"

is_label(22, "Gross Profit", bold=True)
ws["D22"] = "=D20+D21"
ws["E22"] = "=E20+E21"
ws["F22"] = "=F20+F21"
ws["G22"] = "=G20+G21"
ws["H22"] = "=H20+H21"
ws["I22"] = "=I20+I21"

is_label(23, "Operating expenses", indent=True)
ws["D23"] = "=-D20*D9"
ws["E23"] = "=-E20*E9"
ws["F23"] = "=-F20*F9"
ws["G23"] = "=-G20*G9"
ws["H23"] = "=-H20*H9"
ws["I23"] = "=-I20*I9"

is_label(24, "EBITDA", bold=True)
ws["D24"] = "=D22+D23"
ws["E24"] = "=E22+E23"
ws["F24"] = "=F22+F23"
ws["G24"] = "=G22+G23"
ws["H24"] = "=H22+H23"
ws["I24"] = "=I22+I23"

is_label(25, "D&A", indent=True)
ws["D25"] = -18
ws["E25"] = -19
ws["F25"] = -21
ws["G25"] = "=-G20*G10"
ws["H25"] = "=-H20*H10"
ws["I25"] = "=-I20*I10"

is_label(26, "EBIT", bold=True)
ws["D26"] = "=D24+D25"
ws["E26"] = "=E24+E25"
ws["F26"] = "=F24+F25"
ws["G26"] = "=G24+G25"
ws["H26"] = "=H24+H25"
ws["I26"] = "=I24+I25"

is_label(27, "Interest expense", indent=True)
ws["D27"] = -8
ws["E27"] = -7.5
ws["F27"] = -7
# Interest on beginning-of-year debt (row 48 will be Debt)
ws["G27"] = "=-F48*G14"
ws["H27"] = "=-G48*H14"
ws["I27"] = "=-H48*I14"

is_label(28, "EBT", bold=True)
ws["D28"] = "=D26+D27"
ws["E28"] = "=E26+E27"
ws["F28"] = "=F26+F27"
ws["G28"] = "=G26+G27"
ws["H28"] = "=H26+H27"
ws["I28"] = "=I26+I27"

is_label(29, "Tax", indent=True)
ws["D29"] = "=-D28*D11"
ws["E29"] = "=-E28*E11"
ws["F29"] = "=-F28*F11"
ws["G29"] = "=-G28*G11"
ws["H29"] = "=-H28*H11"
ws["I29"] = "=-I28*I11"

is_label(30, "Net Income", bold=True)
ws["D30"] = "=D28+D29"
ws["E30"] = "=E28+E29"
ws["F30"] = "=F28+F29"
ws["G30"] = "=G28+G29"
ws["H30"] = "=H28+H29"
ws["I30"] = "=I28+I29"

# margins
is_label(32, "Memo: EBITDA margin")
ws["D32"] = "=IF(D20=0,\"-\",D24/D20)"
ws["E32"] = "=E24/E20"
ws["F32"] = "=F24/F20"
ws["G32"] = "=G24/G20"
ws["H32"] = "=H24/H20"
ws["I32"] = "=I24/I20"

is_label(33, "Memo: NI margin")
ws["D33"] = "=D30/D20"
ws["E33"] = "=E30/E20"
ws["F33"] = "=F30/F20"
ws["G33"] = "=G30/G20"
ws["H33"] = "=H30/H20"
ws["I33"] = "=I30/I20"

is_label(34, "Memo: YoY revenue growth")
ws["D34"] = "n.a."
ws["E34"] = "=E20/D20-1"
ws["F34"] = "=F20/E20-1"
ws["G34"] = "=G20/F20-1"
ws["H34"] = "=H20/G20-1"
ws["I34"] = "=I20/H20-1"

# ── CASH FLOW ──
section_bar(ws, 36, 2, 10, "CASH FLOW STATEMENT")

is_label(37, "Net Income")
ws["D37"] = "=D30"
ws["E37"] = "=E30"
ws["F37"] = "=F30"
ws["G37"] = "=G30"
ws["H37"] = "=H30"
ws["I37"] = "=I30"

is_label(38, "(+) D&A", indent=True)
ws["D38"] = "=-D25"
ws["E38"] = "=-E25"
ws["F38"] = "=-F25"
ws["G38"] = "=-G25"
ws["H38"] = "=-H25"
ws["I38"] = "=-I25"

is_label(39, "(−) Increase in NWC", indent=True)
# NWC row 47
ws["D39"] = "n.a."
ws["E39"] = "=-(E47-D47)"
ws["F39"] = "=-(F47-E47)"
ws["G39"] = "=-(G47-F47)"
ws["H39"] = "=-(H47-G47)"
ws["I39"] = "=-(I47-H47)"

is_label(40, "(−) Capex", indent=True)
ws["D40"] = -22
ws["E40"] = -24
ws["F40"] = -26
ws["G40"] = "=-G20*G13"
ws["H40"] = "=-H20*H13"
ws["I40"] = "=-I20*I13"

is_label(41, "Free cash flow", bold=True)
ws["D41"] = "n.a."
ws["E41"] = "=E37+E38+E39+E40"
ws["F41"] = "=F37+F38+F39+F40"
ws["G41"] = "=G37+G38+G39+G40"
ws["H41"] = "=H37+H38+H39+H40"
ws["I41"] = "=I37+I38+I39+I40"

is_label(42, "(−) Debt paydown", indent=True)
ws["D42"] = "=-D15"
ws["E42"] = "=-E15"
ws["F42"] = "=-F15"
ws["G42"] = "=-G15"
ws["H42"] = "=-H15"
ws["I42"] = "=-I15"

is_label(43, "(−) Dividends", indent=True)
ws["D43"] = "=-D30*D16"
ws["E43"] = "=-E30*E16"
ws["F43"] = "=-F30*F16"
ws["G43"] = "=-G30*G16"
ws["H43"] = "=-H30*H16"
ws["I43"] = "=-I30*I16"

is_label(44, "Net change in cash", bold=True)
ws["D44"] = "n.a."
ws["E44"] = "=E41+E42+E43"
ws["F44"] = "=F41+F42+F43"
ws["G44"] = "=G41+G42+G43"
ws["H44"] = "=H41+H42+H43"
ws["I44"] = "=I41+I42+I43"

# ── BALANCE SHEET ──
section_bar(ws, 46, 2, 10, "BALANCE SHEET")

is_label(47, "Net working capital")
ws["D47"] = "=D20*D12"
ws["E47"] = "=E20*E12"
ws["F47"] = "=F20*F12"
ws["G47"] = "=G20*G12"
ws["H47"] = "=H20*H12"
ws["I47"] = "=I20*I12"

is_label(48, "Debt")
ws["D48"] = 140
ws["E48"] = "=D48-E15"
ws["F48"] = "=E48-F15"
ws["G48"] = "=F48-G15"
ws["H48"] = "=G48-H15"
ws["I48"] = "=H48-I15"

is_label(49, "Net PP&E")
# Rollforward: beg + capex + DA (DA is negative on IS)
ws["D49"] = 210
ws["E49"] = "=D49-E40+E25"  # - capex(neg) wait capex is stored negative on CFS so -E40 is positive spend
# CFS capex D40 = -22, so adding PP&E should add 22: PP&E_end = PP&E_beg + (-capex_cfs) + DA_is
# DA_is is -18, so PP&E_end = 210 + 22 - 18 = 214
ws["E49"] = "=D49+(-E40)+E25"
ws["F49"] = "=E49+(-F40)+F25"
ws["G49"] = "=F49+(-G40)+G25"
ws["H49"] = "=G49+(-H40)+H25"
ws["I49"] = "=H49+(-I40)+I25"

is_label(50, "Cash", bold=True)
ws["D50"] = 35
ws["E50"] = "=D50+E44"
ws["F50"] = "=E50+F44"
ws["G50"] = "=F50+G44"
ws["H50"] = "=G50+H44"
ws["I50"] = "=H50+I44"

is_label(51, "Total assets", bold=True)
ws["D51"] = "=D50+D47+D49"
ws["E51"] = "=E50+E47+E49"
ws["F51"] = "=F50+F47+F49"
ws["G51"] = "=G50+G47+G49"
ws["H51"] = "=H50+H47+H49"
ws["I51"] = "=I50+I47+I49"

is_label(52, "Equity (RE rollforward)")
# Equity plug for 2023 so BS balances: Assets - Debt
ws["D52"] = "=D51-D48"
ws["E52"] = "=D52+E30+E43"  # +NI + dividends(neg)
ws["F52"] = "=E52+F30+F43"
ws["G52"] = "=F52+G30+G43"
ws["H52"] = "=G52+H30+H43"
ws["I52"] = "=H52+I30+I43"

is_label(53, "Total liabilities & equity", bold=True)
ws["D53"] = "=D48+D52"
ws["E53"] = "=E48+E52"
ws["F53"] = "=F48+F52"
ws["G53"] = "=G48+G52"
ws["H53"] = "=H48+H52"
ws["I53"] = "=I48+I52"

is_label(55, "Balance check (A − L−E)", bold=True)
ws["D55"] = "=D51-D53"
ws["E55"] = "=E51-E53"
ws["F55"] = "=F51-F53"
ws["G55"] = "=G51-G53"
ws["H55"] = "=H51-H53"
ws["I55"] = "=I51-I53"

is_label(56, "Cash roll check vs CFS")
ws["D56"] = "—"
ws["E56"] = "=E50-(D50+E44)"
ws["F56"] = "=F50-(E50+F44)"
ws["G56"] = "=G50-(F50+G44)"
ws["H56"] = "=H50-(G50+H44)"
ws["I56"] = "=I50-(H50+I44)"

# formatting for IS/CFS/BS number cells
input_cells = {
    # historical hardcodes
    "D20", "E20", "F20",
    "D25", "E25", "F25",
    "D27", "E27", "F27",
    "D40", "E40", "F40",
    "D48",
    "D49",
    "D50",
}
pct_rows = {7, 8, 9, 10, 11, 12, 13, 14, 16, 32, 33, 34}
mm_rows = set(range(20, 31)) | set(range(37, 45)) | set(range(47, 57)) | {15}

for row in range(7, 57):
    for col in range(4, 10):
        cell = ws.cell(row, col)
        addr = f"{get_column_letter(col)}{row}"
        if row in pct_rows and isinstance(cell.value, (int, float, str)) and isinstance(cell.value, str) and cell.value.startswith("="):
            cell.number_format = FMT_PCT
        elif row in pct_rows:
            if cell.value not in ("—", "n.a.", None, "n.a."):
                cell.number_format = FMT_PCT
        elif row in mm_rows:
            if cell.value not in ("—", "n.a.", None):
                cell.number_format = FMT_MM1

        # fonts
        if addr in input_cells:
            cell.font = font_input
            cell.fill = fill_input
        elif isinstance(cell.value, str) and cell.value.startswith("="):
            if row == 4:
                cell.font = Font(name="Arial", size=12, bold=True, color="008000")
            else:
                cell.font = font_formula
        cell.alignment = center

# borders on totals
for row in (22, 24, 26, 30, 41, 44, 51, 53):
    for col in range(2, 10):
        ws.cell(row, col).border = bottom_single
for col in range(2, 10):
    ws.cell(30, col).border = top_bottom
    ws.cell(55, col).border = top_bottom

# check CF on status row
ws.conditional_formatting.add(
    "C4",
    FormulaRule(formula=['C4="BALANCES"'], fill=fill_green, font=Font(name="Arial", size=12, bold=True, color="006000")),
)
ws.conditional_formatting.add(
    "C4",
    FormulaRule(formula=['C4="BREAKS"'], fill=fill_red, font=Font(name="Arial", size=12, bold=True, color="9C0006")),
)
ws.conditional_formatting.add(
    "D55:I56",
    FormulaRule(formula=["ABS(D55)<0.05"], fill=fill_green),
)
ws.conditional_formatting.add(
    "D55:I56",
    FormulaRule(formula=["ABS(D55)>=0.05"], fill=fill_red),
)

ws["J20"] = "Change growth in G7:I7 and watch the whole forecast rewrite. That is the point."
ws["J20"].font = font_hint
ws["J30"] = "NI is the bridge from IS → CFS → Equity."
ws["J30"].font = font_hint
ws["J44"] = "Ending cash on BS must equal beginning cash + this line."
ws["J44"].font = font_hint
ws["J50"] = "Cash is NOT a plug here — it is computed from the CFS. If BS breaks, a flow is missing."
ws["J50"].font = font_hint
ws["J55"] = "Must be 0.000 in every year. If it is not, do not pass the file on."
ws["J55"].font = Font(name="Arial", size=9, italic=True, color="C00000")

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 05 Sensitivity
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("05_Sensitivity")
banner(ws, "05  ·  SENSITIVITY & CASES", "Bankers do not present one number. They present a grid. Yellow cells drive everything.", 12)
width(ws, {"A": 3, "B": 22, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 28})

section_bar(ws, 4, 2, 12, "SIMPLE DCF ON APEX  —  FY28 EBITDA exit")

ws["B5"] = "FY28 EBITDA"
ws["C5"] = "=IFERROR('04_3Statement'!I24,0)"
ws["C5"].font = font_link
ws["C5"].number_format = FMT_MM1
ws["D5"] = "← green: live link to the 3-statement"
ws["D5"].font = font_hint
ws.merge_cells("D5:G5")

ws["B6"] = "Exit EV/EBITDA"
ws["C6"] = 8.5
ws["C6"].font = font_input
ws["C6"].fill = fill_input
ws["C6"].number_format = FMT_X

ws["B7"] = "FY28 net debt"
ws["C7"] = "='04_3Statement'!I48-'04_3Statement'!I50"
ws["C7"].font = font_link
ws["C7"].number_format = FMT_MM1

ws["B8"] = "Exit enterprise value"
ws["C8"] = "=C5*C6"
ws["C8"].font = font_formula
ws["C8"].number_format = FMT_MM1

ws["B9"] = "Exit equity value"
ws["C9"] = "=C8-C7"
ws["C9"].font = font_formula
ws["C9"].number_format = FMT_MM1

ws["B10"] = "Mid-year discount years"
ws["C10"] = 2.5
ws["C10"].font = font_input
ws["C10"].fill = fill_input
ws["C10"].number_format = "0.0"

ws["B11"] = "WACC (base)"
ws["C11"] = 0.095
ws["C11"].font = font_input
ws["C11"].fill = fill_input
ws["C11"].number_format = FMT_PCT

ws["B12"] = "PV of exit equity"
ws["C12"] = "=C9/(1+C11)^C10"
ws["C12"].font = font_formula
ws["C12"].number_format = FMT_MM1

ws["B13"] = "Shares (mm)"
ws["C13"] = 40
ws["C13"].font = font_input
ws["C13"].fill = fill_input
ws["C13"].number_format = FMT_NUM1

ws["B14"] = "Implied value / share"
ws["C14"] = "=IF(C13=0,\"-\",C12/C13)"
ws["C14"].font = Font(name="Arial", size=12, bold=True)
ws["C14"].number_format = '"$"#,##0.00;("$"#,##0.00);"-"'
ws["C14"].fill = fill_pale

for r in range(5, 15):
    ws.cell(r, 2).font = font_body
    ws.cell(r, 3).alignment = center

# Two-way table: WACC rows × exit multiple cols
section_bar(ws, 16, 2, 12, "TWO-WAY SENSITIVITY  —  implied value / share  (rebuild this with Data Table once you know Alt A W T)")

ws["B17"] = "WACC \\ Exit x"
ws["B17"].font = font_navy_b
ws["B17"].fill = fill_section

multiples = [7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0]
waccs = [0.080, 0.085, 0.090, 0.095, 0.100, 0.105, 0.110]
for i, m in enumerate(multiples):
    cell = ws.cell(17, 3 + i, m)
    cell.font = font_input
    cell.fill = fill_input
    cell.number_format = FMT_X
    cell.alignment = center
for i, w in enumerate(waccs):
    cell = ws.cell(18 + i, 2, w)
    cell.font = font_input
    cell.fill = fill_input
    cell.number_format = FMT_PCT
    cell.alignment = center

# Corner cell is the formula Data Table would hook to
ws["B17"] = "=C14"
ws["B17"].number_format = '"$"#,##0.00'
ws["B17"].font = font_formula
ws["B17"].fill = fill_pale

# Explicit formulas so it works without Data Table
# value = (EBITDA * multiple - net_debt) / (1+wacc)^years / shares
for i, w in enumerate(waccs):
    for j, m in enumerate(multiples):
        # C18 is first data cell (row 18, col 3)
        cell = ws.cell(18 + i, 3 + j)
        # reference the header WACC and header multiple
        col_letter = get_column_letter(3 + j)
        cell.value = f"=(($C$5*{col_letter}$17)-$C$7)/(1+$B{18+i})^$C$10/$C$13"
        cell.number_format = '"$"#,##0.00;("$"#,##0.00);"-"'
        cell.font = font_formula
        cell.alignment = center
        cell.border = thin

# Color scale on the grid
ws.conditional_formatting.add(
    "C18:I24",
    ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B",
    ),
)

ws.merge_cells("B26:I28")
ws["B26"] = (
    "HOW TO TURN THIS INTO A REAL DATA TABLE: put =C14 in the corner (B17). Clear C18:I24. "
    "Select B17:I24. Alt → A → W → T. Row input = C6 (the multiple). Column input = C11 (the WACC). "
    "Excel writes {=TABLE(C6,C11)} as an array. Set calc to Automatic except data tables or F9 will crawl. "
    "The formulas above already compute the same grid so you can see the math either way."
)
ws["B26"].font = font_body
ws["B26"].alignment = left_top

# Case toggle
section_bar(ws, 30, 2, 12, "OPERATING CASE TOGGLE  —  feed this back into 04 if you want to get fancy")

ws["B31"] = "Active case"
ws["C31"] = "Base"
ws["C31"].font = font_input
ws["C31"].fill = fill_input
ws["C31"].alignment = center
dv = DataValidation(type="list", formula1='"Downside,Base,Upside"', allow_blank=False)
dv.error = "Pick Downside, Base, or Upside"
dv.errorTitle = "Case"
ws.add_data_validation(dv)
dv.add(ws["C31"])

ws["B32"] = "Case #"
ws["C32"] = '=MATCH(C31,{"Downside","Base","Upside"},0)'
ws["C32"].font = font_formula
ws["C32"].alignment = center

headers = ["Driver", "Downside", "Base", "Upside", "Active (CHOOSE)"]
for i, h in enumerate(headers, start=2):
    c = ws.cell(34, i, h)
    c.font = font_white_b
    c.fill = fill_navy2
    c.alignment = center

cases = [
    (35, "FY26 rev growth", 0.03, 0.08, 0.12, FMT_PCT),
    (36, "FY26 COGS %", 0.63, 0.605, 0.58, FMT_PCT),
    (37, "Exit multiple", 7.0, 8.5, 10.0, FMT_X),
    (38, "WACC", 0.11, 0.095, 0.085, FMT_PCT),
]
for row, label, d, b, u, fmt in cases:
    ws.cell(row, 2, label).font = font_body
    ws.cell(row, 3, d).font = font_input
    ws.cell(row, 3).fill = fill_input
    ws.cell(row, 4, b).font = font_input
    ws.cell(row, 4).fill = fill_input
    ws.cell(row, 5, u).font = font_input
    ws.cell(row, 5).fill = fill_input
    ws.cell(row, 6, f"=CHOOSE($C$32,C{row},D{row},E{row})")
    ws.cell(row, 6).font = font_formula
    ws.cell(row, 6).fill = fill_pale
    for col in range(3, 7):
        ws.cell(row, col).number_format = fmt
        ws.cell(row, col).alignment = center

ws["B40"] = "CHOOSE + a dropdown is how many 'case' models work before anyone writes VBA."
ws["B40"].font = font_hint
ws.merge_cells("B40:F40")

ws["B42"] = "Active value / share (uses case WACC + case multiple, same EBITDA)"
ws["B42"].font = font_body
ws.merge_cells("B42:E42")
ws["F42"] = "=(($C$5*F37)-$C$7)/(1+F38)^$C$10/$C$13"
ws["F42"].font = Font(name="Arial", size=12, bold=True)
ws["F42"].number_format = '"$"#,##0.00;("$"#,##0.00);"-"'
ws["F42"].fill = fill_pale
ws["F42"].alignment = center

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 06 Comps
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("06_Comps")
banner(ws, "06  ·  TRADING COMPS DRILL", "Build EV and multiples from the yellow inputs. Stats at the bottom feed a football field.", 12)
width(ws, {"A": 3, "B": 22, "C": 10, "D": 14, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12, "L": 14})

headers = ["Company", "Ticker", "Price", "Shares (mm)", "Equity val", "Net debt", "EV", "EBITDA", "NI", "EV/EBITDA", "P/E"]
for i, h in enumerate(headers, start=2):
    c = ws.cell(4, i, h)
    c.font = font_white_b
    c.fill = fill_navy
    c.alignment = center
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[4].height = 28

# inputs: price, shares, net debt, ebitda, ni
# equity = price * shares
# EV = equity + net debt
comps = [
    # name, ticker, price, shares, net debt, ebitda, ni
    ("Apex Industrial", "APX", 22.50, 40.0, 85.0, 109.7, 61.2),
    ("Northline Mfg", "NLM", 34.00, 38.0, 128.0, 155.0, 88.0),
    ("Cedar Forge", "CFR", 18.25, 36.0, 103.0, 92.0, 41.5),
    ("Helix Components", "HLX", 41.80, 48.0, 98.4, 240.0, 130.4),
    ("Pinnacle Tools", "PNT", 15.10, 28.0, 117.2, 61.0, 22.0),
    ("Riverbend Co", "RVB", 27.40, 40.0, 84.0, 150.0, 70.1),
]
for i, (name, tkr, px, sh, nd, ebitda, ni) in enumerate(comps):
    r = 5 + i
    ws.cell(r, 2, name).font = font_body
    ws.cell(r, 3, tkr).font = font_body
    ws.cell(r, 3).alignment = center
    ws.cell(r, 4, px).font = font_input
    ws.cell(r, 4).fill = fill_input
    ws.cell(r, 4).number_format = '"$"#,##0.00'
    ws.cell(r, 5, sh).font = font_input
    ws.cell(r, 5).fill = fill_input
    ws.cell(r, 5).number_format = FMT_NUM1
    ws.cell(r, 6, f"=D{r}*E{r}")
    ws.cell(r, 6).font = font_formula
    ws.cell(r, 6).number_format = FMT_MM1
    ws.cell(r, 7, nd).font = font_input
    ws.cell(r, 7).fill = fill_input
    ws.cell(r, 7).number_format = FMT_MM1
    ws.cell(r, 8, f"=F{r}+G{r}")
    ws.cell(r, 8).font = font_formula
    ws.cell(r, 8).number_format = FMT_MM1
    ws.cell(r, 9, ebitda).font = font_input
    ws.cell(r, 9).fill = fill_input
    ws.cell(r, 9).number_format = FMT_MM1
    ws.cell(r, 10, ni).font = font_input
    ws.cell(r, 10).fill = fill_input
    ws.cell(r, 10).number_format = FMT_MM1
    ws.cell(r, 11, f"=IF(I{r}=0,\"-\",H{r}/I{r})")
    ws.cell(r, 11).font = font_formula
    ws.cell(r, 11).number_format = FMT_X
    ws.cell(r, 12, f"=IF(J{r}=0,\"-\",F{r}/J{r})")
    ws.cell(r, 12).font = font_formula
    ws.cell(r, 12).number_format = FMT_X
    for col in range(3, 13):
        ws.cell(r, col).alignment = center
    if name == "Apex Industrial":
        fill_row(ws, r, 2, 12, PatternFill("solid", fgColor="E2EFDA"))
        ws.cell(r, 4).fill = fill_input
        ws.cell(r, 5).fill = fill_input
        ws.cell(r, 7).fill = fill_input
        ws.cell(r, 9).fill = fill_input
        ws.cell(r, 10).fill = fill_input

# Highlight APX note
ws["B11"] = "Apex row is shaded — it is the target. Peers only go into the stats below."
ws["B11"].font = font_hint
ws.merge_cells("B11:L11")

section_bar(ws, 13, 2, 12, "PEER STATS  (excludes Apex — INDEX/SMALL or just explicit ranges C6:C10 wait rows 6-10 are peers)")

# peers are rows 6-10
ws["B14"] = "Statistic"
ws["K14"] = "EV/EBITDA"
ws["L14"] = "P/E"
for col in (2, 11, 12):
    ws.cell(14, col).font = font_white_b
    ws.cell(14, col).fill = fill_navy2
    ws.cell(14, col).alignment = center

stats = [
    (15, "High", "=MAX(K6:K10)", "=MAX(L6:L10)"),
    (16, "75th pctile", "=PERCENTILE(K6:K10,0.75)", "=PERCENTILE(L6:L10,0.75)"),
    (17, "Median", "=MEDIAN(K6:K10)", "=MEDIAN(L6:L10)"),
    (18, "Mean", "=AVERAGE(K6:K10)", "=AVERAGE(L6:L10)"),
    (19, "25th pctile", "=PERCENTILE(K6:K10,0.25)", "=PERCENTILE(L6:L10,0.25)"),
    (20, "Low", "=MIN(K6:K10)", "=MIN(L6:L10)"),
    (21, "Apex (target)", "=K5", "=L5"),
]
for row, label, a, b in stats:
    ws.cell(row, 2, label).font = font_section if row == 21 else font_body
    ws.cell(row, 11, a).font = font_link if row == 21 else font_formula
    ws.cell(row, 12, b).font = font_link if row == 21 else font_formula
    ws.cell(row, 11).number_format = FMT_X
    ws.cell(row, 12).number_format = FMT_X
    ws.cell(row, 11).alignment = center
    ws.cell(row, 12).alignment = center
    if row == 17:
        fill_row(ws, row, 2, 12, fill_pale)
    if row == 21:
        fill_row(ws, row, 2, 12, PatternFill("solid", fgColor="E2EFDA"))

ws["B23"] = "Implied Apex equity value at peer median EV/EBITDA"
ws["B23"].font = font_body
ws.merge_cells("B23:F23")
ws["G23"] = "=K17*I5-G5"
ws["G23"].font = font_formula
ws["G23"].number_format = FMT_MM1
ws["H23"] = "$mm"
ws["H23"].font = font_hint

ws["B24"] = "Implied Apex price at peer median EV/EBITDA"
ws["B24"].font = font_body
ws.merge_cells("B24:F24")
ws["G24"] = "=IF(E5=0,\"-\",G23/E5)"
ws["G24"].font = Font(name="Arial", size=12, bold=True)
ws["G24"].number_format = '"$"#,##0.00'
ws["G24"].fill = fill_pale

ws["B25"] = "Implied Apex price at peer median P/E"
ws["B25"].font = font_body
ws.merge_cells("B25:F25")
ws["G25"] = "=IF(E5=0,\"-\",(L17*J5)/E5)"
ws["G25"].font = Font(name="Arial", size=12, bold=True)
ws["G25"].number_format = '"$"#,##0.00'
ws["G25"].fill = fill_pale

ws.merge_cells("B27:L29")
ws["B27"] = (
    "FOOTBALL-FIELD LOGIC: a deck slide is just these ranges drawn as bars — DCF (from tab 05), "
    "52-week high/low, peer EV/EBITDA (low–high and 25–75), peer P/E, precedent transactions. "
    "The work is not the drawing. The work is that every bar traces back to a cell a VP can audit. "
    "Change Helix's price in D8 and watch the median and implied Apex price move. That is a live comps file."
)
ws["B27"].font = font_body
ws["B27"].alignment = left_top

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 07 Checks
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("07_Checks")
banner(ws, "07  ·  ERROR CHECK DASHBOARD", "If this page is red, the file does not leave your desk. Green font = pulled from other tabs.", 8)
width(ws, {"A": 3, "B": 42, "C": 16, "D": 16, "E": 16, "F": 16, "G": 18, "H": 28})

ws["B4"] = "Check"
ws["C4"] = "2023A"
ws["D4"] = "2024A"
ws["E4"] = "2025A"
ws["F4"] = "2026E"
ws["G4"] = "2027E"
ws["H4"] = "2028E"
for col in range(2, 9):
    ws.cell(4, col).font = font_white_b
    ws.cell(4, col).fill = fill_navy
    ws.cell(4, col).alignment = center

# C6 used by 04_3Statement status — max abs balance check
ws["B5"] = "BS balances (Assets − L&E)"
ws["C5"] = "='04_3Statement'!D55"
ws["D5"] = "='04_3Statement'!E55"
ws["E5"] = "='04_3Statement'!F55"
ws["F5"] = "='04_3Statement'!G55"
ws["G5"] = "='04_3Statement'!H55"
ws["H5"] = "='04_3Statement'!I55"

ws["B6"] = "MAX abs BS break (feeds the status flag)"
ws["C6"] = "=MAX(ABS(C5),ABS(D5),ABS(E5),ABS(F5),ABS(G5),ABS(H5))"
ws.merge_cells("C6:H6")
ws["C6"].alignment = Alignment(horizontal="left", vertical="center")

ws["B7"] = "Cash roll (BS cash − beg cash − Δ cash)"
ws["C7"] = "='04_3Statement'!D56"
ws["D7"] = "='04_3Statement'!E56"
ws["E7"] = "='04_3Statement'!F56"
ws["F7"] = "='04_3Statement'!G56"
ws["G7"] = "='04_3Statement'!H56"
ws["H7"] = "='04_3Statement'!I56"

ws["B8"] = "Debt roll (end − beg + paydown should be 0)"
ws["C8"] = "—"
ws["D8"] = "='04_3Statement'!E48-('04_3Statement'!D48-'04_3Statement'!E15)"
ws["E8"] = "='04_3Statement'!F48-('04_3Statement'!E48-'04_3Statement'!F15)"
ws["F8"] = "='04_3Statement'!G48-('04_3Statement'!F48-'04_3Statement'!G15)"
ws["G8"] = "='04_3Statement'!H48-('04_3Statement'!G48-'04_3Statement'!H15)"
ws["H8"] = "='04_3Statement'!I48-('04_3Statement'!H48-'04_3Statement'!I15)"

ws["B9"] = "NI identity (EBT + Tax − NI should be 0)"
ws["C9"] = "='04_3Statement'!D28+'04_3Statement'!D29-'04_3Statement'!D30"
ws["D9"] = "='04_3Statement'!E28+'04_3Statement'!E29-'04_3Statement'!E30"
ws["E9"] = "='04_3Statement'!F28+'04_3Statement'!F29-'04_3Statement'!F30"
ws["F9"] = "='04_3Statement'!G28+'04_3Statement'!G29-'04_3Statement'!G30"
ws["G9"] = "='04_3Statement'!H28+'04_3Statement'!H29-'04_3Statement'!H30"
ws["H9"] = "='04_3Statement'!I28+'04_3Statement'!I29-'04_3Statement'!I30"

ws["B10"] = "Cash negative? (1 = warning)"
ws["C10"] = '=IF(\'04_3Statement\'!D50<0,1,0)'
ws["D10"] = '=IF(\'04_3Statement\'!E50<0,1,0)'
ws["E10"] = '=IF(\'04_3Statement\'!F50<0,1,0)'
ws["F10"] = '=IF(\'04_3Statement\'!G50<0,1,0)'
ws["G10"] = '=IF(\'04_3Statement\'!H50<0,1,0)'
ws["H10"] = '=IF(\'04_3Statement\'!I50<0,1,0)'

ws["B11"] = "Interest coverage (EBIT / |Int|)  —  flag if < 2.0x"
ws["C11"] = '=IFERROR(\'04_3Statement\'!D26/ABS(\'04_3Statement\'!D27),"-")'
ws["D11"] = '=IFERROR(\'04_3Statement\'!E26/ABS(\'04_3Statement\'!E27),"-")'
ws["E11"] = '=IFERROR(\'04_3Statement\'!F26/ABS(\'04_3Statement\'!F27),"-")'
ws["F11"] = '=IFERROR(\'04_3Statement\'!G26/ABS(\'04_3Statement\'!G27),"-")'
ws["G11"] = '=IFERROR(\'04_3Statement\'!H26/ABS(\'04_3Statement\'!H27),"-")'
ws["H11"] = '=IFERROR(\'04_3Statement\'!I26/ABS(\'04_3Statement\'!I27),"-")'

for row in range(5, 12):
    ws.cell(row, 2).font = font_body
    for col in range(3, 9):
        cell = ws.cell(row, col)
        cell.font = font_link
        cell.alignment = center
        if row == 11:
            cell.number_format = FMT_X
        elif row == 10:
            cell.number_format = FMT_NUM
        elif row == 6:
            cell.number_format = FMT_MM1
        else:
            cell.number_format = "0.000"

ws["C6"].number_format = "0.000"

ws.conditional_formatting.add(
    "C5:H9",
    FormulaRule(formula=["AND(ISNUMBER(C5),ABS(C5)<0.05)"], fill=fill_green),
)
ws.conditional_formatting.add(
    "C5:H9",
    FormulaRule(formula=["AND(ISNUMBER(C5),ABS(C5)>=0.05)"], fill=fill_red),
)
ws.conditional_formatting.add(
    "C10:H10",
    FormulaRule(formula=["C10=1"], fill=fill_red),
)
ws.conditional_formatting.add(
    "C10:H10",
    FormulaRule(formula=["C10=0"], fill=fill_green),
)
ws.conditional_formatting.add(
    "C11:H11",
    FormulaRule(formula=["AND(ISNUMBER(C11),C11<2)"], fill=fill_amber),
)

section_bar(ws, 13, 2, 8, "HOW TO USE CHECKS LIKE AN ANALYST")
notes = [
    "Put this tab first when you send a file internally. VPs scroll here before they scroll the IS.",
    "A check cell is a formula that equals ZERO when the world is right. Color it with conditional formatting, never by hand.",
    "ABS() so −0.0003 and +0.0003 both pass. Use a tolerance (here 0.05) because floats are messy.",
    "If cash goes negative, you needed a revolver plug. Real models have a min-cash + draw / paydown revolver.",
    "Coverage, leverage, and liquidity flags belong here too — not buried in a footnote of the credit memo.",
    "Never delete a check because it is red. Fix the model.",
]
for i, n in enumerate(notes, start=14):
    ws.cell(i, 2, "▸").font = font_navy_b
    ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
    ws.cell(i, 3, n).font = font_body
    ws.row_dimensions[i].height = 18

ws["B21"] = "Aggregate break"
ws["C21"] = '=IF(AND(C6<0.05,SUM(C10:H10)=0),"ALL CLEAR","FIX BEFORE SEND")'
ws["C21"].font = Font(name="Arial", size=14, bold=True, color="008000")
ws.merge_cells("C21:H21")
ws.conditional_formatting.add(
    "C21",
    FormulaRule(formula=['C21="ALL CLEAR"'], fill=fill_green, font=Font(name="Arial", size=14, bold=True, color="006000")),
)
ws.conditional_formatting.add(
    "C21",
    FormulaRule(formula=['C21="FIX BEFORE SEND"'], fill=fill_red, font=Font(name="Arial", size=14, bold=True, color="9C0006")),
)

# ═══════════════════════════════════════════════════════════════════════════
# SHEET: 08 Plan
# ═══════════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("08_Plan")
banner(ws, "08  ·  14-DAY PLAN + EXTRA TIPS", "25–40 minutes a day. Track it. Bankers get good at this the same way pianists do — scales, then pieces.", 8)
width(ws, {"A": 3, "B": 12, "C": 18, "D": 52, "E": 18, "F": 14, "G": 16, "H": 16})

ws["B4"] = "DAY"
ws["C4"] = "FOCUS"
ws["D4"] = "SESSION"
ws["E4"] = "TAB"
ws["F4"] = "MINUTES"
ws["G4"] = "DONE? (Y)"
ws["H4"] = "NOTES"
for col in range(2, 9):
    ws.cell(4, col).font = font_white_b
    ws.cell(4, col).fill = fill_navy
    ws.cell(4, col).alignment = center

plan = [
    (5, "1", "Navigation", "Ctrl+Arrow, Ctrl+Shift+Arrow, F5, Ctrl+PgUp/Dn only. Cross every tab without a mouse.", "01_Shortcuts", 25),
    (6, "2", "Edit muscle", "F2, F4 anchors, Ctrl+D, Ctrl+R, Alt+=, Ctrl+Z. Rewrite one formula and copy it across years.", "01 + 04", 25),
    (7, "3", "Paste Special", "Values, formats, formulas, add/subtract, skip blanks, transpose (Alt E S E). Snapshot a case.", "01_Shortcuts", 25),
    (8, "4", "Format gym", "Fix tab 02 until it matches the AFTER block. Ctrl+1 custom format. Blue the constants (F5 Special).", "02_Format_Gym", 40),
    (9, "5", "Lookups", "Drills 01–07 on tab 03. INDEX/MATCH until you stop thinking about VLOOKUP.", "03_Formula_Dojo", 35),
    (10, "6", "Aggregation", "Drills 08–13. SUMIFS, AVERAGEIFS, LARGE, MEDIAN of a multiple.", "03_Formula_Dojo", 30),
    (11, "7", "Value math", "Drills 14–20. NPV vs Excel-NPV trap, IRR, SUMPRODUCT, TEXT labels.", "03_Formula_Dojo", 35),
    (12, "8", "3-statement", "Read every formula on tab 04 with F2 and Ctrl+[. Change FY26 growth from 8% to 3%. Explain what moved.", "04_3Statement", 40),
    (13, "9", "Break it", "Set tax rate to 0, capex % to 25%, payout to 150%. Watch checks go red. Then put it back.", "04 + 07", 30),
    (14, "10", "Sensitivity", "Rebuild the WACC × multiple grid with a real Data Table (Alt A W T). Toggle the case dropdown.", "05_Sensitivity", 40),
    (15, "11", "Comps", "Change three prices. Recalc implied Apex price. Write one sentence on why median ≠ mean here.", "06_Comps", 30),
    (16, "12", "Audit", "Ctrl+` formula view on tab 04. Ctrl+[ from NI to EBT to Revenue. Trace the cash line to the CFS.", "04 + 07", 30),
    (17, "13", "Print / polish", "Landscape, fit-to-width 1, print titles, header/footer, hide column N on the Dojo, freeze panes.", "all", 30),
    (18, "14", "Timed build", "Blank workbook. 45 minutes. Rebuild Apex IS + 4-year forecast from memory. Then compare to tab 04.", "blank file", 45),
]
for row, day, focus, session, tab, mins in plan:
    ws.cell(row, 2, day).font = font_navy_b
    ws.cell(row, 2).alignment = center
    ws.cell(row, 3, focus).font = font_navy_b
    ws.cell(row, 4, session).font = font_body
    ws.cell(row, 4).alignment = left
    ws.cell(row, 5, tab).font = font_hint
    ws.cell(row, 6, mins).font = font_formula
    ws.cell(row, 6).alignment = center
    ws.cell(row, 7).fill = fill_input
    ws.cell(row, 7).alignment = center
    ws.cell(row, 7).border = thin
    ws.row_dimensions[row].height = 32
    if row % 2 == 0:
        for col in range(2, 7):
            if ws.cell(row, col).fill.fgColor.rgb in (None, "00000000", "0"):
                ws.cell(row, col).fill = fill_off

ws["B20"] = "Total minutes"
ws["F20"] = "=SUM(F5:F18)"
ws["F20"].font = font_navy_b
ws["F20"].number_format = FMT_NUM
ws["G20"] = '=COUNTIF(G5:G18,"Y")&" / 14 days"'
ws["G20"].font = font_formula

section_bar(ws, 22, 2, 8, "TIPS THAT ARE NOT SHORTCUTS (THESE MATTER MORE)")
tips = [
    "Speed is a side effect of structure. A clean model with 40 shortcuts beats a messy model with 200.",
    "Read the 10-K footnote before you type the number. The formula is the easy part; the accounting is the job.",
    "Say the units out loud. '$mm, year ended Dec-31, diluted shares.' Most model errors are unit errors.",
    "If a VP asks 'what do I toggle?', the answer must be one yellow block, not seventeen hidden tabs.",
    "Build the check before you build the forecast. If you add the alarm after the fire, you will ship the fire.",
    "Circular interest (interest depends on average debt, debt depends on cash, cash depends on interest) is normal. Turn on iteration. Keep a toggle to break the circ when debugging.",
    "Don't merge cells in a grid you will ever copy from. Merged cells break Ctrl+Arrow and paste.",
    "Name the file like a person who will hate you in six weeks: ProjectApex_3stmt_v7_pre-IC_2026-08-29.xlsx.",
    "Print to PDF before you send. You will catch the column that still says 0.0000000001.",
    "The best next course after this file: rebuild a real public company from its 10-K (IS + BS + CFS) over a weekend. Then layer a DCF.",
]
for i, t in enumerate(tips, start=23):
    ws.cell(i, 2, f"{i-22:02d}").font = font_navy_b
    ws.cell(i, 2).alignment = center
    ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
    ws.cell(i, 3, t).font = font_body
    ws.row_dimensions[i].height = 22

section_bar(ws, 34, 2, 8, "WHEN YOU ARE READY FOR THE REAL THING")
ws.merge_cells("B35:H38")
ws["B35"] = (
    "Courses people actually use on the street: Wall Street Prep, Breaking Into Wall Street, Training The Street. "
    "They are expensive because they force reps, not because the functions are secret. "
    "Free complement: a public 10-K (start with a simple industrials name — you already have Apex as the toy version), "
    "plus the WSP / BIWS shortcut PDFs printed and taped next to the laptop. "
    "Skip fancy VBA until you can rebuild tab 04 from a blank sheet in under 20 minutes without looking. "
    "That is the line between 'I know Excel' and 'I can staff a live deal.'"
)
ws["B35"].font = font_body
ws["B35"].alignment = left_top

# Named ranges
wb.defined_names.add(DefinedName(name="TaxRate_FY26", attr_text="'04_3Statement'!$G$11"))
wb.defined_names.add(DefinedName(name="WACC_Base", attr_text="'05_Sensitivity'!$C$11"))
wb.defined_names.add(DefinedName(name="ExitMultiple", attr_text="'05_Sensitivity'!$C$6"))

# sheet order already created in sequence
# print settings already on banners

# tab colors
colors = {
    "00_Start": "1F4E79",
    "01_Shortcuts": "2E75B6",
    "02_Format_Gym": "C65911",
    "03_Formula_Dojo": "548235",
    "04_3Statement": "1F4E79",
    "05_Sensitivity": "7030A0",
    "06_Comps": "C45911",
    "07_Checks": "C00000",
    "08_Plan": "548235",
}
for name, col in colors.items():
    wb[name].sheet_properties.tabColor = col

out = "/home/workdir/artifacts/IB_Excel_Dojo.xlsx"
wb.save(out)
print("saved", out)
