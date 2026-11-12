"""
Build Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx  (v2)

A Power BI-ready marketing dashboard workbook populated with dummy data.

v2 adds an **Inputs** control panel: every table is formula-driven from the
Inputs sheet (fiscal year, data-as-of month, scenario, target leads, database
size, campaign / email / social / event assumptions, seasonality, database
mix).  Spin buttons, a month drop-down, scenario option buttons and campaign
check-boxes are added as native Excel form controls (via Excel COM) so the
figures can be changed with clicks instead of typing.

Every data sheet is an Excel Table (tbl_*) so Power BI imports it by name.
Column names of the tbl_* tables are unchanged from v1, so an existing Power
BI model keeps working after refresh.

Usage:
    python build_marketing_dashboard_v2.py [output_path]

Requires: openpyxl   (pywin32 + Excel: pre-calculates formulas and adds the
form controls; without Excel the workbook still works, inputs are typed).
"""

import random
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(2026)

YEAR = 2026                      # default fiscal year (input)
AS_OF_MONTH = 11                 # default "data as of" month index (input)
TARGET_LEADS = 1000              # default annual lead target (input)
SAUDI_CONTACTS = 1950            # default database size (input)
CAMPAIGNS = ["HVAC", "Data Centre", "Wastewater", "Hospitality", "Healthcare"]
PLATFORMS = ["LinkedIn", "Facebook", "WhatsApp Channel"]
EVENT_TYPES = ["Exhibition", "Conference", "Seminar", "Webinar", "Roadshow", "Networking"]
SEASON = [0.85, 0.90, 1.00, 1.05, 1.00, 0.95, 0.70, 0.75, 1.05, 1.15, 1.20, 1.10]

# Grundfos brand colours
BLUE_DARK = "11497B"
BLUE = "126AF3"
BLUE_LIGHT = "E9F1F7"

HEADER_FILL = PatternFill("solid", fgColor=BLUE_DARK)
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", bold=True, size=16, color=BLUE_DARK)
SUBTITLE_FONT = Font(name="Calibri", italic=True, size=10, color="666666")
KPI_FONT = Font(name="Calibri", bold=True, size=20, color=BLUE)
LABEL_FONT = Font(name="Calibri", bold=True, color=BLUE_DARK)
NOTE_FONT = Font(name="Calibri", italic=True, size=9, color="7F7F7F")
INPUT_FONT = Font(name="Calibri", bold=True, color=BLUE_DARK)
CALC_FONT = Font(name="Calibri", color="595959")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INPUT_SIDE = Side(style="thin", color=BLUE)
INPUT_BOX = Border(left=INPUT_SIDE, right=INPUT_SIDE, top=INPUT_SIDE, bottom=INPUT_SIDE)
INPUT_FILL = PatternFill("solid", fgColor=BLUE_LIGHT)
SECTION_FILL = PatternFill("solid", fgColor=BLUE_DARK)

FILL_GREEN = PatternFill("solid", fgColor="C6EFCE")
FILL_AMBER = PatternFill("solid", fgColor="FFEB9C")
FILL_RED = PatternFill("solid", fgColor="FFC7CE")
FILL_GREY = PatternFill("solid", fgColor="EDEDED")

FMT_INT = "#,##0"
FMT_MONEY = '#,##0 "SAR"'
FMT_PCT = "0.0%"
FMT_DATE = "dd-mmm-yyyy"
FMT_MONTH = "mmm-yyyy"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def add_table(ws, name, headers, rows, number_formats=None, widths=None,
              start_row=1, start_col=1, freeze=True, set_widths=True):
    """Write headers + rows, wrap them in an Excel Table and apply formats."""
    number_formats = number_formats or {}
    for c, h in enumerate(headers):
        cell = ws.cell(row=start_row, column=start_col + c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, row in enumerate(rows, start_row + 1):
        for c, val in enumerate(row):
            cell = ws.cell(row=r, column=start_col + c, value=val)
            fmt = number_formats.get(headers[c])
            if fmt:
                cell.number_format = fmt
    last_row = start_row + len(rows)
    ref = (f"{get_column_letter(start_col)}{start_row}:"
           f"{get_column_letter(start_col + len(headers) - 1)}{last_row}")
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)
    if freeze:
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    if set_widths:
        for i, h in enumerate(headers):
            w = (widths or {}).get(h, max(12, min(32, len(h) + 4)))
            ws.column_dimensions[get_column_letter(start_col + i)].width = w
    ws.row_dimensions[start_row].height = 30
    return last_row


def col(headers, name):
    return get_column_letter(headers.index(name) + 1)


def rng(headers, sheet, name, last_row, first_row=2):
    c = col(headers, name)
    return f"{sheet}!${c}${first_row}:${c}${last_row}"


def safe_div(num, den):
    return f"=IFERROR({num}/{den},0)"


def lit(lo, hi, nd=2):
    """Baked-in per-row variance factor (so dummy rows differ from each other)."""
    return f"{random.uniform(lo, hi):.{nd}f}"


def status_cf(ws, ref, mapping):
    """Colour cells by exact text (mapping: text -> fill)."""
    for text, fill in mapping.items():
        ws.conditional_formatting.add(
            ref, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill))


# --------------------------------------------------------------------------- #
# Inputs sheet (control panel)
# --------------------------------------------------------------------------- #
CONTROLS = []   # (kind, anchor_cell, options) – created later via Excel COM


def build_inputs(wb):
    ws = wb.create_sheet("Inputs")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    for c_, w in {"A": 2, "B": 36, "C": 14, "D": 16, "E": 46, "F": 16, "G": 16, "H": 16,
                  "I": 14, "J": 14, "K": 16, "L": 7, "M": 16, "N": 14}.items():
        ws.column_dimensions[c_].width = w

    def name(n, ref):
        wb.defined_names[n] = DefinedName(n, attr_text=f"Inputs!{ref}")

    def section(row, text):
        for c_ in range(2, 12):
            ws.cell(row=row, column=c_).fill = SECTION_FILL
        cell = ws.cell(row=row, column=2, value=text)
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        ws.row_dimensions[row].height = 20

    def inp(addr, value, fmt=None):
        c = ws[addr]
        c.value = value
        c.fill, c.border, c.font = INPUT_FILL, INPUT_BOX, INPUT_FONT
        c.alignment = Alignment(horizontal="center")
        if fmt:
            c.number_format = fmt
        return c

    def calc(addr, value, fmt=None):
        c = ws[addr]
        c.value = value
        c.font = CALC_FONT
        c.alignment = Alignment(horizontal="center")
        if fmt:
            c.number_format = fmt
        return c

    def label(row, text, note=None):
        ws.cell(row=row, column=2, value=text).font = LABEL_FONT
        if note:
            ws.cell(row=row, column=5, value=note).font = NOTE_FONT

    def style_table_inputs(ref, pct_cols=(), fmts=None):
        """Input styling + number formats for the body of a small inputs table."""
        fmts = fmts or {}
        for row in ws[ref]:
            for c in row:
                c.fill, c.border = INPUT_FILL, INPUT_BOX
                if c.column_letter in fmts:
                    c.number_format = fmts[c.column_letter]

    ws["B2"] = "Control panel – change the inputs, every sheet updates"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = ("Light-blue cells are inputs: use the buttons / drop-down or type a value. "
                "Grey cells are calculated. Dummy data – all figures are fictional.")
    ws["B3"].font = SUBTITLE_FONT

    # ---- 1. General settings ------------------------------------------- #
    section(5, "1   General settings")
    label(6, "Fiscal year", "▲▼ 2024–2030. All month, event and activity dates follow the year.")
    inp("C6", YEAR, "0")
    name("Inp_Year", "$C$6")
    CONTROLS.append(("spinner", "D6", dict(linked="$C$6", min=2024, max=2030, step=1)))

    label(7, "Data as of (last closed month)",
          "Pick the month. Actuals after it are 0 and event / activity status follows it.")
    calc("L7", AS_OF_MONTH, "0")
    name("Inp_AsOfIdx", "$L$7")
    calc("C7", "=EOMONTH(DATE(Inp_Year,Inp_AsOfIdx,1),0)", FMT_DATE)
    name("Inp_AsOf", "$C$7")
    CONTROLS.append(("dropdown", "D7", dict(linked="$L$7", list="Inputs!$M$6:$M$17", lines=12)))

    label(8, "Scenario", None)
    calc("L8", 2, "0")
    name("Inp_ScenarioIdx", "$L$8")
    calc("C8", "=INDEX(Inp_ScenarioNames,Inp_ScenarioIdx)")
    name("Inp_ScenarioName", "$C$8")
    CONTROLS.append(("options", "D8", dict(linked="$L$8", captions=["Conservative", "Base", "Optimistic"],
                                           default=2)))
    label(9, "Scenario lead multiplier", "Applied to campaign and event leads. Edit the multipliers in the list at the right (M–N).")
    calc("C9", "=INDEX(Inp_ScenarioMults,Inp_ScenarioIdx)", "0.00")
    name("Inp_ScenarioMult", "$C$9")

    label(10, "Target leads (annual)", "▲▼ steps of 50. Split evenly per month on Overview; drives Targets and Dashboard.")
    inp("C10", TARGET_LEADS, FMT_INT)
    name("Inp_TargetLeads", "$C$10")
    CONTROLS.append(("spinner", "D10", dict(linked="$C$10", min=0, max=30000, step=50)))

    label(11, "Saudi database contacts", "▲▼ steps of 50. Split by the segment / region mix in section 8.")
    inp("C11", SAUDI_CONTACTS, FMT_INT)
    name("Inp_Contacts", "$C$11")
    CONTROLS.append(("spinner", "D11", dict(linked="$C$11", min=0, max=30000, step=50)))

    # ---- 2. KPI targets ------------------------------------------------- #
    section(13, "2   KPI targets (Targets sheet scorecard)")
    tgt_rows = [
        ("MQLs", 550, "MQLs", "Inp_Tgt_MQLs"),
        ("SQLs", 250, "SQLs", "Inp_Tgt_SQLs"),
        ("Saudi database contacts", 2000, "contacts", "Inp_Tgt_Contacts"),
        ("Email open rate", 0.30, "%", "Inp_Tgt_EmailOpen"),
        ("Email click rate", 0.05, "%", "Inp_Tgt_EmailClick"),
        ("Social engagements", 45000, "engagements", "Inp_Tgt_SocialEng"),
        ("LinkedIn followers", 10000, "followers", "Inp_Tgt_LinkedIn"),
        ("Events delivered", 12, "events", "Inp_Tgt_Events"),
        ("Budget utilisation", 0.95, "%", "Inp_Tgt_BudgetUtil"),
    ]
    add_table(ws, "tbl_Inputs_Targets", ["KPI", "Target", "Unit"],
              [[k, v, u] for k, v, u, _ in tgt_rows], start_row=14, start_col=2,
              freeze=False, set_widths=False)
    for i, (_, v, u, nm) in enumerate(tgt_rows):
        r = 15 + i
        inp(f"C{r}", v, FMT_PCT if u == "%" else FMT_INT)
        name(nm, f"$C${r}")

    # ---- 3. Campaign assumptions ---------------------------------------- #
    section(25, "3   Campaign assumptions (tick = campaign active)")
    camp_hdr = ["Campaign", "Active", "Segment", "Primary_Channel", "Annual_Budget",
                "Monthly_Leads_Base", "Monthly_Impressions_Base", "Lead_to_MQL", "MQL_to_SQL",
                "Email_Sends_per_Month"]
    camp_rows = [
        ["HVAC", True, "Commercial Buildings", "LinkedIn Ads", 32000, 16, 42000, 0.57, 0.47, 520],
        ["Data Centre", True, "Data Centres", "Trade Media", 38000, 10, 28000, 0.57, 0.47, 450],
        ["Wastewater", True, "Water Utilities", "Events", 26000, 9, 22000, 0.57, 0.47, 480],
        ["Hospitality", True, "Hotels & Resorts", "Email", 22000, 11, 30000, 0.57, 0.47, 500],
        ["Healthcare", True, "Hospitals", "LinkedIn Ads", 20000, 8, 24000, 0.57, 0.47, 450],
    ]
    add_table(ws, "tbl_Inputs_Campaigns", camp_hdr, camp_rows, start_row=26, start_col=2,
              freeze=False, set_widths=False)
    style_table_inputs("C27:K31", fmts={"F": FMT_MONEY, "G": FMT_INT, "H": FMT_INT,
                                        "I": FMT_PCT, "J": FMT_PCT, "K": FMT_INT})
    for r in range(27, 32):
        ws[f"C{r}"].number_format = ";;;"      # value hidden behind the check-box
        CONTROLS.append(("checkbox", f"C{r}", dict(linked=f"$C${r}")))
    for nm, c_ in [("Inp_Camp_Names", "B"), ("Inp_Camp_Active", "C"), ("Inp_Camp_Seg", "D"),
                   ("Inp_Camp_Ch", "E"), ("Inp_Camp_Budget", "F"), ("Inp_Camp_Leads", "G"),
                   ("Inp_Camp_Imp", "H"), ("Inp_Camp_L2M", "I"), ("Inp_Camp_M2S", "J"),
                   ("Inp_Camp_Email", "K")]:
        name(nm, f"${c_}$27:${c_}$31")

    # ---- 4. Email assumptions ------------------------------------------- #
    section(33, "4   Email assumptions")
    for r, (lbl, pct, nm, note) in enumerate([
        ("Delivery rate", 97, "Inp_Em_Deliv", "▲▼ 1 % steps (or overwrite the % cell)."),
        ("Open rate (of delivered)", 30, "Inp_Em_Open", "▲▼ 1 % steps. Base open rate; each row varies around it."),
        ("Click-to-open rate (CTR)", 16, "Inp_Em_CTO", "▲▼ 1 % steps. Click rate = open rate × click-to-open."),
    ], start=34):
        label(r, lbl, note)
        calc(f"L{r}", pct, "0")
        inp(f"C{r}", f"=L{r}/100", FMT_PCT)
        name(nm, f"$C${r}")
        CONTROLS.append(("spinner", f"D{r}", dict(linked=f"$L${r}", min=0, max=100, step=1)))
    label(37, "Unsubscribe rate (of delivered)", "Type a value, e.g. 0.4 %.")
    inp("C37", 0.004, "0.00%")
    name("Inp_Em_Unsub", "$C$37")

    # ---- 5. Social media assumptions ------------------------------------ #
    section(39, "5   Social media assumptions (tick = platform active)")
    soc_hdr = ["Platform", "Active", "Starting_Followers", "New_Followers_per_Month", "Posts_per_Month",
               "Impressions_per_Month", "Like_Rate", "Comment_Rate", "Repost_Rate", "Click_Rate"]
    soc_rows = [
        ["LinkedIn", True, 8200, 160, 14, 38000, 0.028, 0.004, 0.006, 0.012],
        ["Facebook", True, 5400, 70, 10, 18000, 0.022, 0.003, 0.004, 0.007],
        ["WhatsApp Channel", True, 1100, 55, 8, 6000, 0.060, 0.000, 0.020, 0.030],
    ]
    add_table(ws, "tbl_Inputs_Social", soc_hdr, soc_rows, start_row=40, start_col=2,
              freeze=False, set_widths=False)
    style_table_inputs("C41:K43", fmts={"D": FMT_INT, "E": FMT_INT, "F": FMT_INT, "G": FMT_INT,
                                        "H": FMT_PCT, "I": FMT_PCT, "J": FMT_PCT, "K": FMT_PCT})
    for r in range(41, 44):
        ws[f"C{r}"].number_format = ";;;"
        CONTROLS.append(("checkbox", f"C{r}", dict(linked=f"$C${r}")))
    for nm, c_ in [("Inp_Soc_Names", "B"), ("Inp_Soc_Active", "C"), ("Inp_Soc_Start", "D"),
                   ("Inp_Soc_NewF", "E"), ("Inp_Soc_Posts", "F"), ("Inp_Soc_Imp", "G"),
                   ("Inp_Soc_Like", "H"), ("Inp_Soc_Com", "I"), ("Inp_Soc_Rep", "J"),
                   ("Inp_Soc_Clk", "K")]:
        name(nm, f"${c_}$41:${c_}$43")

    # ---- 6. Event assumptions ------------------------------------------- #
    section(45, "6   Event assumptions (per event type)")
    evt_rows = [
        ["Exhibition", 120000, 420, 0.11],
        ["Conference", 60000, 260, 0.12],
        ["Seminar", 25000, 90, 0.14],
        ["Webinar", 8000, 150, 0.10],
        ["Roadshow", 45000, 120, 0.12],
        ["Networking", 20000, 60, 0.15],
    ]
    add_table(ws, "tbl_Inputs_Event_Types", ["Event_Type", "Budget", "Attendees_Base", "Lead_Rate"],
              evt_rows, start_row=46, start_col=2, freeze=False, set_widths=False)
    style_table_inputs("C47:E52", fmts={"C": FMT_MONEY, "D": FMT_INT, "E": FMT_PCT})
    for nm, c_ in [("Inp_Evt_Types", "B"), ("Inp_Evt_Budget", "C"), ("Inp_Evt_Att", "D"),
                   ("Inp_Evt_LeadRate", "E")]:
        name(nm, f"${c_}$47:${c_}$52")

    # ---- 7. Seasonality ------------------------------------------------- #
    section(54, "7   Seasonality – monthly activity factor (1.00 = average month)")
    sea_rows = [[m, f'=TEXT(DATE(Inp_Year,{m},1),"mmm")', SEASON[m - 1]] for m in range(1, 13)]
    add_table(ws, "tbl_Inputs_Seasonality", ["Month_No", "Month", "Factor"], sea_rows,
              start_row=55, start_col=2, freeze=False, set_widths=False)
    style_table_inputs("D56:D67", fmts={"D": "0.00"})
    for r in range(56, 68):
        ws[f"C{r}"].font = CALC_FONT
    name("Inp_Season", "$D$56:$D$67")

    # ---- 8. Database mix ------------------------------------------------ #
    section(69, "8   Database mix – segment and region shares must each total 100 %")
    seg_rows = [
        ["HVAC", 0.26, "End User / Facility Owner"],
        ["Data Centre", 0.12, "End User / Facility Owner"],
        ["Wastewater", 0.14, "End User / Facility Owner"],
        ["Hospitality", 0.13, "End User / Facility Owner"],
        ["Healthcare", 0.11, "End User / Facility Owner"],
        ["Industrial", 0.10, "End User / Facility Owner"],
        ["Consultants & EPCs", 0.14, "Consultant / EPC"],
    ]
    add_table(ws, "tbl_Inputs_Segments", ["Segment", "Share", "Contact_Type"], seg_rows,
              start_row=70, start_col=2, freeze=False, set_widths=False)
    style_table_inputs("C71:D77", fmts={"C": FMT_PCT})
    name("Inp_Seg_Names", "$B$71:$B$77")
    name("Inp_Seg_Share", "$C$71:$C$77")
    name("Inp_Seg_Type", "$D$71:$D$77")
    ws["B78"] = "Check – segment shares"
    ws["B78"].font = NOTE_FONT
    calc("C78", '=IF(ABS(SUM(C71:C77)-1)<0.0001,"OK 100%","Must total 100% – now "&TEXT(SUM(C71:C77),"0%"))')
    ws["C78"].alignment = Alignment(horizontal="left")

    reg_rows = [["Central (Riyadh)", 0.48], ["Western (Jeddah/Makkah)", 0.30], ["Eastern (Dammam/Khobar)", 0.22]]
    add_table(ws, "tbl_Inputs_Regions", ["Region", "Share"], reg_rows,
              start_row=80, start_col=2, freeze=False, set_widths=False)
    style_table_inputs("C81:C83", fmts={"C": FMT_PCT})
    name("Inp_Reg_Names", "$B$81:$B$83")
    name("Inp_Reg_Share", "$C$81:$C$83")
    ws["B84"] = "Check – region shares"
    ws["B84"].font = NOTE_FONT
    calc("C84", '=IF(ABS(SUM(C81:C83)-1)<0.0001,"OK 100%","Must total 100% – now "&TEXT(SUM(C81:C83),"0%"))')
    ws["C84"].alignment = Alignment(horizontal="left")
    for addr in ("C78", "C84"):
        ws.conditional_formatting.add(addr, FormulaRule(formula=[f'LEFT({addr},2)="OK"'], fill=FILL_GREEN))
        ws.conditional_formatting.add(addr, FormulaRule(formula=[f'LEFT({addr},2)<>"OK"'], fill=FILL_RED))

    # ---- Helper lists (right-hand side) -------------------------------- #
    ws["M5"] = "Lists used by the controls"
    ws["M5"].font = LABEL_FONT
    for i in range(12):
        calc(f"M{6 + i}", f'=TEXT(DATE(Inp_Year,{i + 1},1),"mmm-yyyy")')
        ws[f"M{6 + i}"].alignment = Alignment(horizontal="left")
    name("Inp_MonthList", "$M$6:$M$17")
    ws["M19"], ws["N19"] = "Scenario", "Lead multiplier"
    ws["M19"].font = ws["N19"].font = LABEL_FONT
    for i, (nm_, mult) in enumerate([("Conservative", 0.85), ("Base", 1.00), ("Optimistic", 1.20)]):
        ws[f"M{20 + i}"] = nm_
        inp(f"N{20 + i}", mult, "0.00")
    name("Inp_ScenarioNames", "$M$20:$M$22")
    name("Inp_ScenarioMults", "$N$20:$N$22")

    # ---- Data validation (keeps typed inputs sane) ---------------------- #
    dv_year = DataValidation(type="whole", operator="between", formula1="2020", formula2="2035",
                             error="Enter a year between 2020 and 2035", showErrorMessage=True)
    dv_month = DataValidation(type="whole", operator="between", formula1="1", formula2="12", showErrorMessage=True)
    dv_scn = DataValidation(type="whole", operator="between", formula1="1", formula2="3", showErrorMessage=True)
    dv_pos = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", showErrorMessage=True)
    dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                            error="Enter a percentage between 0% and 100%", showErrorMessage=True)
    dv_bool = DataValidation(type="list", formula1='"TRUE,FALSE"', showErrorMessage=True)
    for dv in (dv_year, dv_month, dv_scn, dv_pos, dv_pct, dv_bool):
        ws.add_data_validation(dv)
    dv_year.add("C6")
    dv_month.add("L7")
    dv_scn.add("L8")
    for ref in ("C10", "C11", "C15:C17", "C20:C22", "F27:H31", "K27:K31", "D41:G43", "C47:D52", "L34:L36"):
        dv_pos.add(ref)
    for ref in ("C18:C19", "C23", "I27:J31", "H41:K43", "E47:E52", "C71:C77", "C81:C83", "C37"):
        dv_pct.add(ref)
    dv_bool.add("C27:C31")
    dv_bool.add("C41:C43")
    return ws


# --------------------------------------------------------------------------- #
# Build
# --------------------------------------------------------------------------- #
def build(path: Path):
    wb = Workbook()
    wb.remove(wb.active)
    build_inputs(wb)

    # ------------------------------------------------------------------ #
    # 1. Campaigns  (5 campaigns × 12 months, driven by Inputs §3)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Campaigns")
    camp_headers = [
        "Month", "Month_Label", "Campaign", "Segment", "Primary_Channel",
        "Impressions", "Clicks", "Leads", "MQLs", "SQLs", "Budget", "Spend",
        "CTR", "Cost_per_Lead", "Lead_to_MQL", "MQL_to_SQL",
    ]
    camp_rows = []
    for m in range(1, 13):
        for camp in CAMPAIGNS:
            r = len(camp_rows) + 2
            mi = f"MATCH($C{r},Inp_Camp_Names,0)"
            act = f"AND(INDEX(Inp_Camp_Active,{mi}),$A{r}<=Inp_AsOf)"
            sea = f"INDEX(Inp_Season,MONTH($A{r}))"
            camp_rows.append([
                f"=DATE(Inp_Year,{m},1)",
                f'=TEXT(A{r},"mmm-yyyy")',
                camp,
                f"=INDEX(Inp_Camp_Seg,{mi})",
                f"=INDEX(Inp_Camp_Ch,{mi})",
                f"=IF({act},ROUND(INDEX(Inp_Camp_Imp,{mi})*{sea}*{lit(0.75, 1.25)},0),0)",
                f"=ROUND(F{r}*{lit(0.012, 0.03, 4)},0)",
                f"=IF({act},ROUND(INDEX(Inp_Camp_Leads,{mi})*{sea}*Inp_ScenarioMult*{lit(0.65, 1.35)},0),0)",
                f"=ROUND(H{r}*INDEX(Inp_Camp_L2M,{mi})*{lit(0.8, 1.2)},0)",
                f"=ROUND(I{r}*INDEX(Inp_Camp_M2S,{mi})*{lit(0.8, 1.2)},0)",
                f"=IF(INDEX(Inp_Camp_Active,{mi}),INDEX(Inp_Camp_Budget,{mi})/12,0)",
                f"=IF({act},ROUND(K{r}*{lit(0.75, 1.08)},-1),0)",
                safe_div(f"G{r}", f"F{r}"),
                safe_div(f"L{r}", f"H{r}"),
                safe_div(f"I{r}", f"H{r}"),
                safe_div(f"J{r}", f"I{r}"),
            ])
    camp_last = add_table(
        ws, "tbl_Campaigns", camp_headers, camp_rows,
        {"Month": FMT_MONTH, "Impressions": FMT_INT, "Clicks": FMT_INT, "Leads": FMT_INT,
         "MQLs": FMT_INT, "SQLs": FMT_INT, "Budget": FMT_MONEY, "Spend": FMT_MONEY,
         "CTR": FMT_PCT, "Cost_per_Lead": FMT_MONEY, "Lead_to_MQL": FMT_PCT, "MQL_to_SQL": FMT_PCT},
        {"Segment": 22, "Primary_Channel": 16},
    )

    # ------------------------------------------------------------------ #
    # 2. Events  (Leads, MQLs, SQLs – driven by Inputs §6 and as-of date)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Events")
    ev_headers = [
        "Event_ID", "Event_Name", "Event_Date", "Month", "City", "Event_Type", "Campaign",
        "Attendees", "Leads", "MQLs", "SQLs", "Budget", "Actual_Cost",
        "Cost_per_Lead", "Lead_Conversion", "Status",
    ]
    events = [
        ("Saudi HVAC-R Expo",                 (1, 27),  "Riyadh", "Exhibition", "HVAC"),
        ("Data Centre World Saudi",           (2, 17),  "Riyadh", "Exhibition", "Data Centre"),
        ("Grundfos Distributed Pumping Day",  (3, 10),  "Jeddah", "Seminar",    "HVAC"),
        ("Saudi Water Forum",                 (4, 21),  "Riyadh", "Conference", "Wastewater"),
        ("Hospitality Innovation Summit",     (5, 12),  "Jeddah", "Conference", "Hospitality"),
        ("Healthcare Facilities Webinar",     (6, 9),   "Online", "Webinar",    "Healthcare"),
        ("Energy Efficiency Roadshow",        (9, 15),  "Dammam", "Roadshow",   "HVAC"),
        ("Wastewater Lifecycle Cost Webinar", (9, 29),  "Online", "Webinar",    "Wastewater"),
        ("Saudi Build",                       (10, 20), "Riyadh", "Exhibition", "Hospitality"),
        ("Data Centre Cooling Masterclass",   (11, 10), "Riyadh", "Seminar",    "Data Centre"),
        ("Global Health Exhibition",          (11, 24), "Riyadh", "Exhibition", "Healthcare"),
        ("Year-End Consultant Appreciation",  (12, 8),  "Jeddah", "Networking", "HVAC"),
    ]
    ev_rows = []
    for n, (name_, (mm, dd), city, etype, camp) in enumerate(events, 1):
        r = n + 1
        ti = f"MATCH($F{r},Inp_Evt_Types,0)"
        ci = f"MATCH($G{r},Inp_Camp_Names,0)"
        deliv = f'$P{r}="Delivered"'
        ev_rows.append([
            f"EV-{n:03d}", name_,
            f"=DATE(Inp_Year,{mm},{dd})",
            f"=DATE(YEAR(C{r}),MONTH(C{r}),1)",
            city, etype, camp,
            f"=IF({deliv},ROUND(INDEX(Inp_Evt_Att,{ti})*{lit(0.8, 1.2)},0),0)",
            f"=ROUND(H{r}*INDEX(Inp_Evt_LeadRate,{ti})*Inp_ScenarioMult*{lit(0.8, 1.2)},0)",
            f"=ROUND(I{r}*INDEX(Inp_Camp_L2M,{ci})*{lit(0.8, 1.2)},0)",
            f"=ROUND(J{r}*INDEX(Inp_Camp_M2S,{ci})*{lit(0.8, 1.2)},0)",
            f"=INDEX(Inp_Evt_Budget,{ti})",
            f"=IF({deliv},ROUND(L{r}*{lit(0.85, 1.05)},-2),0)",
            safe_div(f"M{r}", f"I{r}"),
            safe_div(f"I{r}", f"H{r}"),
            f'=IF(C{r}<=Inp_AsOf,"Delivered","Planned")',
        ])
    ev_last = add_table(
        ws, "tbl_Events", ev_headers, ev_rows,
        {"Event_Date": FMT_DATE, "Month": FMT_MONTH, "Attendees": FMT_INT, "Leads": FMT_INT,
         "MQLs": FMT_INT, "SQLs": FMT_INT, "Budget": FMT_MONEY, "Actual_Cost": FMT_MONEY,
         "Cost_per_Lead": FMT_MONEY, "Lead_Conversion": FMT_PCT},
        {"Event_Name": 34},
    )
    status_cf(ws, f"P2:P{ev_last}", {"Delivered": FILL_GREEN, "Planned": FILL_GREY})

    # ------------------------------------------------------------------ #
    # 3. Email performance (Inputs §4 + per-campaign sends)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Email")
    em_headers = [
        "Month", "Month_Label", "Campaign", "Email_Name", "Sent", "Delivered", "Bounced",
        "Opens", "Clicks", "Unsubscribes",
        "Delivery_Rate", "Open_Rate", "Click_Rate", "CTR_Click_to_Open", "Unsub_Rate",
    ]
    em_topics = {
        "HVAC": "Distributed Pumping Energy Savings",
        "Data Centre": "Cooling Reliability Insights",
        "Wastewater": "Lifecycle Cost of Pumping Stations",
        "Hospitality": "Guest Comfort, Lower OPEX",
        "Healthcare": "Hygienic Water Systems for Hospitals",
    }
    em_rows = []
    for m in range(1, 13):
        for camp in CAMPAIGNS:
            r = len(em_rows) + 2
            mi = f"MATCH($C{r},Inp_Camp_Names,0)"
            act = f"AND(INDEX(Inp_Camp_Active,{mi}),$A{r}<=Inp_AsOf)"
            sea = f"INDEX(Inp_Season,MONTH($A{r}))"
            em_rows.append([
                f"=DATE(Inp_Year,{m},1)",
                f'=TEXT(A{r},"mmm-yyyy")',
                camp,
                f'="{em_topics[camp]} – "&B{r}',
                f"=IF({act},ROUND(INDEX(Inp_Camp_Email,{mi})*{sea}*{lit(0.85, 1.15)},0),0)",
                f"=MIN(E{r},ROUND(E{r}*Inp_Em_Deliv*{lit(0.99, 1.01, 3)},0))",
                f"=E{r}-F{r}",
                f"=ROUND(F{r}*Inp_Em_Open*{lit(0.75, 1.25)},0)",
                f"=ROUND(H{r}*Inp_Em_CTO*{lit(0.7, 1.3)},0)",
                f"=ROUND(F{r}*Inp_Em_Unsub*{lit(0, 2)},0)",
                safe_div(f"F{r}", f"E{r}"),
                safe_div(f"H{r}", f"F{r}"),
                safe_div(f"I{r}", f"F{r}"),
                safe_div(f"I{r}", f"H{r}"),
                safe_div(f"J{r}", f"F{r}"),
            ])
    em_last = add_table(
        ws, "tbl_Email", em_headers, em_rows,
        {"Month": FMT_MONTH, "Sent": FMT_INT, "Delivered": FMT_INT, "Bounced": FMT_INT,
         "Opens": FMT_INT, "Clicks": FMT_INT, "Delivery_Rate": FMT_PCT, "Open_Rate": FMT_PCT,
         "Click_Rate": FMT_PCT, "CTR_Click_to_Open": FMT_PCT, "Unsub_Rate": FMT_PCT},
        {"Email_Name": 46},
    )

    # ------------------------------------------------------------------ #
    # 4. Social media (Inputs §5)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Social_Media")
    so_headers = [
        "Month", "Month_Label", "Platform", "Followers", "New_Followers", "Posts",
        "Impressions", "Likes", "Comments", "Reposts", "Clicks",
        "Total_Engagements", "Engagement_Rate", "Avg_Engagement_per_Post",
    ]
    so_rows = []
    npl = len(PLATFORMS)
    for m in range(1, 13):
        for p in PLATFORMS:
            r = len(so_rows) + 2
            pi = f"MATCH($C{r},Inp_Soc_Names,0)"
            act = f"AND(INDEX(Inp_Soc_Active,{pi}),$A{r}<=Inp_AsOf)"
            sea = f"INDEX(Inp_Season,MONTH($A{r}))"
            prev = f"INDEX(Inp_Soc_Start,{pi})" if m == 1 else f"D{r - npl}"
            so_rows.append([
                f"=DATE(Inp_Year,{m},1)",
                f'=TEXT(A{r},"mmm-yyyy")',
                p,
                f"={prev}+E{r}",
                f"=IF({act},ROUND(INDEX(Inp_Soc_NewF,{pi})*{sea}*{lit(0.7, 1.3)},0),0)",
                f"=IF({act},MAX(3,ROUND(INDEX(Inp_Soc_Posts,{pi})*{lit(0.8, 1.2)},0)),0)",
                f"=IF({act},ROUND(INDEX(Inp_Soc_Imp,{pi})*{sea}*{lit(0.75, 1.25)},0),0)",
                f"=ROUND(G{r}*INDEX(Inp_Soc_Like,{pi})*{lit(0.8, 1.2)},0)",
                f"=ROUND(G{r}*INDEX(Inp_Soc_Com,{pi})*{lit(0.7, 1.3)},0)",
                f"=ROUND(G{r}*INDEX(Inp_Soc_Rep,{pi})*{lit(0.7, 1.3)},0)",
                f"=ROUND(G{r}*INDEX(Inp_Soc_Clk,{pi})*{lit(0.8, 1.2)},0)",
                f"=H{r}+I{r}+J{r}+K{r}",
                safe_div(f"L{r}", f"G{r}"),
                safe_div(f"L{r}", f"F{r}"),
            ])
    so_last = add_table(
        ws, "tbl_Social_Media", so_headers, so_rows,
        {"Month": FMT_MONTH, "Followers": FMT_INT, "New_Followers": FMT_INT, "Posts": FMT_INT,
         "Impressions": FMT_INT, "Likes": FMT_INT, "Comments": FMT_INT, "Reposts": FMT_INT,
         "Clicks": FMT_INT, "Total_Engagements": FMT_INT, "Engagement_Rate": FMT_PCT,
         "Avg_Engagement_per_Post": "#,##0.0"},
        {"Platform": 18},
    )

    # ------------------------------------------------------------------ #
    # 5. Saudi database  (total = Inp_Contacts, split by Inputs §8)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Database")
    db_headers = ["Country", "Region", "Segment", "Contact_Type", "Contacts",
                  "Opted_In", "Engaged_90d", "Opt_In_Rate", "Engagement_Rate"]
    regions = ["Central (Riyadh)", "Western (Jeddah/Makkah)", "Eastern (Dammam/Khobar)"]
    segments = CAMPAIGNS + ["Industrial", "Consultants & EPCs"]
    combos = [(seg, reg) for seg in segments for reg in regions]
    db_rows = []
    last = len(combos) + 1
    for n, (seg, reg) in enumerate(combos):
        r = n + 2
        si = f"MATCH($C{r},Inp_Seg_Names,0)"
        ri = f"MATCH($B{r},Inp_Reg_Names,0)"
        contacts = (f"=Inp_Contacts-SUM(E2:E{last - 1})" if r == last else
                    f"=ROUND(Inp_Contacts*INDEX(Inp_Seg_Share,{si})*INDEX(Inp_Reg_Share,{ri}),0)")
        db_rows.append([
            "Saudi Arabia", reg, seg, f"=INDEX(Inp_Seg_Type,{si})", contacts,
            f"=ROUND(E{r}*{lit(0.7, 0.9)},0)",
            f"=ROUND(F{r}*{lit(0.3, 0.5)},0)",
            safe_div(f"F{r}", f"E{r}"), safe_div(f"G{r}", f"F{r}"),
        ])
    db_last = add_table(
        ws, "tbl_Database", db_headers, db_rows,
        {"Contacts": FMT_INT, "Opted_In": FMT_INT, "Engaged_90d": FMT_INT,
         "Opt_In_Rate": FMT_PCT, "Engagement_Rate": FMT_PCT},
        {"Region": 26, "Segment": 20, "Contact_Type": 26},
    )
    ws.cell(row=db_last + 2, column=4, value="Total Saudi contacts").font = LABEL_FONT
    c = ws.cell(row=db_last + 2, column=5, value=f"=SUM(E2:E{db_last})")
    c.font = LABEL_FONT
    c.number_format = FMT_INT

    # ------------------------------------------------------------------ #
    # 6. Budget
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Budget")
    bu_headers = ["Budget_Line", "Category", "Annual_Budget", "Spend_YTD", "Remaining",
                  "Pct_Used", "Status"]
    status_f = lambda r: f'=IF(F{r}>1,"Over budget",IF(F{r}>0.9,"Watch","On track"))'  # noqa: E731
    bu_rows = []
    r = 2
    for camp in CAMPAIGNS:
        bu_rows.append([
            f"{camp} Campaign", "Campaign",
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Budget", camp_last)},'
            f'{rng(camp_headers, "Campaigns", "Campaign", camp_last)},"{camp}")',
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Spend", camp_last)},'
            f'{rng(camp_headers, "Campaigns", "Campaign", camp_last)},"{camp}")',
            f"=C{r}-D{r}", safe_div(f"D{r}", f"C{r}"), status_f(r),
        ])
        r += 1
    bu_rows.append([
        "Events & Exhibitions", "Events",
        f'=SUM({rng(ev_headers, "Events", "Budget", ev_last)})',
        f'=SUM({rng(ev_headers, "Events", "Actual_Cost", ev_last)})',
        f"=C{r}-D{r}", safe_div(f"D{r}", f"C{r}"), status_f(r),
    ])
    r += 1
    for line, cat, bud, spent in [
        ("Social Media Content & Paid", "Digital", 90000, 71500),
        ("Email Platform & Automation", "Digital", 36000, 31200),
        ("Database & Market Intelligence", "Data", 45000, 38900),
        ("Brand & Collateral (AR/EN)", "Brand", 60000, 44800),
        ("Agency & Creative", "Agency", 120000, 101300),
    ]:
        bu_rows.append([line, cat, bud, spent, f"=C{r}-D{r}", safe_div(f"D{r}", f"C{r}"), status_f(r)])
        r += 1
    bu_last = add_table(
        ws, "tbl_Budget", bu_headers, bu_rows,
        {"Annual_Budget": FMT_MONEY, "Spend_YTD": FMT_MONEY, "Remaining": FMT_MONEY, "Pct_Used": FMT_PCT},
        {"Budget_Line": 32, "Status": 14},
    )
    # the five fixed lines are typed inputs – mark them like the Inputs sheet
    for rr in range(bu_last - 4, bu_last + 1):
        for cc in "CD":
            ws[f"{cc}{rr}"].fill, ws[f"{cc}{rr}"].border = INPUT_FILL, INPUT_BOX
    ws.cell(row=bu_last + 2, column=2, value="Total").font = LABEL_FONT
    for cc in "CDE":
        c = ws[f"{cc}{bu_last + 2}"]
        c.value = f"=SUM({cc}2:{cc}{bu_last})"
        c.font = LABEL_FONT
        c.number_format = FMT_MONEY
    c = ws[f"F{bu_last + 2}"]
    c.value = safe_div(f"D{bu_last + 2}", f"C{bu_last + 2}")
    c.font = LABEL_FONT
    c.number_format = FMT_PCT
    ws.cell(row=bu_last + 3, column=2,
            value="Light-blue cells (fixed budget lines) are typed inputs; the rest roll up from Campaigns and Events.").font = NOTE_FONT
    status_cf(ws, f"G2:G{bu_last}", {"On track": FILL_GREEN, "Watch": FILL_AMBER, "Over budget": FILL_RED})
    ws.conditional_formatting.add(f"F2:F{bu_last}", DataBarRule(start_type="num", start_value=0,
                                                                 end_type="num", end_value=1, color=BLUE))

    # ------------------------------------------------------------------ #
    # 7. Overview – 12 months, formula roll-ups from the detail tables
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Overview")
    ov_headers = [
        "Month", "Month_Label", "Quarter", "Website_Visits", "New_Contacts",
        "Leads", "MQLs", "SQLs", "Target_Leads", "Lead_Achievement",
        "Budget", "Spend", "Cost_per_Lead", "Email_Sent", "Social_Engagements",
        "Cumulative_Leads", "Cumulative_Target",
    ]
    ov_rows = []
    c_month = rng(camp_headers, "Campaigns", "Month", camp_last)
    e_month = rng(ev_headers, "Events", "Month", ev_last)
    em_month = rng(em_headers, "Email", "Month", em_last)
    so_month = rng(so_headers, "Social_Media", "Month", so_last)
    for m in range(1, 13):
        r = m + 1
        cm = f"$A{r}"
        sea = SEASON[m - 1]

        def both(metric):
            return (f'=SUMIFS({rng(camp_headers, "Campaigns", metric, camp_last)},{c_month},{cm})'
                    f'+SUMIFS({rng(ev_headers, "Events", metric, ev_last)},{e_month},{cm})')

        visits = int(9500 * sea * random.uniform(0.85, 1.15))
        new_contacts = int(95 * sea * random.uniform(0.7, 1.3))
        ov_rows.append([
            f"=DATE(Inp_Year,{m},1)",
            f'=TEXT(A{r},"mmm-yyyy")',
            f'="Q"&ROUNDUP(MONTH(A{r})/3,0)',
            f"=IF(A{r}<=Inp_AsOf,{visits},0)",
            f"=IF(A{r}<=Inp_AsOf,{new_contacts},0)",
            both("Leads"), both("MQLs"), both("SQLs"),
            (f"=ROUND(Inp_TargetLeads/12,0)" if m < 12 else f"=Inp_TargetLeads-SUM(I2:I{r - 1})"),
            safe_div(f"F{r}", f"I{r}"),
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Budget", camp_last)},{c_month},{cm})'
            f'+SUMIFS({rng(ev_headers, "Events", "Budget", ev_last)},{e_month},{cm})',
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Spend", camp_last)},{c_month},{cm})'
            f'+SUMIFS({rng(ev_headers, "Events", "Actual_Cost", ev_last)},{e_month},{cm})',
            safe_div(f"L{r}", f"F{r}"),
            f'=SUMIFS({rng(em_headers, "Email", "Sent", em_last)},{em_month},{cm})',
            f'=SUMIFS({rng(so_headers, "Social_Media", "Total_Engagements", so_last)},{so_month},{cm})',
            f"=SUM($F$2:F{r})", f"=SUM($I$2:I{r})",
        ])
    ov_last = add_table(
        ws, "tbl_Overview", ov_headers, ov_rows,
        {"Month": FMT_MONTH, "Website_Visits": FMT_INT, "New_Contacts": FMT_INT, "Leads": FMT_INT,
         "MQLs": FMT_INT, "SQLs": FMT_INT, "Target_Leads": FMT_INT, "Lead_Achievement": FMT_PCT,
         "Budget": FMT_MONEY, "Spend": FMT_MONEY, "Cost_per_Lead": FMT_MONEY, "Email_Sent": FMT_INT,
         "Social_Engagements": FMT_INT, "Cumulative_Leads": FMT_INT, "Cumulative_Target": FMT_INT},
    )
    ws.cell(row=ov_last + 2, column=1,
            value="Website_Visits and New_Contacts are typed dummy values (editable); everything else rolls up by formula.").font = NOTE_FONT

    # ------------------------------------------------------------------ #
    # 8. Targets / KPI scorecard  (targets from Inputs §1–2)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Targets")
    tg_headers = ["KPI", "Area", "Target", "Actual_YTD", "Achievement", "Gap", "Status", "Unit"]
    L = lambda hdrs, sh, name, last: rng(hdrs, sh, name, last)  # noqa: E731
    kpis = [
        ("Total Leads", "Funnel", "=Inp_TargetLeads", f'=SUM({L(ov_headers,"Overview","Leads",ov_last)})', "leads"),
        ("MQLs", "Funnel", "=Inp_Tgt_MQLs", f'=SUM({L(ov_headers,"Overview","MQLs",ov_last)})', "MQLs"),
        ("SQLs", "Funnel", "=Inp_Tgt_SQLs", f'=SUM({L(ov_headers,"Overview","SQLs",ov_last)})', "SQLs"),
        ("Saudi Database Contacts", "Database", "=Inp_Tgt_Contacts",
         f'=SUM({L(db_headers,"Database","Contacts",db_last)})', "contacts"),
        ("Email Open Rate", "Email", "=Inp_Tgt_EmailOpen",
         safe_div(f'SUM({L(em_headers,"Email","Opens",em_last)})', f'SUM({L(em_headers,"Email","Delivered",em_last)})'), "%"),
        ("Email Click Rate", "Email", "=Inp_Tgt_EmailClick",
         safe_div(f'SUM({L(em_headers,"Email","Clicks",em_last)})', f'SUM({L(em_headers,"Email","Delivered",em_last)})'), "%"),
        ("Social Engagements", "Social", "=Inp_Tgt_SocialEng",
         f'=SUM({L(so_headers,"Social_Media","Total_Engagements",so_last)})', "engagements"),
        ("LinkedIn Followers", "Social", "=Inp_Tgt_LinkedIn",
         f'=_xlfn.MAXIFS({L(so_headers,"Social_Media","Followers",so_last)},'
         f'{L(so_headers,"Social_Media","Platform",so_last)},"LinkedIn")', "followers"),
        ("Events Delivered", "Events", "=Inp_Tgt_Events",
         f'=COUNTIF({L(ev_headers,"Events","Status",ev_last)},"Delivered")', "events"),
        ("Budget Utilisation", "Budget", "=Inp_Tgt_BudgetUtil",
         safe_div(f'SUM({L(bu_headers,"Budget","Spend_YTD",bu_last)})', f'SUM({L(bu_headers,"Budget","Annual_Budget",bu_last)})'), "%"),
    ]
    tg_rows = []
    for n, (kpi, area, tgt, actual, unit) in enumerate(kpis):
        r = n + 2
        tg_rows.append([
            kpi, area, tgt, actual,
            safe_div(f"D{r}", f"C{r}"),
            f"=D{r}-C{r}",
            f'=IF(E{r}>=1,"Achieved",IF(E{r}>=0.8,"On track","At risk"))',
            unit,
        ])
    tg_last = add_table(ws, "tbl_Targets", tg_headers, tg_rows,
                        {"Achievement": FMT_PCT}, {"KPI": 28, "Area": 12, "Status": 12})
    for n, (_, _, _, _, unit) in enumerate(kpis):
        r = n + 2
        fmt = FMT_PCT if unit == "%" else FMT_INT
        for cc in "CDF":
            ws[f"{cc}{r}"].number_format = fmt
    status_cf(ws, f"G2:G{tg_last}", {"Achieved": FILL_GREEN, "On track": FILL_AMBER, "At risk": FILL_RED})
    ws.conditional_formatting.add(f"E2:E{tg_last}", DataBarRule(start_type="num", start_value=0,
                                                                 end_type="num", end_value=1, color=BLUE))
    ws.cell(row=tg_last + 2, column=1, value="Targets come from the Inputs sheet (sections 1–2).").font = NOTE_FONT

    # ------------------------------------------------------------------ #
    # 9. Activities – status derived from dates vs. data-as-of
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Activities")
    ac_headers = ["Activity_ID", "Activity", "Campaign", "Channel", "Owner", "Start_Date", "End_Date",
                  "Status", "Pct_Complete", "Deliverable", "Notes"]
    owners = ["Bandar", "Sara", "Faisal", "Noura", "Agency"]
    activities = [
        ("HVAC Distributed Pumping landing page (AR/EN)", "HVAC", "Web", (1, 5), (2, 10), "Landing page live"),
        ("HVAC LinkedIn lead-gen campaign Q1", "HVAC", "LinkedIn", (1, 15), (3, 31), "Campaign report"),
        ("Energy Efficiency Challenge interactive tool", "HVAC", "Web", (8, 1), (12, 10), "Web tool + lead form"),
        ("Data Centre cooling whitepaper", "Data Centre", "Content", (2, 1), (4, 15), "Whitepaper PDF"),
        ("Data Centre trade media programme", "Data Centre", "Trade Media", (5, 1), (12, 31), "Monthly articles"),
        ("Wastewater lifecycle cost calculator", "Wastewater", "Web", (3, 1), (6, 30), "Calculator tool"),
        ("Wastewater utilities nurture email series", "Wastewater", "Email", (7, 1), (12, 15), "6-email series"),
        ("Hospitality case study – Red Sea resort", "Hospitality", "Content", (4, 1), (5, 20), "Case study"),
        ("Hospitality Saudi Build exhibition stand", "Hospitality", "Events", (8, 15), (10, 22), "Stand + leads"),
        ("Healthcare hygienic water webinar", "Healthcare", "Webinar", (5, 1), (6, 9), "Webinar recording"),
        ("Healthcare Global Health Exhibition stand", "Healthcare", "Events", (9, 1), (12, 5), "Stand + leads"),
        ("WhatsApp Channel launch", "All", "WhatsApp Channel", (1, 10), (2, 1), "Channel live"),
        ("Social media content calendar H2", "All", "Social", (6, 15), (12, 31), "Monthly calendar"),
        ("Saudi database cleansing & enrichment", "All", "Database", (2, 1), (4, 30), "Verified contacts"),
        ("Marketing automation scoring model (MQL/SQL)", "All", "Email", (3, 1), (5, 31), "Scoring rules live"),
        ("Power BI marketing dashboard", "All", "Reporting", (10, 1), (12, 15), "Dashboard v1"),
        ("2027 marketing plan & budget", "All", "Planning", (11, 1), (12, 20), "Plan deck"),
        ("Year-end consultant appreciation evening", "HVAC", "Events", (11, 15), (12, 8), "Event"),
        ("Arabic brand guideline refresh", "All", "Brand", (6, 1), (8, 31), "Guideline PDF"),
        ("Customer satisfaction survey", "All", "Research", (9, 1), (10, 31), "Survey report"),
        ("2027 kick-off webinar for consultants", "All", "Webinar", (12, 10), (12, 20), "Webinar"),
    ]
    ac_rows = []
    for n, (act, camp, ch, (sm, sd), (em_, ed), deliv) in enumerate(activities, 1):
        r = n + 1
        ac_rows.append([
            f"ACT-{n:03d}", act, camp, ch, random.choice(owners),
            f"=DATE(Inp_Year,{sm},{sd})", f"=DATE(Inp_Year,{em_},{ed})",
            f'=IF(G{r}<=Inp_AsOf,"Delivered",IF(F{r}<=Inp_AsOf,"In Progress","Planned"))',
            f'=IF(H{r}="Delivered",1,IF(H{r}="Planned",0,{random.choice([0.35, 0.5, 0.65, 0.8])}))',
            deliv, "",
        ])
    ac_last = add_table(ws, "tbl_Activities", ac_headers, ac_rows,
                        {"Start_Date": FMT_DATE, "End_Date": FMT_DATE, "Pct_Complete": "0%"},
                        {"Activity": 46, "Channel": 18, "Status": 13, "Deliverable": 24, "Notes": 20})
    status_cf(ws, f"H2:H{ac_last}", {"Delivered": FILL_GREEN, "In Progress": FILL_AMBER, "Planned": FILL_GREY})
    ws.conditional_formatting.add(f"I2:I{ac_last}", DataBarRule(start_type="num", start_value=0,
                                                                 end_type="num", end_value=1, color=BLUE))
    ws.cell(row=ac_last + 2, column=1,
            value="Status and % complete follow the data-as-of date on Inputs. Overwrite a cell to set a status manually.").font = NOTE_FONT

    # ------------------------------------------------------------------ #
    # 10. Executive Action Center
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Action_Center")
    xa_headers = ["Action_ID", "Priority", "Action", "Linked_KPI", "Campaign", "Owner", "Due_Date",
                  "Status", "Expected_Impact", "Decision_Needed"]
    actions = [
        ("High", "Reallocate 15% of Hospitality budget to HVAC LinkedIn lead-gen (lowest CPL)", "Cost per Lead", "HVAC", "Bandar", (11, 20), "Open", "+40 leads in Q4", "Yes – budget approval"),
        ("High", "Fast-track SQL handover with sales for Data Centre MQLs older than 14 days", "MQL to SQL", "Data Centre", "Faisal", (11, 18), "In Progress", "+8% MQL→SQL", "No"),
        ("High", "Launch Energy Efficiency Challenge tool before Dec consultant event", "Total Leads", "HVAC", "Agency", (12, 1), "In Progress", "+60 leads", "No"),
        ("Medium", "Re-engage 'Engaged_90d' Saudi contacts with Wastewater nurture series", "Email Click Rate", "Wastewater", "Sara", (11, 30), "Open", "+1.5 pt click rate", "No"),
        ("Medium", "Grow WhatsApp Channel to 2,000 subscribers via event QR codes", "Social Engagements", "All", "Noura", (12, 31), "In Progress", "+600 subscribers", "No"),
        ("Medium", "Approve next-year plan: Target Leads 1,200, budget +8%", "Total Leads", "All", "Bandar", (12, 15), "Open", "Next-year baseline", "Yes – leadership sign-off"),
        ("Low", "Archive underperforming Facebook paid ads; keep organic only", "Engagement Rate", "All", "Sara", (12, 10), "Open", "Save ~SAR 12k", "No"),
        ("Low", "Refresh Healthcare case study with Global Health Exhibition outcomes", "Leads", "Healthcare", "Noura", (12, 20), "Planned", "Sales enablement", "No"),
    ]
    xa_rows = [[f"ACT-EX-{n:02d}", p, a, k, c, o, f"=DATE(Inp_Year,{dm},{dd})", s, imp, dec]
               for n, (p, a, k, c, o, (dm, dd), s, imp, dec) in enumerate(actions, 1)]
    xa_last = add_table(ws, "tbl_Action_Center", xa_headers, xa_rows, {"Due_Date": FMT_DATE},
                        {"Action": 70, "Linked_KPI": 18, "Expected_Impact": 20, "Decision_Needed": 26, "Priority": 10})
    status_cf(ws, f"B2:B{xa_last}", {"High": FILL_RED, "Medium": FILL_AMBER, "Low": FILL_GREY})
    status_cf(ws, f"H2:H{xa_last}", {"Done": FILL_GREEN, "In Progress": FILL_AMBER, "Open": FILL_GREY, "Planned": FILL_GREY})
    dv = DataValidation(type="list", formula1='"Open,In Progress,Done,Planned"', showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add(f"H2:H{xa_last}")
    dv2 = DataValidation(type="list", formula1='"High,Medium,Low"', showErrorMessage=True)
    ws.add_data_validation(dv2)
    dv2.add(f"B2:B{xa_last}")
    for cc in "BH":
        for rr in range(2, xa_last + 1):
            ws[f"{cc}{rr}"].fill, ws[f"{cc}{rr}"].border = INPUT_FILL, INPUT_BOX

    # ------------------------------------------------------------------ #
    # 11. Management Dashboard (Excel view – formulas only)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    for c_ in "BCDEFGH":
        ws.column_dimensions[c_].width = 26
    ws["B2"] = "Grundfos Saudi Arabia – Marketing Dashboard"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = ('="FY"&Inp_Year&"  ·  Data as of "&TEXT(Inp_AsOf,"dd-mmm-yyyy")&"  ·  Scenario: "'
                '&Inp_ScenarioName&"  ·  Dummy data – all figures roll up from the data tables"')
    ws["B3"].font = SUBTITLE_FONT
    link = ws["G2"]
    link.value = "▶  Change inputs (Inputs sheet)"
    link.hyperlink = Hyperlink(ref="G2", location="Inputs!C6", display=link.value)
    link.font = Font(name="Calibri", bold=True, color="FFFFFF")
    link.fill = PatternFill("solid", fgColor=BLUE)
    link.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("G2:H2")
    ws.row_dimensions[2].height = 24

    def kpi_card(cell_label, label, formula, fmt):
        ws[cell_label] = label
        ws[cell_label].font = LABEL_FONT
        ws[cell_label].fill = PatternFill("solid", fgColor=BLUE_LIGHT)
        v = ws.cell(row=ws[cell_label].row + 1, column=ws[cell_label].column, value=formula)
        v.font = KPI_FONT
        v.number_format = fmt
        v.fill = PatternFill("solid", fgColor=BLUE_LIGHT)
        v.alignment = Alignment(horizontal="left")

    ov = lambda name: rng(ov_headers, "Overview", name, ov_last)  # noqa: E731
    kpi_card("B5", "Total Leads YTD", f"=SUM({ov('Leads')})", FMT_INT)
    kpi_card("C5", "Target Leads", "=Inp_TargetLeads", FMT_INT)
    kpi_card("D5", "Lead Achievement", safe_div(f"SUM({ov('Leads')})", "Inp_TargetLeads"), FMT_PCT)
    kpi_card("E5", "MQLs", f"=SUM({ov('MQLs')})", FMT_INT)
    kpi_card("F5", "SQLs", f"=SUM({ov('SQLs')})", FMT_INT)
    kpi_card("G5", "Saudi Database", f"=SUM({rng(db_headers,'Database','Contacts',db_last)})", FMT_INT)

    kpi_card("B8", "Budget (annual)", f"=SUM({rng(bu_headers,'Budget','Annual_Budget',bu_last)})", FMT_MONEY)
    kpi_card("C8", "Spend YTD", f"=SUM({rng(bu_headers,'Budget','Spend_YTD',bu_last)})", FMT_MONEY)
    kpi_card("D8", "Budget Used", safe_div("C9", "B9"), FMT_PCT)
    kpi_card("E8", "Blended Cost per Lead", safe_div("C9", "B6"), FMT_MONEY)
    kpi_card("F8", "Email Open Rate",
             safe_div(f"SUM({rng(em_headers,'Email','Opens',em_last)})",
                      f"SUM({rng(em_headers,'Email','Delivered',em_last)})"), FMT_PCT)
    kpi_card("G8", "Email Click Rate",
             safe_div(f"SUM({rng(em_headers,'Email','Clicks',em_last)})",
                      f"SUM({rng(em_headers,'Email','Delivered',em_last)})"), FMT_PCT)

    kpi_card("B11", "Social Engagements", f"=SUM({rng(so_headers,'Social_Media','Total_Engagements',so_last)})", FMT_INT)
    kpi_card("C11", "Social Posts", f"=SUM({rng(so_headers,'Social_Media','Posts',so_last)})", FMT_INT)
    kpi_card("D11", "Events Delivered", f'=COUNTIF({rng(ev_headers,"Events","Status",ev_last)},"Delivered")', FMT_INT)
    kpi_card("E11", "Event Leads", f"=SUM({rng(ev_headers,'Events','Leads',ev_last)})", FMT_INT)
    kpi_card("F11", "Activities Delivered", '=COUNTIF(Activities!$H:$H,"Delivered")', FMT_INT)
    kpi_card("G11", "Activities In Progress", '=COUNTIF(Activities!$H:$H,"In Progress")', FMT_INT)

    # Campaign summary block
    ws["B14"] = "Campaign performance (YTD)"
    ws["B14"].font = LABEL_FONT
    hdr = ["Campaign", "Leads", "MQLs", "SQLs", "Spend", "Cost per Lead", "Share of Leads"]
    for i, h in enumerate(hdr):
        c = ws.cell(row=15, column=2 + i, value=h)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BOX
    cc = lambda name: rng(camp_headers, "Campaigns", name, camp_last)  # noqa: E731
    for n, camp in enumerate(CAMPAIGNS):
        r = 16 + n
        ws.cell(row=r, column=2, value=camp).border = BOX
        vals = [
            (f'=SUMIFS({cc("Leads")},{cc("Campaign")},$B{r})', FMT_INT),
            (f'=SUMIFS({cc("MQLs")},{cc("Campaign")},$B{r})', FMT_INT),
            (f'=SUMIFS({cc("SQLs")},{cc("Campaign")},$B{r})', FMT_INT),
            (f'=SUMIFS({cc("Spend")},{cc("Campaign")},$B{r})', FMT_MONEY),
            (safe_div(f"F{r}", f"C{r}"), FMT_MONEY),
            (safe_div(f"C{r}", f"SUM($C$16:$C$20)"), FMT_PCT),
        ]
        for i, (f, fmt) in enumerate(vals):
            c = ws.cell(row=r, column=3 + i, value=f)
            c.number_format, c.border = fmt, BOX

    # Monthly trend block
    ws["B23"] = "Monthly trend"
    ws["B23"].font = LABEL_FONT
    hdr = ["Month", "Leads", "Target", "Achievement", "Spend", "Email Sent", "Social Engagements"]
    for i, h in enumerate(hdr):
        c = ws.cell(row=24, column=2 + i, value=h)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BOX
    src_cols = {"Leads": "F", "Target": "I", "Achievement": "J", "Spend": "L",
                "Email Sent": "N", "Social Engagements": "O"}
    for i in range(12):
        r = 25 + i
        src = i + 2
        cells = [(f"=Overview!B{src}", None)] + [
            (f"=Overview!{src_cols[h]}{src}", {"Achievement": FMT_PCT, "Spend": FMT_MONEY}.get(h, FMT_INT))
            for h in hdr[1:]
        ]
        for j, (f, fmt) in enumerate(cells):
            c = ws.cell(row=r, column=2 + j, value=f)
            c.border = BOX
            if fmt:
                c.number_format = fmt
    ws.conditional_formatting.add("E25:E36", DataBarRule(start_type="num", start_value=0,
                                                         end_type="num", end_value=1.5, color=BLUE))

    # ------------------------------------------------------------------ #
    # 12. README – Power BI instructions
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("README")
    ws.column_dimensions["A"].width = 112
    lines = [
        ("Grundfos Marketing Dashboard – Master workbook (dummy data)", TITLE_FONT),
        ("Generated by build_marketing_dashboard_v2.py · Currency SAR · All data is fictional.", SUBTITLE_FONT),
        ("", None),
        ("How to change the data", LABEL_FONT),
        ("1. Open the Inputs sheet (or click the blue button on the Dashboard).", None),
        ("2. Use the ▲▼ spin buttons, the 'Data as of' drop-down, the scenario buttons and the campaign / platform tick-boxes,", None),
        ("   or simply type a new value into any light-blue cell. Every table recalculates immediately.", None),
        ("3. You can also overwrite any single number inside a data table (e.g. the real leads for a month) – it stays an input.", None),
        ("4. Save the file, then in Power BI Desktop press Home → Refresh.", None),
        ("", None),
        ("How to connect in Power BI", LABEL_FONT),
        ("1. Power BI Desktop → Get data → Excel workbook → select this file.", None),
        ("2. In the Navigator tick the TABLES (tbl_*) – not the sheets – then Load / Transform.", None),
        ("3. Mark a Date table on Month; join the fact tables on Month and Campaign.", None),
        ("4. Formula columns (rates, CPL) are pre-calculated and stored as values inside the file, so Power BI reads them directly.", None),
        ("   Tables starting with tbl_Inputs_ are assumption tables – you do not need them in Power BI.", None),
        ("", None),
        ("Tables", LABEL_FONT),
        ("tbl_Overview       12-month roll-up: leads, MQLs, SQLs, target, budget, spend, cumulative (formula-driven from detail tables)", None),
        ("tbl_Campaigns      HVAC, Data Centre, Wastewater, Hospitality, Healthcare × 12 months", None),
        ("tbl_Email          Sent, Delivered, Opens, Clicks + Delivery/Open/Click/CTR formulas", None),
        ("tbl_Social_Media   LinkedIn, Facebook, WhatsApp Channel: posts, likes, comments, reposts, clicks", None),
        ("tbl_Events         Events with attendees, leads, MQLs, SQLs, budget vs cost", None),
        ("tbl_Database       Saudi database by region and segment – total = Inputs › Saudi database contacts", None),
        ("tbl_Budget         Budget lines, spend YTD, remaining, % used", None),
        ("tbl_Targets        KPI scorecard (targets from the Inputs sheet)", None),
        ("tbl_Activities     Marketing activities – Delivered / In Progress / Planned (status follows the data-as-of date)", None),
        ("tbl_Action_Center  Executive Action Center – priorities, owners, decisions needed", None),
        ("tbl_Inputs_*       Assumption tables used by the formulas (campaigns, social, event types, seasonality, database mix)", None),
        ("Dashboard          Excel management view (formulas only – not needed in Power BI)", None),
        ("", None),
        ("Definitions", LABEL_FONT),
        ("Open Rate = Opens / Delivered · Click Rate = Clicks / Delivered · CTR (click-to-open) = Clicks / Opens", None),
        ("Engagement Rate = (Likes + Comments + Reposts + Clicks) / Impressions · Cost per Lead = Spend / Leads", None),
        ("Data as of = last closed month: months after it show 0 actuals; events / activities after it are 'Planned'.", None),
    ]
    for i, (text, font) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font

    # Sheet order
    order = ["Dashboard", "Inputs", "README", "Overview", "Targets", "Campaigns", "Events", "Email",
             "Social_Media", "Database", "Budget", "Activities", "Action_Center"]
    wb._sheets = [wb[n] for n in order]
    for n in order:
        wb[n].sheet_properties.tabColor = BLUE_DARK if n in ("Dashboard", "README") else (
            "F2A900" if n == "Inputs" else BLUE)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------- #
# Excel COM: add form controls + pre-calculate
# --------------------------------------------------------------------------- #
XL_CHECKBOX, XL_DROPDOWN, XL_OPTION, XL_SPINNER = 1, 2, 7, 9


def add_form_controls(ws):
    """Create the native Excel buttons described in CONTROLS on the Inputs sheet."""
    for shp in list(ws.Shapes):
        if shp.Name.startswith("ctl_"):
            shp.Delete()
    for i, (kind, anchor, opt) in enumerate(CONTROLS):
        cell = ws.Range(anchor)
        left, top, width, height = cell.Left, cell.Top, cell.Width, cell.Height
        linked = opt.get("linked")
        keep = ws.Range(linked).Value if linked else None   # linking resets the cell → restore it
        if kind == "spinner":
            s = ws.Shapes.AddFormControl(XL_SPINNER, left + 2, top + 1, 22, height - 2)
            s.ControlFormat.Min, s.ControlFormat.Max = opt["min"], opt["max"]
            s.ControlFormat.SmallChange = opt["step"]
            s.ControlFormat.LinkedCell = linked
            s.ControlFormat.Value = int(keep)
        elif kind == "dropdown":
            s = ws.Shapes.AddFormControl(XL_DROPDOWN, left + 2, top + 1, width - 4, height - 2)
            s.ControlFormat.ListFillRange = opt["list"]
            s.ControlFormat.LinkedCell = linked
            s.ControlFormat.DropDownLines = opt["lines"]
            s.ControlFormat.ListIndex = int(keep)
        elif kind == "options":
            for j, cap in enumerate(opt["captions"]):
                s = ws.Shapes.AddFormControl(XL_OPTION, left + 2 + j * 96, top, 94, height)
                s.TextFrame.Characters().Text = cap
                s.ControlFormat.LinkedCell = linked
                if j + 1 == opt["default"]:
                    s.ControlFormat.Value = 1
                s.Name = f"ctl_{i}_{j}"
            ws.Range(linked).Value = opt["default"]
            continue
        elif kind == "checkbox":
            s = ws.Shapes.AddFormControl(XL_CHECKBOX, left + width / 2 - 10, top, 20, height)
            s.TextFrame.Characters().Text = ""
            s.ControlFormat.LinkedCell = linked
            s.ControlFormat.Value = 1 if bool(keep) else 0
        s.Name = f"ctl_{i}"


def finalise_with_excel(path: Path) -> bool:
    """Open in Excel: add form controls, full recalculation, save (Power BI reads cached values)."""
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(path.resolve()))
        add_form_controls(wb.Worksheets("Inputs"))
        excel.CalculateFullRebuild()
        wb.Worksheets("Dashboard").Activate()
        wb.Save()
        wb.Close(SaveChanges=True)
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"Excel step skipped: {exc}")
        return False
    finally:
        if excel is not None:
            excel.Quit()


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx")
    build(out)
    print(f"Workbook written: {out.resolve()}")
    if finalise_with_excel(out):
        print("Form controls added and formulas pre-calculated with Excel.")
    else:
        print("Note: open the file once in Excel and save, so Power BI can read formula results "
              "(inputs are typed; no buttons without Excel).")
