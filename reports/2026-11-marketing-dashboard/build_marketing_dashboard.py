"""
Build Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx

A Power BI-ready marketing dashboard workbook populated with dummy data.
Every data sheet is an Excel Table (ListObject) so Power BI can import it
by name. Rates, totals and KPI roll-ups are real Excel formulas.

Usage:
    python build_marketing_dashboard.py [output_path]

Requires: openpyxl  (optional: pywin32 + Excel, to pre-calculate formulas)
"""

import random
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

random.seed(2026)

YEAR = 2026
MONTHS = [date(YEAR, m, 1) for m in range(1, 13)]
MONTH_LABELS = [d.strftime("%b-%Y") for d in MONTHS]
CAMPAIGNS = ["HVAC", "Data Centre", "Wastewater", "Hospitality", "Healthcare"]
PLATFORMS = ["LinkedIn", "Facebook", "WhatsApp Channel"]
TARGET_LEADS = 1000
SAUDI_CONTACTS = 1950
CURRENCY = "SAR"

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
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_INT = "#,##0"
FMT_MONEY = '#,##0 "SAR"'
FMT_PCT = "0.0%"
FMT_DATE = "dd-mmm-yyyy"
FMT_MONTH = "mmm-yyyy"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def add_table(ws, name, headers, rows, number_formats=None, widths=None, start_row=1):
    """Write headers + rows, wrap them in an Excel Table and apply formats."""
    number_formats = number_formats or {}
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, row in enumerate(rows, start_row + 1):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            fmt = number_formats.get(headers[c - 1])
            if fmt:
                cell.number_format = fmt
    last_row = start_row + len(rows)
    ref = f"A{start_row}:{get_column_letter(len(headers))}{last_row}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    for i, h in enumerate(headers, 1):
        w = (widths or {}).get(h, max(12, min(32, len(h) + 4)))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[start_row].height = 30
    return last_row


def col(headers, name):
    return get_column_letter(headers.index(name) + 1)


def rng(headers, sheet, name, last_row, first_row=2):
    c = col(headers, name)
    return f"{sheet}!${c}${first_row}:${c}${last_row}"


def seasonal(base, i, spread=0.25):
    """Dummy seasonality: dip in summer (Jul/Aug) and Ramadan-ish lull, lift in Q4."""
    curve = [0.85, 0.9, 1.0, 1.05, 1.0, 0.95, 0.7, 0.75, 1.05, 1.15, 1.2, 1.1]
    return max(0, int(base * curve[i] * random.uniform(1 - spread, 1 + spread)))


def safe_div(num, den):
    return f'=IFERROR({num}/{den},0)'


# --------------------------------------------------------------------------- #
# Build
# --------------------------------------------------------------------------- #
def build(path: Path):
    wb = Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ #
    # 1. Campaigns  (HVAC, Data Centre, Wastewater, Hospitality, Healthcare)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Campaigns")
    camp_headers = [
        "Month", "Month_Label", "Campaign", "Segment", "Primary_Channel",
        "Impressions", "Clicks", "Leads", "MQLs", "SQLs", "Budget", "Spend",
        "CTR", "Cost_per_Lead", "Lead_to_MQL", "MQL_to_SQL",
    ]
    camp_profile = {
        "HVAC":        dict(seg="Commercial Buildings", ch="LinkedIn Ads",  imp=42000, leads=16, budget=32000),
        "Data Centre": dict(seg="Data Centres",         ch="Trade Media",   imp=28000, leads=10, budget=38000),
        "Wastewater":  dict(seg="Water Utilities",      ch="Events",        imp=22000, leads=9,  budget=26000),
        "Hospitality": dict(seg="Hotels & Resorts",     ch="Email",         imp=30000, leads=11, budget=22000),
        "Healthcare":  dict(seg="Hospitals",            ch="LinkedIn Ads",  imp=24000, leads=8,  budget=20000),
    }
    camp_rows = []
    for i, m in enumerate(MONTHS):
        for camp in CAMPAIGNS:
            p = camp_profile[camp]
            imp = seasonal(p["imp"], i)
            clicks = int(imp * random.uniform(0.012, 0.03))
            leads = seasonal(p["leads"], i, 0.35)
            mqls = int(leads * random.uniform(0.45, 0.7))
            sqls = int(mqls * random.uniform(0.35, 0.6))
            budget = round(p["budget"] / 12, -2)
            spend = round(budget * random.uniform(0.75, 1.08), -1)
            r = len(camp_rows) + 2
            camp_rows.append([
                m, MONTH_LABELS[i], camp, p["seg"], p["ch"], imp, clicks, leads, mqls, sqls,
                budget, spend,
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
    # 2. Events  (Leads, MQLs, SQLs)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Events")
    ev_headers = [
        "Event_ID", "Event_Name", "Event_Date", "Month", "City", "Event_Type", "Campaign",
        "Attendees", "Leads", "MQLs", "SQLs", "Budget", "Actual_Cost",
        "Cost_per_Lead", "Lead_Conversion", "Status",
    ]
    events = [
        ("Saudi HVAC-R Expo",                 date(YEAR, 1, 27), "Riyadh",  "Exhibition",  "HVAC"),
        ("Data Centre World Saudi",           date(YEAR, 2, 17), "Riyadh",  "Exhibition",  "Data Centre"),
        ("Grundfos Distributed Pumping Day",  date(YEAR, 3, 10), "Jeddah",  "Seminar",     "HVAC"),
        ("Saudi Water Forum",                 date(YEAR, 4, 21), "Riyadh",  "Conference",  "Wastewater"),
        ("Hospitality Innovation Summit",     date(YEAR, 5, 12), "Jeddah",  "Conference",  "Hospitality"),
        ("Healthcare Facilities Webinar",     date(YEAR, 6, 9),  "Online",  "Webinar",     "Healthcare"),
        ("Energy Efficiency Roadshow",        date(YEAR, 9, 15), "Dammam",  "Roadshow",    "HVAC"),
        ("Wastewater Lifecycle Cost Webinar", date(YEAR, 9, 29), "Online",  "Webinar",     "Wastewater"),
        ("Saudi Build",                       date(YEAR, 10, 20), "Riyadh", "Exhibition",  "Hospitality"),
        ("Data Centre Cooling Masterclass",   date(YEAR, 11, 10), "Riyadh", "Seminar",     "Data Centre"),
        ("Global Health Exhibition",          date(YEAR, 11, 24), "Riyadh", "Exhibition",  "Healthcare"),
        ("Year-End Consultant Appreciation",  date(YEAR, 12, 8),  "Jeddah", "Networking",  "HVAC"),
    ]
    ev_rows = []
    for n, (name, d, city, etype, camp) in enumerate(events, 1):
        attendees = {"Exhibition": 420, "Conference": 260, "Seminar": 90, "Webinar": 150,
                     "Roadshow": 120, "Networking": 60}[etype]
        attendees = int(attendees * random.uniform(0.8, 1.2))
        leads = int(attendees * random.uniform(0.08, 0.15))
        mqls = int(leads * random.uniform(0.4, 0.65))
        sqls = int(mqls * random.uniform(0.3, 0.55))
        budget = {"Exhibition": 120000, "Conference": 60000, "Seminar": 25000, "Webinar": 8000,
                  "Roadshow": 45000, "Networking": 20000}[etype]
        cost = round(budget * random.uniform(0.85, 1.05), -2)
        r = n + 1
        ev_rows.append([
            f"EV-{n:03d}", name, d, date(d.year, d.month, 1), city, etype, camp,
            attendees, leads, mqls, sqls, budget, cost,
            safe_div(f"M{r}", f"I{r}"),
            safe_div(f"I{r}", f"H{r}"),
            "Delivered" if d <= date(YEAR, 11, 12) else "Planned",
        ])
    ev_last = add_table(
        ws, "tbl_Events", ev_headers, ev_rows,
        {"Event_Date": FMT_DATE, "Month": FMT_MONTH, "Attendees": FMT_INT, "Leads": FMT_INT,
         "MQLs": FMT_INT, "SQLs": FMT_INT, "Budget": FMT_MONEY, "Actual_Cost": FMT_MONEY,
         "Cost_per_Lead": FMT_MONEY, "Lead_Conversion": FMT_PCT},
        {"Event_Name": 34},
    )

    # ------------------------------------------------------------------ #
    # 3. Email performance (formulas: delivery, open rate, click rate, CTR)
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
    for i, m in enumerate(MONTHS):
        for camp in CAMPAIGNS:
            sent = random.randint(350, 620)
            delivered = int(sent * random.uniform(0.95, 0.99))
            opens = int(delivered * random.uniform(0.22, 0.38))
            clicks = int(opens * random.uniform(0.1, 0.22))
            unsubs = random.randint(0, 4)
            r = len(em_rows) + 2
            em_rows.append([
                m, MONTH_LABELS[i], camp, f"{em_topics[camp]} – {MONTH_LABELS[i]}",
                sent, delivered, f"=E{r}-F{r}", opens, clicks, unsubs,
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
    # 4. Social media (LinkedIn, Facebook, WhatsApp Channel)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Social_Media")
    so_headers = [
        "Month", "Month_Label", "Platform", "Followers", "New_Followers", "Posts",
        "Impressions", "Likes", "Comments", "Reposts", "Clicks",
        "Total_Engagements", "Engagement_Rate", "Avg_Engagement_per_Post",
    ]
    so_profile = {
        "LinkedIn":         dict(f=8200, nf=160, posts=14, imp=38000, like=0.028, com=0.004, rep=0.006, clk=0.012),
        "Facebook":         dict(f=5400, nf=70,  posts=10, imp=18000, like=0.022, com=0.003, rep=0.004, clk=0.007),
        "WhatsApp Channel": dict(f=1100, nf=55,  posts=8,  imp=6000,  like=0.06,  com=0.0,   rep=0.02,  clk=0.03),
    }
    so_rows = []
    followers = {p: so_profile[p]["f"] for p in PLATFORMS}
    for i, m in enumerate(MONTHS):
        for p in PLATFORMS:
            pr = so_profile[p]
            new_f = seasonal(pr["nf"], i, 0.3)
            followers[p] += new_f
            posts = max(3, int(pr["posts"] * random.uniform(0.8, 1.2)))
            imp = seasonal(pr["imp"], i)
            likes = int(imp * pr["like"] * random.uniform(0.8, 1.2))
            comments = int(imp * pr["com"] * random.uniform(0.7, 1.3))
            reposts = int(imp * pr["rep"] * random.uniform(0.7, 1.3))
            clicks = int(imp * pr["clk"] * random.uniform(0.8, 1.2))
            r = len(so_rows) + 2
            so_rows.append([
                m, MONTH_LABELS[i], p, followers[p], new_f, posts, imp, likes, comments, reposts, clicks,
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
    # 5. Saudi database  (1,950 contacts)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Database")
    db_headers = ["Country", "Region", "Segment", "Contact_Type", "Contacts",
                  "Opted_In", "Engaged_90d", "Opt_In_Rate", "Engagement_Rate"]
    regions = ["Central (Riyadh)", "Western (Jeddah/Makkah)", "Eastern (Dammam/Khobar)"]
    segments = CAMPAIGNS + ["Industrial", "Consultants & EPCs"]
    weights = {"HVAC": 0.26, "Data Centre": 0.12, "Wastewater": 0.14, "Hospitality": 0.13,
               "Healthcare": 0.11, "Industrial": 0.10, "Consultants & EPCs": 0.14}
    region_w = [0.48, 0.30, 0.22]
    raw = []
    for seg in segments:
        for ri, reg in enumerate(regions):
            raw.append([reg, seg, SAUDI_CONTACTS * weights[seg] * region_w[ri]])
    # round so the grand total is exactly SAUDI_CONTACTS
    ints = [int(x[2]) for x in raw]
    diff = SAUDI_CONTACTS - sum(ints)
    for k in range(diff):
        ints[k % len(ints)] += 1
    db_rows = []
    for n, ((reg, seg, _), cnt) in enumerate(zip(raw, ints)):
        ctype = {"Consultants & EPCs": "Consultant / EPC"}.get(seg, "End User / Facility Owner")
        opted = int(cnt * random.uniform(0.7, 0.9))
        engaged = int(opted * random.uniform(0.3, 0.5))
        r = n + 2
        db_rows.append(["Saudi Arabia", reg, seg, ctype, cnt, opted, engaged,
                        safe_div(f"F{r}", f"E{r}"), safe_div(f"G{r}", f"F{r}")])
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
    bu_rows = []
    r = 2
    for camp in CAMPAIGNS:
        bu_rows.append([
            f"{camp} Campaign", "Campaign",
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Budget", camp_last)},'
            f'{rng(camp_headers, "Campaigns", "Campaign", camp_last)},A{r})',
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Spend", camp_last)},'
            f'{rng(camp_headers, "Campaigns", "Campaign", camp_last)},"{camp}")',
            f"=C{r}-D{r}", safe_div(f"D{r}", f"C{r}"),
            f'=IF(F{r}>1,"Over budget",IF(F{r}>0.9,"Watch","On track"))',
        ])
        bu_rows[-1][2] = (f'=SUMIFS({rng(camp_headers, "Campaigns", "Budget", camp_last)},'
                          f'{rng(camp_headers, "Campaigns", "Campaign", camp_last)},"{camp}")')
        r += 1
    bu_rows.append([
        "Events & Exhibitions", "Events",
        f'=SUM({rng(ev_headers, "Events", "Budget", ev_last)})',
        f'=SUMIFS({rng(ev_headers, "Events", "Actual_Cost", ev_last)},'
        f'{rng(ev_headers, "Events", "Status", ev_last)},"Delivered")',
        f"=C{r}-D{r}", safe_div(f"D{r}", f"C{r}"),
        f'=IF(F{r}>1,"Over budget",IF(F{r}>0.9,"Watch","On track"))',
    ]); r += 1
    for line, cat, bud, spent in [
        ("Social Media Content & Paid", "Digital", 90000, 71500),
        ("Email Platform & Automation", "Digital", 36000, 31200),
        ("Database & Market Intelligence", "Data", 45000, 38900),
        ("Brand & Collateral (AR/EN)", "Brand", 60000, 44800),
        ("Agency & Creative", "Agency", 120000, 101300),
    ]:
        bu_rows.append([line, cat, bud, spent, f"=C{r}-D{r}", safe_div(f"D{r}", f"C{r}"),
                        f'=IF(F{r}>1,"Over budget",IF(F{r}>0.9,"Watch","On track"))'])
        r += 1
    bu_last = add_table(
        ws, "tbl_Budget", bu_headers, bu_rows,
        {"Annual_Budget": FMT_MONEY, "Spend_YTD": FMT_MONEY, "Remaining": FMT_MONEY, "Pct_Used": FMT_PCT},
        {"Budget_Line": 32, "Status": 14},
    )
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
    monthly_target = [TARGET_LEADS // 12] * 12
    monthly_target[-1] += TARGET_LEADS - sum(monthly_target)
    for i, m in enumerate(MONTHS):
        r = i + 2
        cm = f"A{r}"
        c_month = rng(camp_headers, "Campaigns", "Month", camp_last)
        e_month = rng(ev_headers, "Events", "Month", ev_last)
        em_month = rng(em_headers, "Email", "Month", em_last)
        so_month = rng(so_headers, "Social_Media", "Month", so_last)

        def both(metric):
            return (f'=SUMIFS({rng(camp_headers, "Campaigns", metric, camp_last)},{c_month},{cm})'
                    f'+SUMIFS({rng(ev_headers, "Events", metric, ev_last)},{e_month},{cm})')

        ov_rows.append([
            m, MONTH_LABELS[i], f"Q{(m.month - 1) // 3 + 1}",
            seasonal(9500, i, 0.15), seasonal(95, i, 0.3),
            both("Leads"), both("MQLs"), both("SQLs"),
            monthly_target[i], safe_div(f"F{r}", f"I{r}"),
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Budget", camp_last)},{c_month},{cm})'
            f'+SUMIFS({rng(ev_headers, "Events", "Budget", ev_last)},{e_month},{cm})',
            f'=SUMIFS({rng(camp_headers, "Campaigns", "Spend", camp_last)},{c_month},{cm})'
            f'+SUMIFS({rng(ev_headers, "Events", "Actual_Cost", ev_last)},{e_month},{cm},'
            f'{rng(ev_headers, "Events", "Status", ev_last)},"Delivered")',
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

    # ------------------------------------------------------------------ #
    # 8. Targets / KPI scorecard
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Targets")
    tg_headers = ["KPI", "Area", "Target", "Actual_YTD", "Achievement", "Gap", "Status", "Unit"]
    L = lambda hdrs, sh, name, last: rng(hdrs, sh, name, last)  # noqa: E731
    kpis = [
        ("Total Leads", "Funnel", TARGET_LEADS, f'=SUM({L(ov_headers,"Overview","Leads",ov_last)})', "leads"),
        ("MQLs", "Funnel", 550, f'=SUM({L(ov_headers,"Overview","MQLs",ov_last)})', "MQLs"),
        ("SQLs", "Funnel", 250, f'=SUM({L(ov_headers,"Overview","SQLs",ov_last)})', "SQLs"),
        ("Saudi Database Contacts", "Database", 2000, f'=SUM({L(db_headers,"Database","Contacts",db_last)})', "contacts"),
        ("Email Open Rate", "Email", 0.30,
         safe_div(f'SUM({L(em_headers,"Email","Opens",em_last)})', f'SUM({L(em_headers,"Email","Delivered",em_last)})'), "%"),
        ("Email Click Rate", "Email", 0.05,
         safe_div(f'SUM({L(em_headers,"Email","Clicks",em_last)})', f'SUM({L(em_headers,"Email","Delivered",em_last)})'), "%"),
        ("Social Engagements", "Social", 45000, f'=SUM({L(so_headers,"Social_Media","Total_Engagements",so_last)})', "engagements"),
        ("LinkedIn Followers", "Social", 10000,
         f'=MAX(IF({L(so_headers,"Social_Media","Platform",so_last)}="LinkedIn",{L(so_headers,"Social_Media","Followers",so_last)}))', "followers"),
        ("Events Delivered", "Events", 12, f'=COUNTIF({L(ev_headers,"Events","Status",ev_last)},"Delivered")', "events"),
        ("Budget Utilisation", "Budget", 0.95,
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
    for n, (_, _, tgt, _, unit) in enumerate(kpis):
        r = n + 2
        fmt = FMT_PCT if unit == "%" else FMT_INT
        for cc in "CDF":
            ws[f"{cc}{r}"].number_format = fmt
    # array formula for LinkedIn followers (works in Excel 365 as dynamic; wrap for legacy)
    from openpyxl.worksheet.formula import ArrayFormula
    r_li = 2 + [k[0] for k in kpis].index("LinkedIn Followers")
    ws[f"D{r_li}"] = ArrayFormula(f"D{r_li}", ws[f"D{r_li}"].value)

    # ------------------------------------------------------------------ #
    # 9. Activities – Delivered / In Progress / Planned
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Activities")
    ac_headers = ["Activity_ID", "Activity", "Campaign", "Channel", "Owner", "Start_Date", "End_Date",
                  "Status", "Pct_Complete", "Deliverable", "Notes"]
    owners = ["Bandar", "Sara", "Faisal", "Noura", "Agency"]
    activities = [
        ("HVAC Distributed Pumping landing page (AR/EN)", "HVAC", "Web", (1, 5), (2, 10), "Delivered", "Landing page live"),
        ("HVAC LinkedIn lead-gen campaign Q1", "HVAC", "LinkedIn", (1, 15), (3, 31), "Delivered", "Campaign report"),
        ("Energy Efficiency Challenge interactive tool", "HVAC", "Web", (8, 1), (11, 30), "In Progress", "Web tool + lead form"),
        ("Data Centre cooling whitepaper", "Data Centre", "Content", (2, 1), (4, 15), "Delivered", "Whitepaper PDF"),
        ("Data Centre trade media programme", "Data Centre", "Trade Media", (5, 1), (12, 31), "In Progress", "Monthly articles"),
        ("Wastewater lifecycle cost calculator", "Wastewater", "Web", (3, 1), (6, 30), "Delivered", "Calculator tool"),
        ("Wastewater utilities nurture email series", "Wastewater", "Email", (7, 1), (12, 15), "In Progress", "6-email series"),
        ("Hospitality case study – Red Sea resort", "Hospitality", "Content", (4, 1), (5, 20), "Delivered", "Case study"),
        ("Hospitality Saudi Build exhibition stand", "Hospitality", "Events", (8, 15), (10, 22), "Delivered", "Stand + leads"),
        ("Healthcare hygienic water webinar", "Healthcare", "Webinar", (5, 1), (6, 9), "Delivered", "Webinar recording"),
        ("Healthcare Global Health Exhibition stand", "Healthcare", "Events", (9, 1), (11, 26), "In Progress", "Stand + leads"),
        ("WhatsApp Channel launch", "All", "WhatsApp Channel", (1, 10), (2, 1), "Delivered", "Channel live"),
        ("Social media content calendar H2", "All", "Social", (6, 15), (12, 31), "In Progress", "Monthly calendar"),
        ("Saudi database cleansing & enrichment", "All", "Database", (2, 1), (4, 30), "Delivered", "1,950 verified contacts"),
        ("Marketing automation scoring model (MQL/SQL)", "All", "Email", (3, 1), (5, 31), "Delivered", "Scoring rules live"),
        ("Power BI marketing dashboard", "All", "Reporting", (10, 1), (12, 15), "In Progress", "Dashboard v1"),
        ("2027 marketing plan & budget", "All", "Planning", (11, 1), (12, 20), "In Progress", "Plan deck"),
        ("Consultant appreciation evening", "HVAC", "Events", (11, 15), (12, 8), "Planned", "Event"),
        ("Arabic brand guideline refresh", "All", "Brand", (6, 1), (8, 31), "Delivered", "Guideline PDF"),
        ("Customer satisfaction survey", "All", "Research", (9, 1), (10, 31), "Delivered", "Survey report"),
    ]
    ac_rows = []
    for n, (act, camp, ch, (sm, sd), (em_, ed), status, deliv) in enumerate(activities, 1):
        pct = 1.0 if status == "Delivered" else (0.0 if status == "Planned" else random.choice([0.35, 0.5, 0.65, 0.8]))
        ac_rows.append([f"ACT-{n:03d}", act, camp, ch, random.choice(owners),
                        date(YEAR, sm, sd), date(YEAR, em_, ed), status, pct, deliv, ""])
    add_table(ws, "tbl_Activities", ac_headers, ac_rows,
              {"Start_Date": FMT_DATE, "End_Date": FMT_DATE, "Pct_Complete": "0%"},
              {"Activity": 46, "Channel": 18, "Status": 13, "Deliverable": 24, "Notes": 20})

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
        ("Medium", "Approve 2027 plan: Target Leads 1,200, budget +8%", "Total Leads", "All", "Bandar", (12, 15), "Open", "2027 baseline", "Yes – leadership sign-off"),
        ("Low", "Archive underperforming Facebook paid ads; keep organic only", "Engagement Rate", "All", "Sara", (12, 10), "Open", "Save ~SAR 12k", "No"),
        ("Low", "Refresh Healthcare case study with Global Health Exhibition outcomes", "Leads", "Healthcare", "Noura", (12, 20), "Planned", "Sales enablement", "No"),
    ]
    xa_rows = [[f"ACT-EX-{n:02d}", p, a, k, c, o, date(YEAR, dm, dd), s, imp, dec]
               for n, (p, a, k, c, o, (dm, dd), s, imp, dec) in enumerate(actions, 1)]
    add_table(ws, "tbl_Action_Center", xa_headers, xa_rows, {"Due_Date": FMT_DATE},
              {"Action": 70, "Linked_KPI": 18, "Expected_Impact": 20, "Decision_Needed": 26, "Priority": 10})

    # ------------------------------------------------------------------ #
    # 11. Management Dashboard (Excel view – formulas only)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    for c_ in "BCDEFGH":
        ws.column_dimensions[c_].width = 22
    ws["B2"] = "Grundfos Saudi Arabia – Marketing Dashboard"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = f"FY{YEAR} · Dummy data · All figures roll up from the data tables via formulas"
    ws["B3"].font = SUBTITLE_FONT

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
    kpi_card("C5", "Target Leads", f"={TARGET_LEADS}", FMT_INT)
    kpi_card("D5", "Lead Achievement", safe_div(f"SUM({ov('Leads')})", str(TARGET_LEADS)), FMT_PCT)
    kpi_card("E5", "MQLs", f"=SUM({ov('MQLs')})", FMT_INT)
    kpi_card("F5", "SQLs", f"=SUM({ov('SQLs')})", FMT_INT)
    kpi_card("G5", "Saudi Database", f"=SUM({rng(db_headers,'Database','Contacts',db_last)})", FMT_INT)

    kpi_card("B8", "Budget (annual)", f"=SUM({rng(bu_headers,'Budget','Annual_Budget',bu_last)})", FMT_MONEY)
    kpi_card("C8", "Spend YTD", f"=SUM({rng(bu_headers,'Budget','Spend_YTD',bu_last)})", FMT_MONEY)
    kpi_card("D8", "Budget Used", safe_div("C9", "B9"), FMT_PCT)
    kpi_card("E8", "Cost per Lead", safe_div("C9", "B6"), FMT_MONEY)
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
    for i in range(12):
        r = 25 + i
        src = i + 2
        cells = [
            (f"=Overview!B{src}", None),
            (f"=Overview!F{src}", FMT_INT),
            (f"=Overview!I{src}", FMT_INT),
            (f"=Overview!J{src}", FMT_PCT),
            (f"=Overview!L{src}", FMT_MONEY),
            (f"=Overview!N{src}", FMT_INT),
            (f"=Overview!O{src}", FMT_INT),
        ]
        for j, (f, fmt) in enumerate(cells):
            c = ws.cell(row=r, column=2 + j, value=f)
            c.border = BOX
            if fmt:
                c.number_format = fmt

    # ------------------------------------------------------------------ #
    # 12. README – Power BI instructions
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("README")
    ws.column_dimensions["A"].width = 110
    lines = [
        ("Grundfos Marketing Dashboard – Master workbook (dummy data)", TITLE_FONT),
        (f"Generated by build_marketing_dashboard.py · FY{YEAR} · Currency {CURRENCY} · All data is fictional.", SUBTITLE_FONT),
        ("", None),
        ("How to connect in Power BI", LABEL_FONT),
        ("1. Power BI Desktop → Get data → Excel workbook → select this file.", None),
        ("2. In the Navigator tick the TABLES (tbl_*) – not the sheets – then Load / Transform.", None),
        ("3. Mark tbl_Overview[Month] (or a Date table) as the date key; join the fact tables on Month and Campaign.", None),
        ("4. Formula columns (rates, CPL) are pre-calculated and stored as values inside the file, so Power BI reads them directly.", None),
        ("   If you edit numbers in Excel, just save; Power BI refresh will pick up the recalculated values.", None),
        ("", None),
        ("Tables", LABEL_FONT),
        ("tbl_Overview       12-month roll-up: leads, MQLs, SQLs, target, budget, spend, cumulative (formula-driven from detail tables)", None),
        ("tbl_Campaigns      HVAC, Data Centre, Wastewater, Hospitality, Healthcare × 12 months", None),
        ("tbl_Email          Sent, Delivered, Opens, Clicks + Delivery/Open/Click/CTR formulas", None),
        ("tbl_Social_Media   LinkedIn, Facebook, WhatsApp Channel: posts, likes, comments, reposts, clicks", None),
        ("tbl_Events         Events with attendees, leads, MQLs, SQLs, budget vs cost", None),
        (f"tbl_Database       Saudi database by region and segment – total {SAUDI_CONTACTS:,} contacts", None),
        ("tbl_Budget         Budget lines, spend YTD, remaining, % used", None),
        (f"tbl_Targets        KPI scorecard (Target Leads = {TARGET_LEADS:,})", None),
        ("tbl_Activities     Marketing activities – Delivered / In Progress / Planned", None),
        ("tbl_Action_Center  Executive Action Center – priorities, owners, decisions needed", None),
        ("Dashboard          Excel management view (formulas only – not needed in Power BI)", None),
        ("", None),
        ("Definitions", LABEL_FONT),
        ("Open Rate = Opens / Delivered · Click Rate = Clicks / Delivered · CTR (click-to-open) = Clicks / Opens", None),
        ("Engagement Rate = (Likes + Comments + Reposts + Clicks) / Impressions · Cost per Lead = Spend / Leads", None),
    ]
    for i, (text, font) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font

    # Sheet order
    order = ["Dashboard", "README", "Overview", "Targets", "Campaigns", "Events", "Email",
             "Social_Media", "Database", "Budget", "Activities", "Action_Center"]
    wb._sheets = [wb[n] for n in order]
    for n in order:
        wb[n].sheet_properties.tabColor = BLUE_DARK if n in ("Dashboard", "README") else BLUE

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def recalc_with_excel(path: Path) -> bool:
    """Open/save in Excel so formula results are cached (Power BI reads cached values)."""
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
        excel.CalculateFullRebuild()
        wb.Save()
        wb.Close(SaveChanges=True)
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"Excel recalculation skipped: {exc}")
        return False
    finally:
        if excel is not None:
            excel.Quit()


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx")
    build(out)
    print(f"Workbook written: {out.resolve()}")
    if recalc_with_excel(out):
        print("Formulas pre-calculated with Excel (Power BI will read computed values).")
    else:
        print("Note: open the file once in Excel and save, so Power BI can read formula results.")
