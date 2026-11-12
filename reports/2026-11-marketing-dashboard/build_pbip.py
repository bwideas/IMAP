"""
Build a Power BI project (.pbip) for the Grundfos Marketing Dashboard workbook.

Creates:
  <out>/Grundfos_Marketing_Dashboard.pbip
  <out>/Grundfos_Marketing_Dashboard.SemanticModel/  (model.bim: typed tables, dims, relationships, measures)
  <out>/Grundfos_Marketing_Dashboard.Report/         (report.json: starter pages)
  <out>/Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx (copy of the data)

Usage: python build_pbip.py <xlsx_path> <output_folder>
"""

import json
import shutil
import sys
import uuid
from pathlib import Path

NAME = "Grundfos_Marketing_Dashboard"
BLUE_DARK = "#11497B"
BLUE = "#126AF3"
BLUE_LIGHT = "#E9F1F7"

# kind -> (M type, TMSL dataType, formatString, summarizeBy)
KINDS = {
    "date":  ("type date",   "dateTime", "dd-mmm-yyyy", "none"),
    "month": ("type date",   "dateTime", "mmm-yyyy",    "none"),
    "text":  ("type text",   "string",   None,          "none"),
    "int":   ("Int64.Type",  "int64",    "#,0",         "sum"),
    "money": ("type number", "double",   '#,0 "SAR"',   "sum"),
    "pct":   ("type number", "double",   "0.0%",        "average"),
    "dec":   ("type number", "double",   "#,0.0",       "average"),
}

TABLES = {
    "tbl_Overview": [
        ("Month", "month"), ("Month_Label", "text"), ("Quarter", "text"),
        ("Website_Visits", "int"), ("New_Contacts", "int"), ("Leads", "int"), ("MQLs", "int"),
        ("SQLs", "int"), ("Target_Leads", "int"), ("Lead_Achievement", "pct"),
        ("Budget", "money"), ("Spend", "money"), ("Cost_per_Lead", "money"),
        ("Email_Sent", "int"), ("Social_Engagements", "int"),
        ("Cumulative_Leads", "int"), ("Cumulative_Target", "int"),
    ],
    "tbl_Campaigns": [
        ("Month", "month"), ("Month_Label", "text"), ("Campaign", "text"), ("Segment", "text"),
        ("Primary_Channel", "text"), ("Impressions", "int"), ("Clicks", "int"), ("Leads", "int"),
        ("MQLs", "int"), ("SQLs", "int"), ("Budget", "money"), ("Spend", "money"),
        ("CTR", "pct"), ("Cost_per_Lead", "money"), ("Lead_to_MQL", "pct"), ("MQL_to_SQL", "pct"),
    ],
    "tbl_Email": [
        ("Month", "month"), ("Month_Label", "text"), ("Campaign", "text"), ("Email_Name", "text"),
        ("Sent", "int"), ("Delivered", "int"), ("Bounced", "int"), ("Opens", "int"),
        ("Clicks", "int"), ("Unsubscribes", "int"), ("Delivery_Rate", "pct"), ("Open_Rate", "pct"),
        ("Click_Rate", "pct"), ("CTR_Click_to_Open", "pct"), ("Unsub_Rate", "pct"),
    ],
    "tbl_Social_Media": [
        ("Month", "month"), ("Month_Label", "text"), ("Platform", "text"), ("Followers", "int"),
        ("New_Followers", "int"), ("Posts", "int"), ("Impressions", "int"), ("Likes", "int"),
        ("Comments", "int"), ("Reposts", "int"), ("Clicks", "int"), ("Total_Engagements", "int"),
        ("Engagement_Rate", "pct"), ("Avg_Engagement_per_Post", "dec"),
    ],
    "tbl_Events": [
        ("Event_ID", "text"), ("Event_Name", "text"), ("Event_Date", "date"), ("Month", "month"),
        ("City", "text"), ("Event_Type", "text"), ("Campaign", "text"), ("Attendees", "int"),
        ("Leads", "int"), ("MQLs", "int"), ("SQLs", "int"), ("Budget", "money"),
        ("Actual_Cost", "money"), ("Cost_per_Lead", "money"), ("Lead_Conversion", "pct"), ("Status", "text"),
    ],
    "tbl_Database": [
        ("Country", "text"), ("Region", "text"), ("Segment", "text"), ("Contact_Type", "text"),
        ("Contacts", "int"), ("Opted_In", "int"), ("Engaged_90d", "int"),
        ("Opt_In_Rate", "pct"), ("Engagement_Rate", "pct"),
    ],
    "tbl_Budget": [
        ("Budget_Line", "text"), ("Category", "text"), ("Annual_Budget", "money"),
        ("Spend_YTD", "money"), ("Remaining", "money"), ("Pct_Used", "pct"), ("Status", "text"),
    ],
    "tbl_Targets": [
        ("KPI", "text"), ("Area", "text"), ("Target", "dec"), ("Actual_YTD", "dec"),
        ("Achievement", "pct"), ("Gap", "dec"), ("Status", "text"), ("Unit", "text"),
    ],
    "tbl_Activities": [
        ("Activity_ID", "text"), ("Activity", "text"), ("Campaign", "text"), ("Channel", "text"),
        ("Owner", "text"), ("Start_Date", "date"), ("End_Date", "date"), ("Status", "text"),
        ("Pct_Complete", "pct"), ("Deliverable", "text"), ("Notes", "text"),
    ],
    "tbl_Action_Center": [
        ("Action_ID", "text"), ("Priority", "text"), ("Action", "text"), ("Linked_KPI", "text"),
        ("Campaign", "text"), ("Owner", "text"), ("Due_Date", "date"), ("Status", "text"),
        ("Expected_Impact", "text"), ("Decision_Needed", "text"),
    ],
}

MEASURES = [
    # name, DAX, format
    ("Total Leads", "SUM(tbl_Overview[Leads])", "#,0"),
    ("Target Leads", "SUM(tbl_Overview[Target_Leads])", "#,0"),
    ("Lead Achievement %", "DIVIDE([Total Leads], [Target Leads])", "0.0%"),
    ("Total MQLs", "SUM(tbl_Overview[MQLs])", "#,0"),
    ("Total SQLs", "SUM(tbl_Overview[SQLs])", "#,0"),
    ("Lead to MQL %", "DIVIDE([Total MQLs], [Total Leads])", "0.0%"),
    ("MQL to SQL %", "DIVIDE([Total SQLs], [Total MQLs])", "0.0%"),
    ("Campaign & Event Spend", "SUM(tbl_Overview[Spend])", '#,0 "SAR"'),
    ("Cost per Lead", "DIVIDE([Campaign & Event Spend], [Total Leads])", '#,0 "SAR"'),
    ("Marketing Budget", "SUM(tbl_Budget[Annual_Budget])", '#,0 "SAR"'),
    ("Spend YTD", "SUM(tbl_Budget[Spend_YTD])", '#,0 "SAR"'),
    ("Budget Used %", "DIVIDE([Spend YTD], [Marketing Budget])", "0.0%"),
    ("Campaign Leads", "SUM(tbl_Campaigns[Leads])", "#,0"),
    ("Campaign MQLs", "SUM(tbl_Campaigns[MQLs])", "#,0"),
    ("Campaign SQLs", "SUM(tbl_Campaigns[SQLs])", "#,0"),
    ("Campaign Spend", "SUM(tbl_Campaigns[Spend])", '#,0 "SAR"'),
    ("Campaign Cost per Lead", "DIVIDE([Campaign Spend], [Campaign Leads])", '#,0 "SAR"'),
    ("Emails Sent", "SUM(tbl_Email[Sent])", "#,0"),
    ("Emails Delivered", "SUM(tbl_Email[Delivered])", "#,0"),
    ("Email Delivery Rate", "DIVIDE([Emails Delivered], [Emails Sent])", "0.0%"),
    ("Email Open Rate", "DIVIDE(SUM(tbl_Email[Opens]), [Emails Delivered])", "0.0%"),
    ("Email Click Rate", "DIVIDE(SUM(tbl_Email[Clicks]), [Emails Delivered])", "0.0%"),
    ("Email CTR (Click-to-Open)", "DIVIDE(SUM(tbl_Email[Clicks]), SUM(tbl_Email[Opens]))", "0.0%"),
    ("Social Posts", "SUM(tbl_Social_Media[Posts])", "#,0"),
    ("Social Engagements", "SUM(tbl_Social_Media[Total_Engagements])", "#,0"),
    ("Social Engagement Rate", "DIVIDE([Social Engagements], SUM(tbl_Social_Media[Impressions]))", "0.0%"),
    ("Saudi Contacts", "SUM(tbl_Database[Contacts])", "#,0"),
    ("Event Leads", "SUM(tbl_Events[Leads])", "#,0"),
    ("Events Delivered", 'CALCULATE(COUNTROWS(tbl_Events), tbl_Events[Status] = "Delivered")', "#,0"),
    ("Activities Delivered", 'CALCULATE(COUNTROWS(tbl_Activities), tbl_Activities[Status] = "Delivered")', "#,0"),
    ("Activities In Progress", 'CALCULATE(COUNTROWS(tbl_Activities), tbl_Activities[Status] = "In Progress")', "#,0"),
    ("Open Actions", 'CALCULATE(COUNTROWS(tbl_Action_Center), tbl_Action_Center[Status] <> "Done")', "#,0"),
]


def m_excel_table(xlsx_path: str, table: str, cols):
    types = ", ".join(f'{{"{c}", {KINDS[k][0]}}}' for c, k in cols)
    p = xlsx_path.replace("\\", "\\\\")
    return [
        "let",
        f'    Source = Excel.Workbook(File.Contents("{p}"), null, true),',
        f'    Data = Source{{[Item="{table}", Kind="Table"]}}[Data],',
        f"    Typed = Table.TransformColumnTypes(Data, {{{types}}})",
        "in",
        "    Typed",
    ]


def tmsl_column(name, kind, **extra):
    _, dtype, fmt, summ = KINDS[kind]
    col = {"name": name, "dataType": dtype, "sourceColumn": name, "summarizeBy": summ,
           "lineageTag": str(uuid.uuid4()),
           "annotations": [{"name": "SummarizationSetBy", "value": "User"}]}
    if fmt:
        col["formatString"] = fmt
    col.update(extra)
    return col


def tmsl_table(name, columns, m_lines, measures=None, hidden=False, **extra):
    t = {
        "name": name,
        "lineageTag": str(uuid.uuid4()),
        "columns": columns,
        "partitions": [{"name": name, "mode": "import",
                        "source": {"type": "m", "expression": m_lines}}],
        "annotations": [{"name": "PBI_ResultType", "value": "Table"}],
    }
    if measures:
        t["measures"] = measures
    if hidden:
        t["isHidden"] = True
    t.update(extra)
    return t


def build_model(xlsx_path: str):
    tables = []
    for tname, cols in TABLES.items():
        tables.append(tmsl_table(tname, [tmsl_column(c, k) for c, k in cols],
                                 m_excel_table(xlsx_path, tname, cols)))

    # Date dimension
    date_cols = [
        tmsl_column("Date", "date", isKey=True),
        tmsl_column("Year", "int", summarizeBy="none"),
        tmsl_column("Month_Number", "int", summarizeBy="none", isHidden=True),
        tmsl_column("Year_Month", "int", summarizeBy="none", isHidden=True),
        tmsl_column("Month_Name", "text", sortByColumn="Month_Number"),
        tmsl_column("Month_Label", "text", sortByColumn="Year_Month"),
        tmsl_column("Quarter", "text"),
        tmsl_column("Month_Start", "month"),
    ]
    date_m = [
        "let",
        "    Source = List.Dates(#date(2026, 1, 1), 365, #duration(1, 0, 0, 0)),",
        '    T0 = Table.FromList(Source, Splitter.SplitByNothing(), {"Date"}),',
        '    T1 = Table.TransformColumnTypes(T0, {{"Date", type date}}),',
        '    T2 = Table.AddColumn(T1, "Year", each Date.Year([Date]), Int64.Type),',
        '    T3 = Table.AddColumn(T2, "Month_Number", each Date.Month([Date]), Int64.Type),',
        '    T4 = Table.AddColumn(T3, "Year_Month", each Date.Year([Date]) * 100 + Date.Month([Date]), Int64.Type),',
        '    T5 = Table.AddColumn(T4, "Month_Name", each Date.ToText([Date], "MMM"), type text),',
        '    T6 = Table.AddColumn(T5, "Month_Label", each Date.ToText([Date], "MMM-yyyy"), type text),',
        '    T7 = Table.AddColumn(T6, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),',
        '    T8 = Table.AddColumn(T7, "Month_Start", each Date.StartOfMonth([Date]), type date)',
        "in",
        "    T8",
    ]
    tables.append(tmsl_table("Dim_Date", date_cols, date_m, dataCategory="Time"))

    # Campaign dimension
    camp_rows = [("HVAC", "Commercial Buildings"), ("Data Centre", "Data Centres"),
                 ("Wastewater", "Water Utilities"), ("Hospitality", "Hotels & Resorts"),
                 ("Healthcare", "Hospitals"), ("All", "Cross-segment")]
    rows = ", ".join(f'{{"{c}", "{s}"}}' for c, s in camp_rows)
    camp_m = ["let",
              f"    Source = #table(type table [Campaign = text, Segment = text], {{{rows}}})",
              "in", "    Source"]
    tables.append(tmsl_table("Dim_Campaign",
                             [tmsl_column("Campaign", "text", isKey=True), tmsl_column("Segment", "text")],
                             camp_m))

    # Measures table
    measures = [{"name": n, "expression": e, "formatString": f, "lineageTag": str(uuid.uuid4())}
                for n, e, f in MEASURES]
    kpi_m = ["let", '    Source = #table(type table [Dummy = Int64.Type], {})', "in", "    Source"]
    tables.append(tmsl_table("KPIs", [tmsl_column("Dummy", "int", isHidden=True, summarizeBy="none")],
                             kpi_m, measures=measures))

    relationships = []

    def rel(from_t, from_c, to_t, to_c, **extra):
        r = {"name": str(uuid.uuid4()), "fromTable": from_t, "fromColumn": from_c,
             "toTable": to_t, "toColumn": to_c}
        r.update(extra)
        relationships.append(r)

    for t in ["tbl_Overview", "tbl_Campaigns", "tbl_Email", "tbl_Social_Media", "tbl_Events"]:
        rel(t, "Month", "Dim_Date", "Date", joinOnDateBehavior="datePartOnly")
    for t in ["tbl_Campaigns", "tbl_Email", "tbl_Events", "tbl_Activities", "tbl_Action_Center"]:
        rel(t, "Campaign", "Dim_Campaign", "Campaign")

    return {
        "name": NAME,
        "compatibilityLevel": 1567,
        "model": {
            "culture": "en-US",
            "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "sourceQueryCulture": "en-US",
            "tables": tables,
            "relationships": relationships,
            "cultures": [{"name": "en-US", "linguisticMetadata": {
                "content": {"Version": "1.0.0", "Language": "en-US"},
                "contentType": "json"}}],
            "annotations": [
                {"name": "PBI_QueryOrder", "value": json.dumps([t["name"] for t in tables])},
                {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
                {"name": "PBIDesktopVersion", "value": "2.159.1231.0 (25.11)"},
            ],
        },
    }


# --------------------------------------------------------------------------- #
# Report
# --------------------------------------------------------------------------- #
def _lit(v):
    return {"expr": {"Literal": {"Value": v}}}


def _vc(x, y, w, h, single_visual, z=0):
    cfg = {"name": uuid.uuid4().hex[:20],
           "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h}}],
           "singleVisual": single_visual}
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(cfg), "filters": "[]"}


def _sel_measure(src, table, name):
    return {"Measure": {"Expression": {"SourceRef": {"Source": src}}, "Property": name},
            "Name": f"{table}.{name}"}


def _sel_column(src, table, name):
    return {"Column": {"Expression": {"SourceRef": {"Source": src}}, "Property": name},
            "Name": f"{table}.{name}"}


def _title(text):
    return {"title": [{"properties": {"show": _lit("true"), "text": _lit(f"'{text}'"),
                                      "fontColor": _lit(f"'{BLUE_DARK}'"), "fontSize": _lit("11D"),
                                      "bold": _lit("true")}}]}


def card(x, y, w, h, measure, title):
    sv = {"visualType": "card",
          "projections": {"Values": [{"queryRef": f"KPIs.{measure}"}]},
          "prototypeQuery": {"Version": 2, "From": [{"Name": "k", "Entity": "KPIs", "Type": 0}],
                             "Select": [_sel_measure("k", "KPIs", measure)]},
          "drillFilterOtherVisuals": True,
          "objects": {"labels": [{"properties": {"color": _lit(f"'{BLUE}'"), "fontSize": _lit("24D")}}],
                      "categoryLabels": [{"properties": {"show": _lit("false")}}]},
          "vcObjects": _title(title)}
    return _vc(x, y, w, h, sv)


def textbox(x, y, w, h, text, size="20pt", color=BLUE_DARK):
    sv = {"visualType": "textbox", "drillFilterOtherVisuals": True,
          "objects": {"general": [{"properties": {"paragraphs": [{"textRuns": [
              {"value": text, "textStyle": {"fontWeight": "bold", "fontSize": size, "color": color}}]}]}}]}}
    return _vc(x, y, w, h, sv)


def combo_leads_vs_target(x, y, w, h):
    sv = {"visualType": "lineClusteredColumnComboChart",
          "projections": {"Category": [{"queryRef": "Dim_Date.Month_Label"}],
                          "Y": [{"queryRef": "KPIs.Total Leads"}],
                          "Y2": [{"queryRef": "KPIs.Target Leads"}]},
          "prototypeQuery": {"Version": 2,
                             "From": [{"Name": "d", "Entity": "Dim_Date", "Type": 0},
                                      {"Name": "k", "Entity": "KPIs", "Type": 0}],
                             "Select": [_sel_column("d", "Dim_Date", "Month_Label"),
                                        _sel_measure("k", "KPIs", "Total Leads"),
                                        _sel_measure("k", "KPIs", "Target Leads")],
                             "OrderBy": [{"Direction": 1, "Expression": {"Column": {
                                 "Expression": {"SourceRef": {"Source": "d"}}, "Property": "Month_Label"}}}]},
          "drillFilterOtherVisuals": True,
          "objects": {"dataPoint": [{"properties": {"fill": {"solid": {"color": _lit(f"'{BLUE}'")}}},
                                     "selector": {"metadata": "KPIs.Total Leads"}},
                                    {"properties": {"fill": {"solid": {"color": _lit(f"'{BLUE_DARK}'")}}},
                                     "selector": {"metadata": "KPIs.Target Leads"}}]},
          "vcObjects": _title("Leads vs target by month")}
    return _vc(x, y, w, h, sv)


def bar_by_campaign(x, y, w, h):
    sv = {"visualType": "clusteredBarChart",
          "projections": {"Category": [{"queryRef": "Dim_Campaign.Campaign"}],
                          "Y": [{"queryRef": "KPIs.Campaign Leads"},
                                {"queryRef": "KPIs.Campaign MQLs"},
                                {"queryRef": "KPIs.Campaign SQLs"}]},
          "prototypeQuery": {"Version": 2,
                             "From": [{"Name": "c", "Entity": "Dim_Campaign", "Type": 0},
                                      {"Name": "k", "Entity": "KPIs", "Type": 0}],
                             "Select": [_sel_column("c", "Dim_Campaign", "Campaign"),
                                        _sel_measure("k", "KPIs", "Campaign Leads"),
                                        _sel_measure("k", "KPIs", "Campaign MQLs"),
                                        _sel_measure("k", "KPIs", "Campaign SQLs")],
                             "OrderBy": [{"Direction": 2, "Expression": {"Measure": {
                                 "Expression": {"SourceRef": {"Source": "k"}}, "Property": "Campaign Leads"}}}]},
          "drillFilterOtherVisuals": True,
          "vcObjects": _title("Funnel by campaign (Leads · MQLs · SQLs)")}
    return _vc(x, y, w, h, sv)


def table_visual(x, y, w, h, title, entity, columns=(), measures=(), src="t"):
    projections, select = [], []
    froms = [{"Name": src, "Entity": entity, "Type": 0}]
    for c in columns:
        projections.append({"queryRef": f"{entity}.{c}"})
        select.append(_sel_column(src, entity, c))
    if measures:
        froms.append({"Name": "k", "Entity": "KPIs", "Type": 0})
        for m in measures:
            projections.append({"queryRef": f"KPIs.{m}"})
            select.append(_sel_measure("k", "KPIs", m))
    sv = {"visualType": "tableEx",
          "projections": {"Values": projections},
          "prototypeQuery": {"Version": 2, "From": froms, "Select": select},
          "drillFilterOtherVisuals": True,
          "objects": {"columnHeaders": [{"properties": {
              "backColor": {"solid": {"color": _lit(f"'{BLUE_DARK}'")}},
              "fontColor": {"solid": {"color": _lit("'#FFFFFF'")}}}}]},
          "vcObjects": _title(title)}
    return _vc(x, y, w, h, sv)


def slicer(x, y, w, h, entity, column, title):
    sv = {"visualType": "slicer",
          "projections": {"Values": [{"queryRef": f"{entity}.{column}"}]},
          "prototypeQuery": {"Version": 2, "From": [{"Name": "s", "Entity": entity, "Type": 0}],
                             "Select": [_sel_column("s", entity, column)]},
          "drillFilterOtherVisuals": True,
          "objects": {"general": [{"properties": {"orientation": _lit("1D")}}],
                      "data": [{"properties": {"mode": _lit("'Basic'")}}]},
          "vcObjects": _title(title)}
    return _vc(x, y, w, h, sv)


def section(name, display, visuals):
    return {"name": name, "displayName": display, "displayOption": 1, "height": 720, "width": 1280,
            "visualContainers": visuals, "config": "{}", "filters": "[]",
            "objects": {"background": [{"properties": {"color": {"solid": {"color": _lit("'#FFFFFF'")}},
                                                       "transparency": _lit("0D")}}]}}


def build_report():
    gap, cw, ch = 12, 200, 84
    overview = [
        textbox(20, 10, 900, 44, "Grundfos Saudi Arabia – Marketing Dashboard FY2026"),
        slicer(950, 8, 310, 50, "Dim_Campaign", "Campaign", "Campaign"),
    ]
    cards = [("Total Leads", "Total leads YTD"), ("Target Leads", "Target leads"),
             ("Lead Achievement %", "Lead achievement"), ("Total MQLs", "MQLs"),
             ("Total SQLs", "SQLs"), ("Saudi Contacts", "Saudi database contacts")]
    for i, (m, t) in enumerate(cards):
        overview.append(card(20 + i * (cw + gap), 66, cw, ch, m, t))
    cards2 = [("Marketing Budget", "Marketing budget"), ("Spend YTD", "Spend YTD"),
              ("Budget Used %", "Budget used"), ("Cost per Lead", "Cost per lead"),
              ("Email Open Rate", "Email open rate"), ("Email Click Rate", "Email click rate")]
    for i, (m, t) in enumerate(cards2):
        overview.append(card(20 + i * (cw + gap), 66 + ch + gap, cw, ch, m, t))
    top = 66 + 2 * (ch + gap)
    overview.append(combo_leads_vs_target(20, top, 620, 250))
    overview.append(bar_by_campaign(652, top, 608, 250))
    top2 = top + 262
    overview.append(table_visual(20, top2, 620, 720 - top2 - 12, "Email performance by campaign",
                                 "Dim_Campaign", ["Campaign"],
                                 ["Emails Delivered", "Email Open Rate", "Email Click Rate",
                                  "Email CTR (Click-to-Open)"], src="c"))
    overview.append(table_visual(652, top2, 608, 720 - top2 - 12, "Social media by platform",
                                 "tbl_Social_Media", ["Platform"],
                                 ["Social Posts", "Social Engagements", "Social Engagement Rate"]))

    events = [
        textbox(20, 10, 900, 44, "Events & campaigns"),
        table_visual(20, 66, 1240, 300, "Events – leads, MQLs, SQLs", "tbl_Events",
                     ["Event_Name", "Event_Date", "City", "Event_Type", "Campaign", "Attendees",
                      "Leads", "MQLs", "SQLs", "Actual_Cost", "Cost_per_Lead", "Status"]),
        table_visual(20, 378, 1240, 330, "Budget lines", "tbl_Budget",
                     ["Budget_Line", "Category", "Annual_Budget", "Spend_YTD", "Remaining", "Pct_Used", "Status"]),
    ]

    actions = [
        textbox(20, 10, 900, 44, "Executive Action Center"),
        card(950, 8, 150, 50, "Open Actions", "Open actions"),
        card(1110, 8, 150, 50, "Activities In Progress", "In progress"),
        table_visual(20, 66, 1240, 300, "Actions requiring decisions", "tbl_Action_Center",
                     ["Priority", "Action", "Linked_KPI", "Campaign", "Owner", "Due_Date", "Status",
                      "Expected_Impact", "Decision_Needed"]),
        table_visual(20, 378, 1240, 330, "Activities – delivered / in progress / planned", "tbl_Activities",
                     ["Activity", "Campaign", "Channel", "Owner", "Start_Date", "End_Date", "Status",
                      "Pct_Complete", "Deliverable"]),
    ]

    config = {"version": "5.43", "themeCollection": {"baseTheme": {"name": "CY24SU06", "version": "5.47", "type": 2}},
              "activeSectionIndex": 0, "defaultDrillFilterOtherVisuals": True,
              "settings": {"useNewFilterPaneExperience": True, "allowChangeFilterTypes": True,
                           "useStylableVisualContainerHeader": True, "exportDataMode": 1,
                           "useEnhancedTooltips": True}}
    return {
        "config": json.dumps(config),
        "layoutOptimization": 0,
        "publicCustomVisuals": [],
        "resourcePackages": [{"resourcePackage": {"name": "SharedResources", "type": 2,
                              "items": [{"name": "CY24SU06", "path": "BaseThemes/CY24SU06.json", "type": 202}]}}],
        "sections": [section("ReportSection1", "Overview", overview),
                     section("ReportSection2", "Events & Budget", events),
                     section("ReportSection3", "Action Center", actions)],
        "filters": "[]",
        "pods": [],
    }


def write_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, ensure_ascii=False), encoding="utf-8")


def build(xlsx_src: Path, out: Path):
    out.mkdir(parents=True, exist_ok=True)
    xlsx_dst = out / xlsx_src.name
    if xlsx_src.resolve() != xlsx_dst.resolve():
        shutil.copy2(xlsx_src, xlsx_dst)

    sm = out / f"{NAME}.SemanticModel"
    rp = out / f"{NAME}.Report"

    write_json(out / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })
    write_json(sm / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "1.0", "settings": {}})
    write_json(sm / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}})
    write_json(sm / "model.bim", build_model(str(xlsx_dst)))

    write_json(rp / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/1.0.0/schema.json",
        "version": "1.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    write_json(rp / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NAME},
        "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}})
    write_json(rp / "report.json", build_report())
    (out / ".gitignore").write_text("**/.pbi/localSettings.json\n**/.pbi/cache.abf\n", encoding="utf-8")
    return out / f"{NAME}.pbip"


if __name__ == "__main__":
    xlsx = Path(sys.argv[1])
    out = Path(sys.argv[2])
    print("Project written:", build(xlsx, out))
