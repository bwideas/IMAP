# Marketing Dashboard Workbook (Power BI-ready, dummy data)

`build_marketing_dashboard.py` generates `Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx`,
a master workbook for the IMAP marketing dashboard filled with dummy 2026 data.

* Every data sheet is an Excel Table (`tbl_*`) so Power BI Desktop can import it by name
  (Get Data → Excel workbook → tick the `tbl_*` tables).
* All `tbl_*` tables are formula-driven from one **Inputs** sheet – change an input and
  every table and KPI updates.
* Table names and columns are unchanged from v1, so existing Power BI connections keep working.
* If `pywin32` and Excel are installed, the script opens the file through Excel COM once to
  pre-calculate the formulas so Power BI sees values immediately. Without them the workbook is
  still valid; Excel calculates on first open.

## Run

```
pip install openpyxl pywin32
python build_marketing_dashboard.py [output_path]
```

## Inputs sheet (v2)

Light-blue cells are the only cells you need to edit.

* **Spinners** – Year, Target Leads, Saudi contacts, email delivery / open / click rates.
* **Month dropdown** – Data-as-of month (drives YTD, cumulative and status columns).
* **Scenario option buttons** – Conservative / Base / Optimistic (scales volumes and spend).
* **Checkboxes** – switch each campaign and social platform on or off.
* **Input tables** – campaigns (annual budget, lead mix), social platforms, event types,
  monthly seasonality, database segments and regions.

Workflow: edit Inputs → Save → Power BI Desktop → Home → Refresh.

## Tables

| Table | Content |
|---|---|
| `tbl_Overview` | 12-month roll-up: leads, MQLs, SQLs, target, budget, spend, cumulative (formula-driven) |
| `tbl_Campaigns` | HVAC, Data Centre, Wastewater, Hospitality, Healthcare × 12 months |
| `tbl_Email` | Sent, delivered, opens, clicks + delivery/open/click/CTR formulas |
| `tbl_Social_Media` | LinkedIn, Facebook, WhatsApp Channel: posts, likes, comments, reposts, clicks |
| `tbl_Events` | Events with attendees, leads, MQLs, SQLs, budget vs cost |
| `tbl_Database` | Saudi contact database by region and segment |
| `tbl_Budget` | Budget lines, spend YTD, remaining, % used |
| `tbl_Targets` | KPI scorecard against targets |
| `tbl_Activities` | Marketing activities – Delivered / In Progress / Planned |
| `tbl_Action_Center` | Executive action list – priorities, owners, decisions needed |

## Power BI project

`build_pbip.py` generates a Power BI Project (.pbip) from the workbook: typed tables,
`Dim_Date` and `Dim_Campaign` dimensions with relationships, KPI measures and three starter
report pages.

```
python build_pbip.py Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx <output_folder>
```

Open the generated `.pbip` in Power BI Desktop (Preview feature "Power BI Project (.pbip) save
option" must be enabled).

## v2 audit fixes

* Removed hard-coded as-of date and Target Leads; both now come from Inputs.
* Replaced the array formula with `MAXIFS`.
* Campaign budget rounding fixed so the twelve monthly budgets sum exactly to the annual figure.
* Status columns use conditional formatting.
* Dashboard column widths corrected.
