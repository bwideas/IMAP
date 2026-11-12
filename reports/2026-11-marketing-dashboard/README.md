# Marketing Dashboard Workbook (Power BI-ready, dummy data)

`build_marketing_dashboard.py` generates `Grundfos_Marketing_Dashboard_Master_With_Dummy_Data.xlsx`,
a master workbook for the IMAP marketing dashboard filled with dummy 2026 data.

* Every data sheet is an Excel Table (`tbl_*`) so Power BI Desktop can import it by name
  (Get Data → Excel workbook → tick the `tbl_*` tables).
* Rates, totals and KPI roll-ups are real Excel formulas, not pasted values.
* If `pywin32` and Excel are installed, the script opens the file through Excel COM once to
  pre-calculate the formulas so Power BI sees values immediately. Without them the workbook is
  still valid; Excel calculates on first open.

## Run

```
pip install openpyxl pywin32
python build_marketing_dashboard.py [output_path]
```

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

Replace the dummy rows with real figures (or point the script at real sources) and refresh
Power BI; the table names and columns stay the same.
