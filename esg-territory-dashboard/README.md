# ESG Territory Hours Dashboard

Three files build two offline HTML pages from one Excel workbook:

| File | What it is |
|---|---|
| `build.py` | All the code. Settings at the top. |
| `template.html` | Design of both pages plus the bundled chart library (Chart.js 4.4.1), so the pages work offline. |
| `ESG_Hours_Sample.xlsx` | Data tabs plus input tabs (team, story text, scenario presets...). Made-up sample data. |

Output (written next to `build.py`):

| Page | What it is |
|---|---|
| `ESG_US_Story_v2.html` | Partner story: why tagging US-focused people to the US is a win-win, with the "same budget, more capacity" calculator and the long-term saving view. Arrow keys move between sections. |
| `ESG_Hours_Dashboard.html` | Full dashboard: team, Energy Transition themes, territory hours (FY/CY), territory detail, US vs Non-US, growth, rate analysis and the US tagging scenario planner. |

## Run it

```
py -m pip install pandas openpyxl     # once
py build.py                           # or double-click build.py
```

It adds any missing input tabs to the workbook (first run only; a backup copy is kept), prints a **data check**, writes both pages and opens the story.

Options: `py build.py --excel "D:\Data\Hours.xlsx" --out-dir "D:\Out" --no-open`

## Use your actual data

1. Put your workbook anywhere and create **`local_settings.py`** next to `build.py`:
   ```python
   EXCEL_FILE = r"C:\path\to\Your_Hours.xlsx"
   CONTACT_NAME = "Your Name"
   CONTACT_EMAIL = "you@example.com"
   LOGO_FILE = "logo.png"          # optional logo next to build.py
   ```
   Any setting from the top of `build.py` can go in this file.
2. Close the workbook in Excel and run `py build.py`.
3. Read the data check: it lists unknown staff, missing sales prices, levels missing from the rate card and the input tabs found.

Git ignores `local_settings.py`, logo files, real workbooks and the built pages (they embed your data), so none of them are committed by accident.

The sample workbook uses made-up hours, people and clients and an **illustrative** rate card. Put your real rate card (`Global Model FY27 CPP` tab) and your Team / Story text in your own workbook.

## Workbook tabs

Data (keep the column headers):

| Tab | Content |
|---|---|
| Project Hours Raw Data | One row per time entry: Territory Name, Staff Name, Management Level, Team Name, Client code/WBS code, Client Name/WBS Name, Job Name, Hours, Date, Memo, Work Request No., Source Sheet, Project Name, Project Match Flag, ESG/ET |
| Staff Master | Employee, Level, Team Name, Home Cost Center, Home Sub LOS 2, Sales Price (or Sales Price 2025, Sales Price 2026, ...) |
| Global Model FY27 CPP | Level, Level Name, Rate/hr, Cost/Person |
| Callouts | Month (Mon-YY), Territory, Event |
| WRN Category | Work Request No. → ESG / ET (fills blank ESG/ET cells) |

Inputs (each tab has a "How to fill" note): README, Assumptions, Tag Mapping, Team, Services, ET Themes, Story, Scenario Presets.

- Reporting month comes from `Source Sheet` (falls back to `Date`).
- Territory spellings like `US`, `United States`, `AU` are normalised.
- One `Sales Price` column is used for every year; `Sales Price YYYY` columns give year-specific prices.

## How the saving is calculated

- **Cost today** = US hours × the person's sales price (Staff Master).
- **Cost once US-tagged** = the same hours × the rate card's Rate/hr for their level.
- **Dedicated seat** = the level's Cost/Person for one full-time person (up to the annual capacity, default 1,800 hrs).
- **Recommended tag** = the territory with at least the threshold share (default 60%) of a person's hours in the look-back window.
- Confirm with Finance that tagged hours are charged at the rate card and how Cost/Person applies to a seat.

## Changing the design

Edit `template.html`. It has three sections, each starting with a marker line: `<!-- @@DASHBOARD -->`, `<!-- @@STORY -->` and `<!-- @@CHARTJS -->`. `build.py` fills the placeholders `/*__DATA__*/null`, `<!--__LOGO__-->` and `/*__CHARTJS__*/`.
