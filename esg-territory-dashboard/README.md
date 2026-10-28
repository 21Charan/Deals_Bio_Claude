# ESG Territory Hours Dashboard

Builds two offline HTML pages from one Excel workbook:

| Page | What it is |
|---|---|
| `ESG_US_Story_v2.html` | Partner story: why tagging US-focused AC people to the US is a win-win, with the "same budget, more capacity" calculator and the long-term saving view. Arrow keys move between sections. |
| `ESG_Hours_Dashboard.html` | Full dashboard: team, Energy Transition themes, territory hours (FY/CY), territory detail, US vs Non-US, growth, rate analysis and the interactive US tagging scenario planner. |

Both pages are single self-contained files (no internet needed) and link to each other.

## Quick start (Windows)

1. Install Python 3.10+ and the two packages (once):
   ```
   py -m pip install -r requirements.txt
   ```
2. Double-click **`Build Dashboards.bat`** (or run `python run.py`).
   It checks the workbook's input tabs, builds both pages and opens the story.

## Use your actual data

1. Copy `local_settings.example.py` to **`local_settings.py`** and fill in your workbook path, contact name/email and logo file name.
   `local_settings.py` is ignored by Git, so your private values are never committed.
2. Optional: copy your logo into `engine/` as `logo.png` (also ignored by Git).
3. Close the workbook in Excel and double-click `Build Dashboards.bat`.
   The first run adds the input tabs to your workbook (pre-filled) and keeps an untouched backup copy.
4. Read the **Data check** printed by the build: it lists unknown staff, missing sales prices, levels missing from the rate card and which input tabs were found.

Ignored by Git (`.gitignore`), so they are never committed by accident: your real workbooks, the generated HTML pages (they embed the data), `local_settings.py` and logo files.

The sample workbook uses made-up hours, people, clients and an **illustrative** rate card - replace the rate card (`Global Model FY27 CPP` tab) and the Team/Story tabs with your own values in your private workbook.

## Workbook structure

Data tabs (keep the column headers):

| Tab | Content |
|---|---|
| Project Hours Raw Data | One row per time entry: Territory Name, Staff Name, Management Level, Team Name, Client code/WBS code, Client Name/WBS Name, Job Name, Hours, Date, Memo, Work Request No., Source Sheet, Project Name, Project Match Flag, ESG/ET |
| Staff Master | Employee, Level, Team Name, Home Cost Center, Home Sub LOS 2, Sales Price (or Sales Price 2025, Sales Price 2026, ...) |
| Global Model FY27 CPP | Level, Level Name, Rate/hr, Cost/Person |
| Callouts | Month (Mon-YY), Territory, Event |
| WRN Category | Work Request No. → ESG / ET (fills blank ESG/ET cells) |

Input tabs (static content you maintain; each tab has a "How to fill" note):

| Tab | Content |
|---|---|
| README | What every tab is and who updates it |
| Assumptions | FY start month, tagging threshold, look-back months, annual capacity, training month, report title, classification, current tagging note |
| Tag Mapping | Which territory each Home Sub LOS 2 tag means |
| Team / Services / ET Themes | Team overview and Energy Transition pages |
| Story | All partner-story wording (title, asks, win-win points, next steps) |
| Scenario Presets | Conservative / Base / Growth settings |

Notes:
- **Reporting month** comes from `Source Sheet` (falls back to `Date`).
- Territory spellings like `US`, `United States`, `AU` are normalised.
- A single `Sales Price` column is used for every year; `Sales Price YYYY` columns give year-specific prices.

## How the saving is calculated

- **Cost today** = US hours × the person's sales price (Staff Master).
- **Cost once US-tagged** = the same hours × the Global Model Rate/hr for their level.
- **Dedicated seat** = the level's Cost/Person for one full-time person (up to the annual capacity, default 1,800 hrs).
- **Recommended tag** = the territory with at least the threshold share (default 60%) of a person's hours in the look-back window.
- Confirm with Finance that US-tagged hours are charged at the Global Model rate and how Cost/Person applies to a seat.

## Repository layout

```
Build Dashboards.bat     double-click to rebuild everything
run.py                   default settings + build steps
local_settings.example.py  template for your private settings (copy to local_settings.py)
requirements.txt
ESG_Hours_Sample.xlsx    sample workbook (made-up hours) with all tabs
engine/
  build_dashboard.py     workbook -> dashboard page (also reads all input tabs)
  build_story_v2.py      workbook -> partner story page
  add_input_tabs.py      adds missing input tabs to a workbook
  chart.umd.min.js       Chart.js 4.4.1 (bundled for offline use)
  logo.png               your logo (not in the repo - add your own; ignored by Git)
  team_content.json      fallback text if a workbook has no Team/Services/ET Themes tabs
```

Advanced: the engine scripts can also be run on their own, e.g.
`python engine/build_dashboard.py --excel "D:\Data\Hours.xlsx" --out "D:\Data\Dashboard.html"`.
