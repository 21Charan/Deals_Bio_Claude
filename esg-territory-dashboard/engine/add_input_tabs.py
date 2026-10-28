"""
Add the input tabs (static content you maintain) to an hours workbook.

    python add_input_tabs.py                       # the workbook set in build_dashboard.py
    python add_input_tabs.py --excel "D:\\Data\\Hours.xlsx"

- Adds only tabs that are missing; existing tabs (and all data tabs) are left untouched.
- New tabs are pre-filled with the current content, so you only edit what changes.
- Close the workbook in Excel before running.
"""
import argparse
import json
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import build_dashboard as bd

HERE = Path(__file__).resolve().parent
ORANGE, CHARCOAL, PEACH = "D04A02", "2D2D2D", "FDEDE4"


def story_defaults():
    try:
        import build_story_v2 as sv
        return sv.STORY
    except Exception:
        return {}


def tables(wb):
    """Return {sheet name: (rows, column widths, note)} for every input tab."""
    content = json.loads((HERE / bd.CONTENT_FILE).read_text(encoding="utf-8")) if (HERE / bd.CONTENT_FILE).exists() else {}
    team, themes = content.get("team", {}), content.get("themes", {})
    story = story_defaults()
    out = {}

    # Tag Mapping: every Home Sub LOS 2 value found in Staff Master
    tags = []
    if bd.SHEET_STAFF in wb.sheetnames:
        sm = pd.DataFrame(wb[bd.SHEET_STAFF].values)
        if len(sm):
            sm.columns = [str(c).strip() for c in sm.iloc[0]]; sm = sm.iloc[1:]
            if "Home Sub LOS 2" in sm.columns:
                tags = sorted({str(v).strip() for v in sm["Home Sub LOS 2"].dropna() if str(v).strip()})
    def guess(tag):
        import re
        for pat, terr in bd.TAG_KEYWORDS.items():
            if re.search(pat, tag, re.I):
                return terr
        return ""
    out[bd.SHEET_TAG_MAP] = ([["Home Sub LOS 2", "Territory"]] + [[t, guess(t)] for t in tags],
                             [36, 16], "Which territory each Home Sub LOS 2 tag means. Leave Territory blank for tags that are not a territory (e.g. Other Territory, Deals).")

    rows = [["Type", "Name", "Role", "Territory", "Note"],
            ["Team name", team.get("teamName", ""), "", "", ""],
            ["Summary", "", "", "", team.get("summary", "")],
            ["Members", team.get("members", ""), "", "", ""]]
    rows += [["Leader", l.get("name", ""), l.get("role", ""), "", l.get("note", "")] for l in team.get("leaders", [])]
    rows += [["Key contact", k.get("name", ""), k.get("role", ""), k.get("territory", ""), ""] for k in team.get("keyContacts", [])]
    out[bd.SHEET_TEAM] = (rows, [14, 26, 20, 14, 70], "Type must be one of: Team name, Summary, Members, Leader, Key contact. Add or delete Leader / Key contact rows freely.")

    rows = [["Category", "Item"]]
    for cat, key in [("Current service", "currentServices"), ("Future service", "futureServices"), ("ESG framework", "frameworks"),
                     ("Database", "databases"), ("Digital & AI tool", "aiTools")]:
        rows += [[cat, x] for x in team.get(key, [])]
    out[bd.SHEET_SERVICES] = (rows, [20, 50], "Category must be one of: Current service, Future service, ESG framework, Database, Digital & AI tool.")

    rows = [["Theme No", "Theme Title", "Tagline", "Work Area", "Description"]]
    for i, it in enumerate(themes.get("items", []), 1):
        for j, (a, d) in enumerate(it.get("points", [])):
            rows.append([i, it["title"] if j == 0 else "", it["tagline"] if j == 0 else "", a, d])
    out[bd.SHEET_THEMES] = (rows, [10, 34, 40, 30, 60], "One row per work area. Theme Title and Tagline only need filling on the first row of each theme. Up to 4 themes are shown.")

    rows = [["Section", "Text", "Heading", "Detail"]]
    for k in ("title", "subtitle", "meeting", "presenter"):
        if story.get(k): rows.append([k.capitalize(), story[k], "", ""])
    rows += [["Ask", x, "", ""] for x in story.get("ask", [])]
    rows += [["Win for US", x, "", ""] for x in story.get("winUS", [])]
    rows += [["Win for AC", x, "", ""] for x in story.get("winAC", [])]
    rows += [["Next step", w, t, d] for w, t, d in story.get("steps", [])]
    out[bd.SHEET_STORY] = (rows, [14, 70, 34, 60], "Section: Title, Subtitle, Meeting, Presenter, Ask, Win for US, Win for AC, Next step. For Next step, Text = when (e.g. 'Now').")

    rows = [["Name", "Tagging threshold", "Look-back months", "US volume change next year", "Costing method"]]
    for p in bd.SCENARIO_PRESETS:
        rows.append([p["name"], p["threshold"], p["window"], p["growth"], "Hourly or seat" if p["method"] == "best" else "Hourly"])
    out[bd.SHEET_PRESETS] = (rows, [16, 18, 18, 26, 18], "Blank threshold / look-back = use the Assumptions tab. Costing method: Hourly, or Hourly or seat.")
    return out


README = [
    ["Tab", "What it is", "Who updates it", "How"],
    [bd.SHEET_RAW, "Time entries (one row per entry)", "Monthly export", "Paste the new export; keep the column headers"],
    [bd.SHEET_STAFF, "People: level, team, Home Sub LOS 2 tag, sales price", "When people join, move or rates change", "One row per person; 'Sales Price' or 'Sales Price 2025', 'Sales Price 2026'..."],
    [bd.SHEET_MODEL, "Global Model rate card (Rate/hr, Cost/Person by level)", "Once a year", "Levels must match Staff Master (SM, M, SA, A...)"],
    [bd.SHEET_CALLOUTS, "Milestones shown on the monthly chart", "As needed", "Month as Mon-YY, e.g. Feb-25"],
    [bd.SHEET_ASSUMPTIONS, "Settings: FY start, tagging threshold, look-back, titles", "Rarely", "Change the Value column only"],
    [bd.SHEET_WRN_CAT, "ESG or ET per Work Request No.", "When new WRNs appear", "Fills blank ESG/ET cells in the raw data"],
    [bd.SHEET_TAG_MAP, "Which territory each Home Sub LOS 2 tag means", "When a new tag appears", "Blank Territory = not a territory tag"],
    [bd.SHEET_TEAM, "Team overview: leaders, members, key contacts", "As the team changes", "See the note on the tab"],
    [bd.SHEET_SERVICES, "Services, frameworks, databases, AI tools", "As needed", "See the note on the tab"],
    [bd.SHEET_THEMES, "Energy Transition themes page", "As needed", "See the note on the tab"],
    [bd.SHEET_STORY, "Text of the partner story page", "Before each meeting", "See the note on the tab"],
    [bd.SHEET_PRESETS, "Scenario presets (Conservative / Base / Growth)", "As needed", "See the note on the tab"],
    ["", "", "", ""],
    ["After editing", "Save and close the workbook, then run:  python build_dashboard.py  and  python build_story_v2.py", "", ""],
]
EXTRA_ASSUMPTIONS = [
    ("Report title", bd.REPORT_TITLE, "Title at the top of the dashboard."),
    ("Classification", bd.CLASSIFICATION, "Shown in the header and footer of both pages."),
    ("Energy Transition page title", "Our Work Across Key Energy Transition Themes", "Heading of the Energy Transition tab."),
    ("Current tagging", bd.DEFAULT_TAGGING_NOTE, "Shown in the assumptions on both pages. Update once people are tagged to territories."),
]


def style_sheet(ws, widths, header_fill=CHARCOAL):
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=header_fill)
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", default=str(bd.EXCEL_PATH))
    args = ap.parse_args()
    path = Path(args.excel).expanduser()
    if not path.exists():
        sys.exit(f"Excel file not found: {path}")
    wb = load_workbook(path)
    added = []

    for name, (rows, widths, note) in tables(wb).items():
        if name in wb.sheetnames:
            continue
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
        style_sheet(ws, widths, ORANGE)
        ws.sheet_properties.tabColor = ORANGE
        nc = ws.cell(row=1, column=len(rows[0]) + 2, value="How to fill: " + note)   # beside the table, so it is never read as data
        nc.font = Font(italic=True, color="7D7D7D"); nc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[get_column_letter(len(rows[0]) + 2)].width = 48
        added.append(name)

    # Assumptions: add the title rows if missing
    if bd.SHEET_ASSUMPTIONS in wb.sheetnames:
        ws = wb[bd.SHEET_ASSUMPTIONS]
        have = {str(r[0].value).strip() for r in ws.iter_rows(min_row=2) if r[0].value}
        for p, v, n in EXTRA_ASSUMPTIONS:
            if p not in have:
                ws.append([p, v, n]); added.append(f"Assumptions: {p}")
    else:
        ws = wb.create_sheet(bd.SHEET_ASSUMPTIONS)
        ws.append(["Parameter", "Value", "Notes"])
        for p, v in bd.DEFAULT_ASSUMPTIONS.items():
            ws.append([p, v, ""])
        for p, v, n in EXTRA_ASSUMPTIONS:
            ws.append([p, v, n])
        style_sheet(ws, [30, 40, 70]); added.append(bd.SHEET_ASSUMPTIONS)

    if "README" not in wb.sheetnames:
        ws = wb.create_sheet("README", 0)
        for r in README:
            ws.append(r)
        style_sheet(ws, [26, 52, 30, 70], ORANGE)
        ws.sheet_properties.tabColor = CHARCOAL
        added.append("README")

    if not added:
        print("Input tabs: all present, workbook not changed")
        return
    import shutil
    backup = path.with_name(path.stem + " (backup before input tabs)" + path.suffix)
    if not backup.exists():
        shutil.copy2(path, backup)     # keep the original once, in case Excel features need restoring
    try:
        wb.save(path)
    except PermissionError:
        sys.exit(f"Could not save {path.name} - close it in Excel and run again.")
    print("Input tabs added:", ", ".join(added), f"(original kept as '{backup.name}')")


if __name__ == "__main__":
    main()
