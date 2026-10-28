"""
ESG Territory Hours Dashboard builder
=====================================
Reads the hours workbook and writes one self-contained HTML page for partner presentations.

Usage
-----
1. Change EXCEL_PATH below (or pass --excel "C:\\path\\to\\file.xlsx").
2. Run:   python build_dashboard.py
          python build_dashboard.py --excel "D:\\Data\\Hours.xlsx" --out "D:\\Data\\ESG_Dashboard.html"
3. Double-click the HTML file. It works offline if chart.umd.min.js sits next to this script.

Requires: pandas, openpyxl
"""
import argparse
import base64
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------- CONFIG ---
HERE = Path(__file__).resolve().parent
EXCEL_PATH = HERE / "ESG_Hours_Sample.xlsx"      # <-- change this to your file
OUTPUT_PATH = HERE / "ESG_Hours_Dashboard.html"  # <-- where the page is written

SHEET_RAW = "Project Hours Raw Data"
SHEET_STAFF = "Staff Master"
SHEET_MODEL = "Global Model FY27 CPP"
SHEET_CALLOUTS = "Callouts"
SHEET_ASSUMPTIONS = "Assumptions"
SHEET_WRN_CAT = "WRN Category"
# Optional input tabs (static content you maintain). Missing tabs fall back to the defaults in the scripts.
SHEET_TAG_MAP = "Tag Mapping"
SHEET_TEAM = "Team"
SHEET_SERVICES = "Services"
SHEET_THEMES = "ET Themes"
SHEET_STORY = "Story"
SHEET_PRESETS = "Scenario Presets"
# Used only if the workbook has no 'Global Model FY27 CPP' sheet.
# Illustrative values only - put your real rate card in the workbook.
DEFAULT_MODEL = {"SM": ("Senior Manager", 95, 96000), "M": ("Manager", 65, 64000),
                 "SA": ("Senior Associate", 48, 43000), "A": ("Associate", 30, 28500)}
#   # optional: Work Request No. -> ESG/ET, fills blanks in the raw ESG/ET column
CAT_COLUMN = "ESG/ET"

# Display order; any other territory found in the data is appended after these.
TERRITORY_ORDER = ["Canada", "USA", "Australia", "UK", "Japan"]
# Spelling variants in the raw data -> standard name.
TERRITORY_ALIASES = {
    "us": "USA", "usa": "USA", "united states": "USA",
    "uk": "UK", "united kingdom": "UK",
    "au": "Australia", "aus": "Australia", "australia": "Australia",
    "ca": "Canada", "canada": "Canada",
    "jp": "Japan", "japan": "Japan",
}
# Keyword in 'Home Sub LOS 2' -> territory the person is currently tagged to.
# Tags such as 'Other Territory' or 'Deals' map to no territory.
TAG_KEYWORDS = {r"\bUS\b|\bUSA\b": "USA", r"\bUK\b": "UK", r"\bCA\b|Canada": "Canada",
                r"\bAU\b|Australia": "Australia", r"\bJP\b|Japan": "Japan"}

# Branding: put your official logo file next to this script (PNG, JPG or SVG) and set its name here.
LOGO_FILE = "logo.png"         # put your logo in engine/ (ignored by Git); blank wordmark if missing
LOGO_TEXT = ""                 # text shown when there is no logo file
# Text for the 'Team overview' and 'Energy Transition themes' tabs. Edit this JSON file to change them;
# delete it to hide both tabs.
CONTENT_FILE = "team_content.json"
REPORT_TITLE = "ESG Delivery: Territory Hours"
CLASSIFICATION = "Confidential | Internal use only"
# Sign-off shown at the end of both pages: "For any changes, additions or further analysis, please reach out to <name>."
# Shown once, at the end of the last dashboard tab and at the end of the story.
CONTACT_NAME = "Your Name"
CONTACT_EMAIL = ""          # optional - shown in brackets after the name, as a clickable email link
CONTACT_NAME_2 = ""         # optional second contact - leave blank to show one name only
CONTACT_EMAIL_2 = ""        # optional
CONTACT_TEAM = ""           # optional; e.g. a team name shown after the contacts (blank = none)
# The partner story page (built by build_story_v2.py). When that script is in the same folder,
# the dashboard header shows a "Back to partner story" link to it.
STORY_FILE = "ESG_US_Story_v2.html"

# Scenario presets shown as one-click buttons on the 'US tagging scenario' tab.
# threshold / window = None -> use the value from the Assumptions sheet.
# growth = expected change in US hours next year (0.10 = +10%).
# method: "hourly" = Global Model rate/hr; "best" = cheaper of hourly or a dedicated seat (Cost/Person).
SCENARIO_PRESETS = [
    {"name": "Conservative", "threshold": 0.75, "window": 12, "growth": 0.00, "method": "hourly"},
    {"name": "Base", "threshold": None, "window": None, "growth": 0.10, "method": "hourly"},
    {"name": "Growth", "threshold": 0.50, "window": 12, "growth": 0.30, "method": "best"},
]

DEFAULT_TAGGING_NOTE = ("Today no one is tagged to a territory: everyone sits under a non-territory tag such as "
                        "Other Territory. This analysis prepares the move to territory tagging.")

DEFAULT_ASSUMPTIONS = {"FY start month": 7, "Tagging threshold": 0.6,
                       "Scenario window (months)": 12, "Annual capacity hours": 1800,
                       "Training month": "Aug-24"}
# ---------------------------------------------------------------------------


def norm_territory(v):
    s = str(v).strip()
    return TERRITORY_ALIASES.get(s.lower(), s)


def to_month(v):
    """'Sep-24', a date, or 'Sep 2024' -> Period('2024-09')."""
    if pd.isna(v) or str(v).strip() == "":
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Period(v, "M")
    s = str(v).strip()
    for fmt in ("%b-%y", "%b-%Y", "%b %y", "%b %Y", "%Y-%m", "%m/%Y"):
        try:
            return pd.Period(datetime.strptime(s, fmt), "M")
        except ValueError:
            pass
    try:
        return pd.Period(pd.to_datetime(s, dayfirst=True), "M")
    except Exception:
        return None


def clean_str(v, default=""):
    return default if pd.isna(v) or str(v).strip() == "" else str(v).strip()


def load(excel_path):
    xl = pd.ExcelFile(excel_path)
    need = [SHEET_RAW]
    missing = [s for s in need if s not in xl.sheet_names]
    if missing:
        sys.exit(f"Missing sheet(s): {missing}. Found: {xl.sheet_names}")
    raw = xl.parse(SHEET_RAW)
    staff = xl.parse(SHEET_STAFF) if SHEET_STAFF in xl.sheet_names else pd.DataFrame()
    model = xl.parse(SHEET_MODEL) if SHEET_MODEL in xl.sheet_names else pd.DataFrame()
    callouts = xl.parse(SHEET_CALLOUTS) if SHEET_CALLOUTS in xl.sheet_names else pd.DataFrame()
    assum = dict(DEFAULT_ASSUMPTIONS)
    if SHEET_ASSUMPTIONS in xl.sheet_names:
        a = xl.parse(SHEET_ASSUMPTIONS)
        for _, r in a.iterrows():
            if clean_str(r.get("Parameter")):
                assum[clean_str(r["Parameter"])] = r["Value"]
    wrn_cat = xl.parse(SHEET_WRN_CAT, dtype=str) if SHEET_WRN_CAT in xl.sheet_names else pd.DataFrame()
    inputs = {name: xl.parse(name).rename(columns=lambda c: str(c).strip()).dropna(how="all")
              for name in (SHEET_TAG_MAP, SHEET_TEAM, SHEET_SERVICES, SHEET_THEMES, SHEET_STORY, SHEET_PRESETS)
              if name in xl.sheet_names}
    return raw, staff, model, callouts, assum, wrn_cat, inputs


def to_share(v, default=None):
    """60, '60%', 0.6 -> 0.6 ; blank -> default"""
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        return default
    s = str(v).strip().replace("%", "")
    try:
        x = float(s)
    except ValueError:
        return default
    return x / 100 if abs(x) > 1.5 or str(v).strip().endswith("%") else x


def read_inputs(inputs, assum):
    """Turn the optional input tabs into: tag map, page content, story text, scenario presets."""
    out = {"tagMap": {}, "content": None, "story": None, "presets": None}
    df = inputs.get(SHEET_TAG_MAP)
    if df is not None and "Home Sub LOS 2" in df.columns:
        for tag, terr in zip(df["Home Sub LOS 2"], df.get("Territory", [""] * len(df))):
            if clean_str(tag):
                out["tagMap"][clean_str(tag).lower()] = norm_territory(terr) if clean_str(terr) else ""

    team, svc, th = inputs.get(SHEET_TEAM), inputs.get(SHEET_SERVICES), inputs.get(SHEET_THEMES)
    if team is not None or svc is not None or th is not None:
        c = {"team": {"teamName": "", "summary": "", "members": "", "leaders": [], "keyContacts": [], "currentServices": [],
                      "futureServices": [], "frameworks": [], "databases": [], "aiTools": []},
             "themes": {"title": clean_str(assum.get("Energy Transition page title"), "Our Work Across Key Energy Transition Themes"), "items": []}}
        t = c["team"]
        if team is not None:
            for _, r in team.iterrows():
                typ = clean_str(r.get("Type")).lower()
                nm, rl, tr, nt = (clean_str(r.get(k)) for k in ("Name", "Role", "Territory", "Note"))
                if typ == "team name": t["teamName"] = nm
                elif typ == "summary": t["summary"] = nt or nm
                elif typ == "members": t["members"] = nm or nt
                elif typ == "leader": t["leaders"].append({"name": nm, "role": rl, "note": nt})
                elif typ == "key contact": t["keyContacts"].append({"territory": norm_territory(tr), "name": nm, "role": rl})
        if svc is not None:
            keymap = {"current service": "currentServices", "future service": "futureServices", "esg framework": "frameworks",
                      "database": "databases", "digital & ai tool": "aiTools", "ai tool": "aiTools"}
            for cat, item in zip(svc.get("Category", []), svc.get("Item", [])):
                k = keymap.get(clean_str(cat).lower())
                if k and clean_str(item): t[k].append(clean_str(item))
        if th is not None:
            for no, g in th.groupby(th["Theme No"].ffill(), sort=True):
                first = g.iloc[0]
                c["themes"]["items"].append({"title": clean_str(first.get("Theme Title")), "tagline": clean_str(first.get("Tagline")),
                    "points": [[clean_str(a), clean_str(b)] for a, b in zip(g.get("Work Area", []), g.get("Description", [])) if clean_str(a)]})
        out["content"] = c

    st = inputs.get(SHEET_STORY)
    if st is not None and "Section" in st.columns:
        story = {}
        lists = {"ask": "ask", "win for us": "winUS", "win for ac": "winAC"}
        for _, r in st.iterrows():
            sec = clean_str(r.get("Section")).lower(); tx = clean_str(r.get("Text"))
            if sec in ("title", "subtitle", "meeting", "presenter"): story[sec] = tx
            elif sec in lists: story.setdefault(lists[sec], []).append(tx)
            elif sec == "next step": story.setdefault("steps", []).append([tx, clean_str(r.get("Heading")), clean_str(r.get("Detail"))])
        out["story"] = story

    pr = inputs.get(SHEET_PRESETS)
    if pr is not None and "Name" in pr.columns:
        out["presets"] = [{"name": clean_str(r.get("Name")),
                           "threshold": to_share(r.get("Tagging threshold")),
                           "window": int(r["Look-back months"]) if pd.notna(r.get("Look-back months")) else None,
                           "growth": to_share(r.get("US volume change next year"), 0.0),
                           "method": "best" if "seat" in clean_str(r.get("Costing method")).lower() else "hourly"}
                          for _, r in pr.iterrows() if clean_str(r.get("Name"))]
    return out


def build_payload(raw, staff, model, callouts, assum, wrn_cat, inputs, source_name):
    inp = read_inputs(inputs or {}, assum)
    raw = raw.rename(columns=lambda c: str(c).strip())
    for col in ["Territory Name", "Staff Name", "Hours"]:
        if col not in raw.columns:
            sys.exit(f"Column '{col}' not found in '{SHEET_RAW}'.")

    # Month: 'Source Sheet' is the reporting month ("consider as date"); fall back to Date.
    month = raw["Source Sheet"].map(to_month) if "Source Sheet" in raw.columns else pd.Series([None] * len(raw))
    if "Date" in raw.columns:
        month = month.where(month.notna(), raw["Date"].map(to_month))
    raw["_month"] = month
    raw["Hours"] = pd.to_numeric(raw["Hours"], errors="coerce")
    bad = raw["_month"].isna() | raw["Hours"].isna() | raw["Staff Name"].isna() | raw["Territory Name"].isna()
    dropped = int(bad.sum())
    raw = raw[~bad].copy()
    raw["_terr"] = raw["Territory Name"].map(norm_territory)
    raw["_staff"] = raw["Staff Name"].map(clean_str)

    # Staff master
    smap = {}
    rate_cols = {}
    if not staff.empty:
        staff = staff.rename(columns=lambda c: str(c).strip())
        for c in staff.columns:
            m = re.search(r"(?:Sales Price|Rate)\D*(\d{4})", c, re.I)
            if m:
                rate_cols[int(m.group(1))] = c
        plain = [c for c in staff.columns if re.fullmatch(r"(sales price|rate)\b.*", c, re.I) and not re.search(r"\d{4}", c)]
        if not rate_cols and plain:   # one 'Sales Price' column: use it for every year in the data
            rate_cols = {y: plain[0] for y in sorted({p.year for p in raw["_month"]})}
        for _, r in staff.iterrows():
            name = clean_str(r.get("Employee"))
            if not name:
                continue
            tag_txt = clean_str(r.get("Home Sub LOS 2"))
            tag_terr = inp["tagMap"].get(tag_txt.lower())
            for pat, terr in ([] if tag_terr is not None else TAG_KEYWORDS.items()):
                if re.search(pat, tag_txt, re.I):
                    tag_terr = terr
                    break
            tag_terr = tag_terr or ""
            smap[name] = {
                "name": name,
                "level": clean_str(r.get("Level")).upper(),
                "team": clean_str(r.get("Team Name")),
                "costCenter": clean_str(r.get("Home Cost Center")),
                "tag": tag_txt,
                "tagTerritory": tag_terr,
                "rates": {str(y): float(r[c]) for y, c in rate_cols.items() if pd.notna(r.get(c))},
            }

    # Fill level / team from the raw data where the master lacks them, and vice versa.
    unknown_staff = sorted(set(raw["_staff"]) - set(smap))
    for name in unknown_staff:
        smap[name] = {"name": name, "level": "", "team": "", "costCenter": "", "tag": "", "tagTerritory": "", "rates": {}}
    if "Management Level" in raw.columns:
        for name, lv in raw.dropna(subset=["Management Level"]).groupby("_staff")["Management Level"].agg(lambda s: s.mode().iat[0]).items():
            if not smap[name]["level"]:
                smap[name]["level"] = str(lv).strip().upper()
    if "Team Name" in raw.columns:
        raw["_team"] = raw["Team Name"].map(clean_str)
        for name, tm in raw[raw["_team"] != ""].groupby("_staff")["_team"].agg(lambda s: s.mode().iat[0]).items():
            if not smap[name]["team"]:
                smap[name]["team"] = tm
        raw.loc[raw["_team"] == "", "_team"] = raw["_staff"].map(lambda n: smap[n]["team"])
    else:
        raw["_team"] = raw["_staff"].map(lambda n: smap[n]["team"])
    raw["_team"] = raw["_team"].replace("", "Unassigned")

    # Projects
    pcols = {"name": "Project Name", "client": "Client Name/WBS Name", "wbs": "Client code/WBS code",
             "wrn": "Work Request No.", "job": "Job Name"}
    for k, c in pcols.items():
        raw["_p_" + k] = raw[c].map(clean_str) if c in raw.columns else ""
    raw.loc[raw["_p_name"] == "", "_p_name"] = raw["_p_job"].where(raw["_p_job"] != "", "Unspecified")
    raw["_pkey"] = raw["_terr"] + "|" + raw["_p_name"] + "|" + raw["_p_client"]
    raw["_flag"] = raw["Project Match Flag"].map(lambda v: clean_str(v, "Not flagged")) if "Project Match Flag" in raw.columns else "Not flagged"

    # ESG / ET category: the row's own ESG/ET value, else the WRN Category sheet, else 'Unclassified'
    norm_cat = lambda v: {"ENERGY TRANSITION": "ET"}.get(clean_str(v).upper(), clean_str(v).upper())
    raw["_cat"] = raw[CAT_COLUMN].map(norm_cat) if CAT_COLUMN in raw.columns else ""
    if not wrn_cat.empty and {"Work Request No.", CAT_COLUMN} <= set(wrn_cat.columns):
        lookup = {clean_str(w): norm_cat(c) for w, c in zip(wrn_cat["Work Request No."], wrn_cat[CAT_COLUMN]) if clean_str(c)}
        blank = raw["_cat"] == ""
        raw.loc[blank, "_cat"] = raw.loc[blank, "_p_wrn"].map(lambda w: lookup.get(w, ""))
    raw.loc[raw["_cat"] == "", "_cat"] = "Unclassified"

    # Index tables
    months = pd.period_range(raw["_month"].min(), raw["_month"].max(), freq="M")
    m_idx = {p: i for i, p in enumerate(months)}
    terrs = [t for t in TERRITORY_ORDER if t in set(raw["_terr"])] + sorted(set(raw["_terr"]) - set(TERRITORY_ORDER))
    t_idx = {t: i for i, t in enumerate(terrs)}
    staff_names = sorted(smap)
    s_idx = {n: i for i, n in enumerate(staff_names)}
    teams = sorted(set(raw["_team"]))
    tm_idx = {t: i for i, t in enumerate(teams)}
    flags = sorted(set(raw["_flag"]))
    f_idx = {f: i for i, f in enumerate(flags)}
    pdf = raw.drop_duplicates("_pkey")
    projects = [{"territory": r["_terr"], "name": r["_p_name"], "client": r["_p_client"], "wbs": r["_p_wbs"],
                 "wrn": r["_p_wrn"], "job": r["_p_job"]} for _, r in pdf.iterrows()]
    p_idx = {k: i for i, k in enumerate(pdf["_pkey"])}
    pcat = raw.groupby("_pkey")["_cat"].agg(lambda x: x.mode().iat[0])
    for k, i in p_idx.items():
        projects[i]["cat"] = pcat[k]
    cats = [c for c in ["ESG", "ET"] if c in set(raw["_cat"])] + sorted(set(raw["_cat"]) - {"ESG", "ET"})
    c_idx = {c: i for i, c in enumerate(cats)}

    rows = [[m_idx[m], t_idx[t], s_idx[s], p_idx[p], tm_idx[tm], round(float(h), 2), f_idx[f], c_idx[c]]
            for m, t, s, p, tm, h, f, c in zip(raw["_month"], raw["_terr"], raw["_staff"], raw["_pkey"],
                                               raw["_team"], raw["Hours"], raw["_flag"], raw["_cat"])]

    model_rates = {}
    if not model.empty:
        model = model.rename(columns=lambda c: str(c).strip())
        for _, r in model.iterrows():
            lv = clean_str(r.get("Level")).upper()
            if lv:
                model_rates[lv] = {"name": clean_str(r.get("Level Name"), lv),
                                   "rate": float(r.get("Rate/hr", 0) or 0),
                                   "cpp": float(r.get("Cost/Person", 0) or 0)}

    model_default = not model_rates
    if model_default:
        model_rates = {lv: {"name": n, "rate": float(r), "cpp": float(c)} for lv, (n, r, c) in DEFAULT_MODEL.items()}

    calls = []
    if not callouts.empty:
        for _, r in callouts.iterrows():
            p = to_month(r.get("Month"))
            if p is not None:
                calls.append({"month": str(p), "territory": clean_str(r.get("Territory")), "event": clean_str(r.get("Event"))})

    tm = to_month(assum.get("Training month"))
    meta = {
        "source": source_name,
        "generated": datetime.now().strftime("%d %b %Y %H:%M"),
        "fyStart": int(assum.get("FY start month", 7)),
        "threshold": float(assum.get("Tagging threshold", 0.6)),
        "window": int(assum.get("Scenario window (months)", 12)),
        "capacity": float(assum.get("Annual capacity hours", 1800)),
        "trainingMonth": str(tm) if tm is not None else None,
        "taggingNote": clean_str(assum.get("Current tagging"), DEFAULT_TAGGING_NOTE),
        "rateYears": [str(y) for y in sorted(set(rate_cols) & {p.year for p in months})][-2:],
        "title": clean_str(assum.get("Report title"), REPORT_TITLE), "classification": clean_str(assum.get("Classification"), CLASSIFICATION),
        "contact": {"people": [{"name": n, "email": e} for n, e in ((CONTACT_NAME, CONTACT_EMAIL), (CONTACT_NAME_2, CONTACT_EMAIL_2)) if n or e],
                    "team": CONTACT_TEAM},
        "storyLink": STORY_FILE if (HERE / "build_story_v2.py").exists() or (HERE / STORY_FILE).exists() else None,
    }
    meta["presets"] = [{**pr, "threshold": meta["threshold"] if pr["threshold"] is None else pr["threshold"],
                        "window": meta["window"] if pr["window"] is None else pr["window"]} for pr in (inp["presets"] or SCENARIO_PRESETS)]
    if not any(p["name"].lower() == "base" for p in meta["presets"]) and meta["presets"]:
        meta["presets"][0]["name"] = meta["presets"][0]["name"] or "Base"
    payload = {"meta": meta, "months": [str(p) for p in months], "territories": terrs,
               "staff": [smap[n] for n in staff_names], "teams": teams, "flags": flags, "categories": cats,
               "projects": projects, "rows": rows, "model": model_rates, "callouts": calls,
               "content": inp["content"], "storyInput": inp["story"]}

    report = {
        "rows used": len(rows), "rows dropped (no month/hours/staff/territory)": dropped,
        "total hours": round(raw["Hours"].sum(), 1),
        "hours by ESG/ET": ", ".join(f"{c} {h:,.0f}" for c, h in raw.groupby("_cat")["Hours"].sum().items()),
        "months": f"{months[0]} to {months[-1]}", "territories": ", ".join(terrs),
        "staff not in Staff Master": ", ".join(unknown_staff) or "none",
        "staff without a level": ", ".join(n for n in staff_names if not smap[n]["level"]) or "none",
        "sales prices": (", ".join(f"{y}: '{c}'" for y, c in sorted(rate_cols.items())) if rate_cols
                         else "MISSING - add 'Sales Price' (or 'Sales Price YYYY') to Staff Master; cost and saving figures need it"),
        "rate card": "Global Model sheet" if not model_default else "sheet missing - using built-in FY27 defaults",
        "input tabs found": ", ".join(inputs) if inputs else "none (using defaults in the scripts)",
        "levels missing from Global Model": ", ".join(sorted({smap[n]["level"] for n in staff_names if smap[n]["level"] and smap[n]["level"] not in model_rates})) or "none",
    }
    return payload, report


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", default=str(EXCEL_PATH), help="Path to the hours workbook")
    ap.add_argument("--out", default=str(OUTPUT_PATH), help="Path of the HTML file to write")
    args = ap.parse_args()

    excel = Path(args.excel).expanduser()
    if not excel.exists():
        sys.exit(f"Excel file not found: {excel}")
    payload, report = build_payload(*load(excel), source_name=excel.name)

    content_path = HERE / CONTENT_FILE
    if payload["content"] is None and content_path.exists():   # Excel tabs win; JSON is the fallback
        payload["content"] = json.loads(content_path.read_text(encoding="utf-8"))

    chart_js = HERE / "chart.umd.min.js"
    if chart_js.exists():
        chart_tag = "<script>" + chart_js.read_text(encoding="utf-8") + "</script>"
    else:
        print("note: chart.umd.min.js not found next to the script - page will load Chart.js from the internet")
        chart_tag = '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>'

    logo = HERE / LOGO_FILE
    if logo.exists():
        mime = {".svg": "image/svg+xml", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}.get(logo.suffix.lower(), "image/png")
        logo_html = f'<img class="logo" alt="logo" src="data:{mime};base64,{base64.b64encode(logo.read_bytes()).decode()}">'
    else:
        logo_html = f'<span class="wordmark">{LOGO_TEXT}</span>'
        print(f"note: {LOGO_FILE} not found next to the script - showing a text placeholder instead of the logo")

    html = (TEMPLATE.replace("/*__CHARTJS__*/", chart_tag).replace("<!--__LOGO__-->", logo_html)
                    .replace("/*__DATA__*/null", json.dumps(payload, separators=(",", ":"))))
    out = Path(args.out).expanduser()
    out.write_text(html, encoding="utf-8")

    print("Data check")
    for k, v in report.items():
        print(f"  {k:<45} {v}")
    print(f"\nDashboard written to: {out}")


TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ESG Territory Hours</title>
<style>
:root{
  color-scheme:light;
  /* brand palette */
  --orange:#D04A02; --tangerine:#EB8C00; --yellow:#FFB600; --rose:#DB536A; --red:#E0301E; --burgundy:#A32020; --charcoal:#2D2D2D;
  --bg:#F3F3F3; --surface:#FFFFFF; --surface-2:#F5F5F5; --border:#DEDEDE;
  --text:#2D2D2D; --text-2:#5F5F5F; --text-3:#7D7D7D; --grid:#EBEBEB;
  --accent:var(--orange); --good:#2E7D32; --bad:#C0392B;
  /* territory colours (fixed per territory) + fallbacks */
  --c-canada:#464646; --c-usa:#D04A02; --c-australia:#FFB600; --c-uk:#DB536A; --c-japan:#A32020;
  --t1:#464646; --t2:#D04A02; --t3:#FFB600; --t4:#DB536A; --t5:#A32020; --t6:#EB8C00; --t7:#7D7D7D; --t8:#E0301E;
  --neutral:#BDBDBD; --seq-lo:#BDBDBD; --seq-hi:#D04A02;
  --pos-bg:#F6C3A6; --neg-bg:#CFCFCF;
  --head:Georgia,"Times New Roman",serif;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:14px/1.45 Arial,Helvetica,sans-serif}
.brandbar{background:var(--surface);border-top:6px solid var(--orange);border-bottom:1px solid var(--border)}
.brandbar .in{max-width:1400px;margin:0 auto;padding:14px 28px;display:flex;align-items:center;gap:18px}
.logo{height:40px;width:auto}
.wordmark{font-family:var(--head);font-weight:700;font-size:28px;color:var(--charcoal);letter-spacing:-.02em}
.brandbar .div{width:1px;height:34px;background:var(--border)}
.brandbar .cls{margin-left:auto;font-size:12px;color:var(--text-3)}
.backlink{display:inline-flex;align-items:center;gap:6px;background:var(--orange);color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:7px 14px;border-radius:4px}
.backlink:hover{background:#B23F02}
header{padding:18px 28px 0;max-width:1400px;margin:0 auto}
h1{font-family:var(--head);font-size:24px;margin:0;font-weight:700;letter-spacing:-.01em}
.card h2,.kpi .v{font-family:var(--head)}
.sub{color:var(--text-2);font-size:13px}
nav{display:flex;gap:4px;flex-wrap:wrap;margin-top:16px;border-bottom:1px solid var(--border)}
nav button{background:none;border:0;border-bottom:2px solid transparent;padding:10px 12px;white-space:nowrap;font:inherit;color:var(--text-2);cursor:pointer}
nav button.on{color:var(--text);border-bottom-color:var(--accent);font-weight:600}
main{max-width:1400px;margin:0 auto;padding:20px 28px 48px}
section{display:none}
section.on{display:block;animation:tabIn .45s cubic-bezier(.2,.7,.2,1)}
@keyframes tabIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
.kpi,.tile,.card{transition:box-shadow .25s ease,transform .25s ease}
.kpi:hover{box-shadow:0 6px 16px rgba(0,0,0,.06);transform:translateY(-2px)}
@media (prefers-reduced-motion:reduce){section.on{animation:none} .kpi:hover{transform:none}}
.grid{display:grid;gap:16px}
.g2{grid-template-columns:1fr 1fr}
.g3{grid-template-columns:2fr 1fr}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:12px;margin-bottom:16px}
.kpi{background:var(--surface);border:1px solid var(--border);border-left:4px solid var(--orange);border-radius:4px;padding:14px 16px}
.kpi .l{color:var(--text-2);font-size:12px}
.kpi .v{font-size:26px;font-weight:650;margin-top:2px;font-variant-numeric:tabular-nums}
.kpi .d{font-size:12px;color:var(--text-2);margin-top:2px}
.up{color:var(--good)} .down{color:var(--bad)}
.card{background:var(--surface);border:1px solid var(--border);border-radius:4px;padding:16px 18px;min-width:0}
.card h2{font-size:17px;margin:0 0 2px;font-weight:700}
.card .note{color:var(--text-2);font-size:12px;margin-bottom:10px}
.chart{position:relative;height:300px}
.chart.tall{height:380px}
.ctl{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
.seg{display:inline-flex;border:1px solid var(--border);border-radius:4px;overflow:hidden}
.seg button{background:var(--surface);border:0;padding:6px 12px;font:inherit;font-size:13px;color:var(--text-2);cursor:pointer}
.seg button+button{border-left:1px solid var(--border)}
.seg button.on{background:var(--accent);color:#fff}
.ctl .lab{font-size:12px;color:var(--text-2);margin-right:2px}
table{width:100%;border-collapse:collapse;font-size:13px;font-variant-numeric:tabular-nums}
th,td{padding:7px 8px;border-bottom:1px solid var(--border);text-align:left}
td.n,th.n,td:first-child,td:last-child{white-space:nowrap}
th{white-space:nowrap;font-weight:700;color:#fff;font-size:12px;background:var(--charcoal);position:sticky;top:0}
td.n,th.n{text-align:right}
tr.tot td{font-weight:650;border-top:2px solid var(--border)}
.tw{overflow-x:auto;max-height:420px;overflow-y:auto}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;vertical-align:-1px}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--text-2);margin:4px 0 10px}
.mix{display:flex;height:12px;width:150px;border-radius:3px;overflow:hidden;gap:1px;background:var(--surface)}
.pill{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;border:1px solid var(--border)}
.pill.act{border-color:var(--accent);color:var(--accent);font-weight:600}
.callouts{margin:0;padding:0;list-style:none}
.callouts li{padding:8px 0;border-bottom:1px solid var(--border);display:flex;gap:12px}
.callouts li b{min-width:58px;font-variant-numeric:tabular-nums}
.story{font-size:14px;color:var(--text);max-width:900px}
.story p{margin:0 0 8px}
.assume{background:var(--surface-2);border-radius:8px;padding:10px 14px;font-size:12.5px;color:var(--text-2)}
.assume li{margin:2px 0}
footer{max-width:1400px;margin:0 auto;padding:0 28px 28px;color:var(--text-3);font-size:12px}
.calc{display:grid;grid-template-columns:320px 1fr;gap:28px;align-items:start;margin-top:8px}
.ctl2{display:grid;gap:16px}
.ctl2 label{font-size:13px;font-weight:700;display:flex;justify-content:space-between;margin-bottom:6px}
.ctl2 label span{font-family:var(--head);color:var(--orange);font-size:16px}
.bars{display:grid;gap:18px;grid-template-columns:minmax(0,1fr)}
.calc>*{min-width:0} .track>.seg-b{min-width:0}
.barrow .bl{display:flex;justify-content:space-between;font-size:13.5px;margin-bottom:6px}
.barrow .bl b{font-family:var(--head);font-size:17px}
.track{display:flex;height:48px;gap:2px}
.seg-b{display:flex;align-items:center;justify-content:center;color:#fff;font-size:12.5px;font-weight:700;overflow:hidden;white-space:nowrap;padding:0 6px;border-radius:2px;transition:width .35s}
.seg-b.today{background:var(--charcoal)} .seg-b.mgr{background:var(--orange)} .seg-b.assoc{background:var(--tangerine)}
.seg-b.save{background:repeating-linear-gradient(45deg,#E9F3EA,#E9F3EA 6px,#D5E8D7 6px,#D5E8D7 12px);color:var(--good);border:1px dashed var(--good)}
.verdict{font-family:var(--head);font-size:20px;line-height:1.35;margin:4px 0 0}
.verdict b{color:var(--orange)}
@media (max-width:900px){.calc{grid-template-columns:1fr}}
.contactline{margin:6px 0 0;color:var(--text-3);font-size:12px}
.contactline a{color:inherit}
.legal{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;border-top:1px solid var(--border);padding-top:12px}
.legal b{color:var(--text)}
.ctlrow{display:grid;grid-template-columns:170px 1fr 64px;align-items:center;column-gap:16px;row-gap:4px;padding:12px 0;border-bottom:1px solid var(--grid)}
.ctlrow:last-child{border-bottom:0}
.ctlrow .lab{font-size:13px;font-weight:700;color:var(--text)}
.ctlrow .val{font-family:var(--head);font-size:17px;font-weight:700;text-align:right;font-variant-numeric:tabular-nums}
.ctlrow .hint{grid-column:2 / 4;font-size:12px;color:var(--text-3)}
.seg.full{display:flex;width:100%}
.seg.full button{flex:1;white-space:nowrap;padding:7px 10px}
.money{display:inline-flex;align-items:center;border:1px solid var(--border);border-radius:4px;background:var(--surface)}
.money span{padding:0 0 0 8px;color:var(--text-3)}
.money input.num{border:0;width:84px}
.money:focus-within{outline:2px solid var(--orange)}
.hero{border-top:4px solid var(--orange);margin-bottom:16px}
.eyebrow{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--text-3)}
.headline{font-family:var(--head);font-size:24px;line-height:1.3;margin:6px 0 10px;color:var(--text)}
.headline b{color:var(--orange)}
.points{margin:0;padding-left:18px;color:var(--text-2);display:grid;gap:4px}
.tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.tile{background:var(--surface);border:1px solid var(--border);border-radius:4px;padding:14px 16px;text-align:left;font:inherit;color:inherit;cursor:pointer;position:relative}
.tile:hover{border-color:var(--text-3)}
.tile.on{border:2px solid var(--orange);padding:13px 15px;background:#FFF8F4}
.tile.on::after{content:"On screen";position:absolute;top:10px;right:12px;font-size:11px;background:var(--orange);color:#fff;padding:1px 7px;border-radius:3px}
.tile.ghost{cursor:default;border-style:dashed;color:var(--text-3)}
.tile .tn{font-family:var(--head);font-weight:700;font-size:17px}
.tile .td{font-size:12px;color:var(--text-3);margin:2px 0 10px;min-height:32px}
.tile .tv{font-family:var(--head);font-size:24px;font-weight:700}
.tile .tl{font-size:12px;color:var(--text-2)}
.sub2{font-size:12px;color:var(--text-3);font-weight:400}
/* ---- team overview */
.band{background:var(--orange);color:#fff;font-weight:700;font-size:17px;padding:8px 14px;border-radius:2px;margin-bottom:12px}
.band.grey{background:#D9D9D9;color:var(--text)}
.team-grid{display:grid;grid-template-columns:minmax(280px,1fr) 3fr;gap:16px;align-items:start}
.meet{padding:0 0 14px}
.meet .band{margin:0 0 12px;border-radius:4px 4px 0 0}
.meet .band.grey{border-radius:0;margin-top:14px}
.meet blockquote{margin:0 18px 14px;padding:0 0 14px 44px;position:relative;color:var(--text-2);font-size:14.5px;border-bottom:1px solid var(--border)}
.meet blockquote::before{content:"\201C";position:absolute;left:-4px;top:-18px;font:700 64px/1 var(--head);color:var(--orange)}
.ladder{list-style:none;margin:0 18px;padding:0 0 0 18px;border-left:2px solid var(--orange);display:grid;gap:12px}
.ladder li{position:relative}
.ladder li::before{content:"";position:absolute;left:-24px;top:6px;width:10px;height:10px;border-radius:50%;background:var(--orange)}
.ladder .nm{color:var(--orange);font-weight:700;font-size:15px}
.ladder .rl{font-weight:700;font-size:13px}
.ladder .nt{font-size:12.5px;color:var(--text-2)}
.members{display:flex;align-items:center;gap:10px;margin:14px 18px 0;font-family:var(--head);font-size:20px;font-weight:700;color:var(--orange)}
.members .mi{display:inline-grid;place-items:center;width:40px;height:40px;border:1.5px solid var(--orange);border-radius:4px}
.contacts{list-style:none;margin:0 18px;padding:0;display:grid;gap:8px}
.contacts li{display:grid;grid-template-columns:14px 82px 1fr;align-items:baseline;gap:6px;font-size:13.5px}
.contacts .tt{font-weight:700;color:var(--orange)}
.contacts .rr{color:var(--text-3);font-size:12px}
.svc-grid{display:grid;grid-template-columns:2fr 1fr;gap:16px}
.tool-grid{display:grid;grid-template-columns:1fr 1.3fr 1.3fr;gap:16px;margin-top:16px}
.scope h3{font-size:15px;margin:0 0 10px;padding-bottom:8px;border-bottom:2px solid var(--orange);font-weight:700}
.future{background:#FFF8F4}
.ticks{list-style:none;margin:0;padding:0;display:grid;gap:7px;font-size:14px}
.ticks.two{grid-template-columns:1fr 1fr;column-gap:24px}
.ticks li{padding-left:18px;position:relative}
.ticks li::before{content:"";position:absolute;left:2px;top:7px;width:7px;height:7px;background:var(--orange);transform:rotate(45deg)}
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chips span{border:1px solid var(--border);border-radius:3px;padding:5px 10px;font-size:13px;background:var(--surface-2)}
.chips.stack{flex-direction:column}
.chips.stack span{text-align:center;padding:9px 10px}
/* ---- energy transition themes */
.page-h{font-family:var(--head);font-size:26px;font-weight:700;margin:4px 0 16px}
.themes-grid{display:grid;grid-template-columns:1fr 340px 1fr;grid-template-areas:"t1 d t3" "t2 d t4";gap:16px 24px;align-items:start}
#th1{grid-area:t1} #th2{grid-area:t2} #th3{grid-area:t3} #th4{grid-area:t4}
.diamond-wrap{grid-area:d;align-self:center;display:grid;place-items:center;height:340px}
.diamond{position:relative;width:240px;height:240px}
.diamond::before{content:"";position:absolute;inset:0;transform:rotate(45deg);background:linear-gradient(135deg,var(--tangerine),var(--orange));border-radius:6px}
.dcore{position:absolute;inset:62px;background:#FBE3D3;display:grid;place-items:center;text-align:center;font-family:var(--head);font-size:16px;line-height:1.25;color:var(--charcoal);padding:4px}
.dn{position:absolute;font:700 26px/1 Arial,Helvetica,sans-serif;color:var(--charcoal)}
.n1{top:-18px;left:50%;transform:translateX(-50%)} .n4{bottom:-18px;left:50%;transform:translateX(-50%)}
.n2{left:-26px;top:50%;transform:translateY(-50%)} .n3{right:-26px;top:50%;transform:translateY(-50%)}
.theme{background:var(--surface);border:1px solid var(--border);border-top:4px solid var(--orange);border-radius:4px;padding:16px 18px}
.theme .num{display:inline-block;font:700 13px/1 Arial,Helvetica,sans-serif;color:#fff;background:var(--orange);padding:4px 7px;border-radius:3px;margin-right:8px;vertical-align:3px}
.theme h3{font-family:var(--head);font-size:18px;margin:0;display:inline}
.theme .tag{color:var(--orange);font-size:13px;margin:6px 0 10px}
.theme ul{margin:0;padding-left:18px;display:grid;gap:6px;font-size:13.5px;color:var(--text-2)}
.theme ul b{color:var(--text)}
@media (max-width:1200px){.team-grid{grid-template-columns:1fr} .tool-grid{grid-template-columns:1fr 1fr}
  .themes-grid{grid-template-columns:1fr 1fr;grid-template-areas:"t1 t3" "t2 t4"} .diamond-wrap{display:none}}
@media (max-width:760px){.svc-grid,.tool-grid,.ticks.two{grid-template-columns:1fr} .themes-grid{grid-template-columns:1fr;grid-template-areas:"t1" "t2" "t3" "t4"}}
.pill{white-space:nowrap}
.nowrap td{white-space:nowrap}
input[type=range]{width:100%;accent-color:var(--orange)}
input.num{width:96px;padding:5px 7px;border:1px solid var(--border);border-radius:4px;text-align:right;font:inherit;font-variant-numeric:tabular-nums}
input.num:focus{outline:2px solid var(--orange);outline-offset:0}
input[type=checkbox]{accent-color:var(--orange);width:16px;height:16px}
.btn{background:var(--surface);border:1px solid var(--charcoal);color:var(--charcoal);border-radius:4px;padding:6px 14px;font:inherit;font-size:13px;cursor:pointer}
.btn:hover{background:var(--surface-2)}
.g-ctl{grid-template-columns:3fr 2fr}
@media (max-width:1200px){.g-ctl{grid-template-columns:1fr} .tiles{grid-template-columns:repeat(2,1fr)}}
td.hl,th.hl{background:#FDEDE4}
th.hl{background:var(--orange)}
.badge{display:inline-block;font-size:11px;padding:1px 7px;border-radius:3px;background:var(--orange);color:#fff;margin-left:6px;vertical-align:1px}
.pill.man{border-color:var(--tangerine);color:#9a5a00}
@media (max-width:900px){.g2,.g3,.g-ctl{grid-template-columns:1fr} .ctlrow{grid-template-columns:1fr 56px} .ctlrow .lab{grid-column:1 / 3} .ctlrow .hint{grid-column:1 / 3} header,main,footer{padding-left:16px;padding-right:16px}}
@media print{section{animation:none !important}}
@media print{nav,.ctl{display:none} section{display:block !important;page-break-after:always} body{background:#fff}}
</style>
/*__CHARTJS__*/
</head>
<body>
<div class="brandbar"><div class="in"><!--__LOGO__--><span class="div"></span><span class="sub" style="font-size:14px">Sustainability &amp; ESG Delivery</span><span class="cls" id="clsTop"></span><a class="backlink" id="backLink" href="#" hidden>&#8592; Back to partner story</a></div></div>
<header>
  <h1 id="title">ESG Delivery: Territory Hours</h1>
  <div class="sub" id="subtitle"></div>
  <div class="ctl" id="catRow" style="margin:12px 0 0"><span class="lab">Work type</span><div class="seg" id="catPick"></div></div>
  <nav id="tabs">
    <button data-tab="team" class="ct">Team</button>
    <button data-tab="themes" class="ct">Energy Transition</button>
    <button data-tab="overview" class="on">Overview</button>
    <button data-tab="territory">Territory</button>
    <button data-tab="usnonus">US vs Non-US</button>
    <button data-tab="growth">Growth</button>
    <button data-tab="rates">Rate analysis</button>
    <button data-tab="tagging">US tagging scenarios</button>
  </nav>
</header>
<main>
<!-- TEAM OVERVIEW -->
<section id="team">
  <div class="team-grid">
    <div class="card meet">
      <div class="band">Meet the team</div>
      <blockquote id="tmSummary"></blockquote>
      <ol class="ladder" id="tmLeaders"></ol>
      <div class="members"><span class="mi"><svg viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><circle cx="8" cy="8" r="3" fill="none" stroke="currentColor" stroke-width="1.6"/><circle cx="16" cy="8" r="3" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M2 20c0-3.3 2.7-6 6-6s6 2.7 6 6M10 20c0-3.3 2.7-6 6-6s6 2.7 6 6" fill="none" stroke="currentColor" stroke-width="1.6"/></svg></span><span id="tmMembers"></span></div>
      <div class="band grey">Key contacts</div>
      <ul class="contacts" id="tmContacts"></ul>
    </div>
    <div class="scope">
      <div class="band">Scope and capabilities</div>
      <div class="svc-grid">
        <div class="card"><h3>Current services</h3><ul class="ticks two" id="tmCurrent"></ul></div>
        <div class="card future"><h3>Future services</h3><ul class="ticks" id="tmFuture"></ul></div>
      </div>
      <div class="tool-grid">
        <div class="card"><h3>Integrated ESG frameworks</h3><div class="chips stack" id="tmFrameworks"></div></div>
        <div class="card"><h3>Database support</h3><div class="chips" id="tmDatabases"></div></div>
        <div class="card"><h3>Digital &amp; AI tools</h3><div class="chips" id="tmTools"></div></div>
      </div>
    </div>
  </div>
</section>

<!-- ENERGY TRANSITION THEMES -->
<section id="themes">
  <h2 class="page-h" id="thTitle"></h2>
  <div class="kpis" id="thKpis"></div>
  <div class="themes-grid">
    <div class="theme" id="th1"></div>
    <div class="diamond-wrap" aria-hidden="true"><div class="diamond"><span class="dn n1">01</span><span class="dn n2">02</span><span class="dn n3">03</span><span class="dn n4">04</span><div class="dcore">Key Energy<br>Transition<br>Themes</div></div></div>
    <div class="theme" id="th3"></div>
    <div class="theme" id="th2"></div>
    <div class="theme" id="th4"></div>
  </div>
</section>

<!-- OVERVIEW -->
<section id="overview" class="on">
  <div class="ctl"><span class="lab">Period</span><div class="seg" id="ovPeriod"><button data-v="FY" class="on">Financial year</button><button data-v="CY">Calendar year</button></div></div>
  <div class="kpis" id="ovKpis"></div>
  <div class="grid g3">
    <div class="card"><h2>Territory-wise hours</h2><div class="note" id="ovPeriodNote"></div><div class="chart"><canvas id="cPeriod"></canvas></div></div>
    <div class="card"><h2>Call-outs</h2><div class="note">Milestones marked on the monthly chart</div><ul class="callouts" id="ovCallouts"></ul></div>
  </div>
  <div class="card" style="margin-top:16px">
    <div class="ctl" style="justify-content:space-between;margin-bottom:4px"><div><h2>Month-on-month hours by territory</h2><div class="note">Dashed lines mark the call-outs</div></div>
    <div class="seg" id="ovRange"><button data-v="12">12 months</button><button data-v="24" class="on">24 months</button><button data-v="0">All</button></div></div>
    <div class="chart tall"><canvas id="cMonthly"></canvas></div>
  </div>
  <div class="grid g2" style="margin-top:16px">
    <div class="card"><h2 id="trainTitle">Training impact</h2><div class="note" id="trainNote"></div><div class="tw" id="trainTable"></div></div>
    <div class="card"><h2>Hours by territory and year</h2><div class="note">Table view of the chart above</div><div class="tw" id="ovPeriodTable"></div></div>
  </div>
</section>

<!-- TERRITORY -->
<section id="territory">
  <div class="ctl"><span class="lab">Territory</span><div class="seg" id="terrPick"></div></div>
  <div class="kpis" id="terrKpis"></div>
  <div class="grid g3">
    <div class="card"><h2>Monthly hours</h2><div class="note" id="terrMonthlyNote"></div><div class="chart"><canvas id="cTerrMonthly"></canvas></div></div>
    <div class="card"><h2>Work stream (team)</h2><div class="note" id="terrTeamNote"></div><div class="chart"><canvas id="cTerrTeam"></canvas></div></div>
  </div>
  <div class="grid g2" style="margin-top:16px">
    <div class="card"><h2>Projects</h2><div class="note">Sorted by hours in the last 12 months</div><div class="tw" id="terrProjects"></div></div>
    <div class="card"><h2>People</h2><div class="note">Last 12 months</div><div class="tw" id="terrPeople"></div></div>
  </div>
</section>

<!-- US VS NON-US -->
<section id="usnonus">
  <div class="kpis" id="usKpis"></div>
  <div class="card"><h2>US vs Non-US hours by month</h2><div class="note">Last 24 months</div><div class="chart"><canvas id="cUsMonthly"></canvas></div></div>
  <div class="grid g2" style="margin-top:16px">
    <div class="card"><h2>1. Territory-wise</h2><div class="note">Hours by financial year</div><div class="tw" id="usTerrTable"></div></div>
    <div class="card"><h2>2. Team-wise</h2><div class="note">Last 12 months</div><div class="chart" style="height:240px"><canvas id="cUsTeam"></canvas></div><div class="tw" id="usTeamTable" style="margin-top:10px"></div></div>
  </div>
</section>

<!-- GROWTH -->
<section id="growth">
  <div class="card"><h2>Year-on-year growth by territory</h2><div class="note" id="growthNote"></div><div class="tw" id="growthYoY"></div></div>
  <div class="grid g2" style="margin-top:16px">
    <div class="card"><h2>Trend: 3-month rolling average hours</h2><div class="note">Smooths out month-to-month noise</div><div class="chart"><canvas id="cRolling"></canvas></div></div>
    <div class="card"><h2>Month-on-month growth %</h2><div class="note">Last 12 months; orange shading = growth, grey = decline</div><div class="tw" id="growthMoM"></div></div>
  </div>
</section>

<!-- RATES -->
<section id="rates">
  <div class="kpis" id="rateKpis"></div>
  <div class="grid g2">
    <div class="card"><h2>Billed value by territory</h2><div class="note">Hours × each person's sales price for that year</div><div class="chart"><canvas id="cRateValue"></canvas></div></div>
    <div class="card"><h2>Effective rate by territory</h2><div class="note">Value ÷ hours, USD per hour</div><div class="tw" id="rateTerr"></div></div>
  </div>
  <div class="card" style="margin-top:16px"><h2>Rate per person</h2><div class="note">Sales price by year vs Global Model FY27 rate for the person's level</div><div class="tw" id="rateStaff" style="max-height:none"></div></div>
</section>

<!-- TAGGING -->
<section id="tagging">
  <div class="card hero">
    <div class="eyebrow">US tagging · <span id="scName"></span></div>
    <div class="headline" id="tagHeadline"></div>
    <ul class="points" id="tagPoints"></ul>
  </div>
  <div class="tiles" id="scTiles"></div>
  <div class="grid g-ctl">
    <div class="card">
      <div class="ctl" style="justify-content:space-between;margin-bottom:0"><div><h2>Adjust the assumptions</h2><div class="note" style="margin:0">Any change here creates "Your scenario"; Reset goes back to Base</div></div><button class="btn" id="scReset">Reset</button></div>
      <div>
        <div class="ctlrow"><span class="lab">Tagging threshold</span><input type="range" id="scThr" min="40" max="90" step="5" aria-label="Tagging threshold"><span class="val" id="scThrV"></span>
          <span class="hint">Tag a person to a territory when at least this share of their hours is in one territory</span></div>
        <div class="ctlrow"><span class="lab">Look-back period</span><div class="seg full" id="scWin"><button data-v="6">6 months</button><button data-v="12">12 months</button><button data-v="24">24 months</button></div><span class="val" id="scWinV"></span>
          <span class="hint">Hours used to judge where each person works; all figures are per year</span></div>
        <div class="ctlrow"><span class="lab">US volume next year</span><input type="range" id="scGrowth" min="-30" max="100" step="5" aria-label="US volume next year"><span class="val" id="scGrowthV"></span>
          <span class="hint">Expected change in US hours, used for the next-year saving</span></div>
        <div class="ctlrow"><span class="lab">Costing method</span><div class="seg full" id="scMethod"><button data-v="hourly">Hourly rate</button><button data-v="best">Hourly or seat, whichever is cheaper</button></div><span class="val" id="scMethodV"></span>
          <span class="hint">A dedicated seat costs the level's FY27 Cost/Person for a full year</span></div>
      </div>
    </div>
    <div class="card"><h2>Rate card</h2><div class="note">Global Model FY27, in USD. Changes apply to every scenario; Reset restores them</div><div class="tw" id="scRates"></div></div>
  </div>
  <div class="kpis" id="tagKpis" style="margin-top:16px"></div>
  <div class="grid g2">
    <div class="card"><h2>Scenario comparison</h2><div class="note">The highlighted column is the scenario on screen</div><div class="tw" id="scCompare"></div></div>
    <div class="card"><h2>Annual saving by scenario</h2><div class="note">On today's US volume and on next year's projected volume</div><div class="chart"><canvas id="cScCompare"></canvas></div></div>
  </div>
  <div class="card" style="margin-top:16px"><h2>Who moves to the US tag</h2><div class="note" id="tagNote"></div>
    <div class="legend" id="tagLegend"></div><div class="tw nowrap" id="tagTable" style="max-height:none"></div></div>
  <div class="card" style="margin-top:16px"><h2>Same budget, more capacity: one full-time US person</h2>
    <div class="note">Today's hourly cost for one person vs a dedicated US seat at FY27 Cost/Person, and how many extra Associates the difference pays for. Uses the rate card above.</div>
    <div class="calc">
      <div class="ctl2">
        <div><label>Level</label><div class="seg full" id="fcLevel"></div></div>
        <div><label>US hours per year (one person) <span id="fcHrsV"></span></label><input type="range" id="fcHrs" min="400" max="1800" step="50" aria-label="US hours per year"></div>
        <div><label>Today's sales price per hour</label><span class="money"><span>$</span><input class="num" type="number" id="fcRate" min="0" step="1" aria-label="Today's sales price per hour"></span><div class="note" style="margin:4px 0 0">Pre-filled with this level's average price in Staff Master</div></div>
        <div class="note" id="fcNote"></div>
      </div>
      <div class="bars">
        <div class="barrow"><div class="bl"><span>Today: hours × sales price</span><b id="fcToday"></b></div><div class="track" id="fcBarToday"></div></div>
        <div class="barrow"><div class="bl"><span>US-tagged: 1 dedicated seat + extra Associates</span><b id="fcTagged"></b></div><div class="track" id="fcBarTagged"></div></div>
        <p class="verdict" id="fcVerdict"></p>
      </div>
    </div>
  </div>
  <div class="card" style="margin-top:16px"><h2>Saving as US volume grows</h2><div class="note" id="ltNote"></div><div class="chart"><canvas id="cLong"></canvas></div></div>
  <div class="grid g2" style="margin-top:16px">
    <div class="card"><h2>US cost per person: today vs scenario</h2><div class="note">Per year, people included in the scenario</div><div class="chart"><canvas id="cTagCost"></canvas></div></div>
    <div class="card"><h2>How the numbers are worked out</h2><ul class="assume" id="tagAssume" style="margin:8px 0 0"></ul></div>
  </div>
</section>
</main>
<footer><div class="legal"><span id="legalLeft"></span><span id="foot"></span></div><div class="contactline" id="signText"></div></footer>

<script>
const D = /*__DATA__*/null;
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const M = D.months, T = D.territories, S = D.staff, P = D.projects, ALL = D.rows;
let R = ALL, curCat = "All";
const NM = M.length, LAST = NM - 1;
const [ri_m, ri_t, ri_s, ri_p, ri_team, ri_h, ri_cat] = [0,1,2,3,4,5,7];
const US = T.indexOf("USA");

// ---------- helpers
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const TVAR = {canada: "--c-canada", usa: "--c-usa", australia: "--c-australia", uk: "--c-uk", japan: "--c-japan"};
const tColor = i => css(TVAR[String(T[i]).toLowerCase()] || "--t" + ((i % 8) + 1));
const ym = i => M[i].split("-").map(Number);
const mLabel = i => { if (i == null || i < 0 || i >= NM) return "–"; const [y, m] = ym(i); return MON[m-1] + "-" + String(y).slice(2); };
const mIndex = s => M.indexOf(s);
const fy = i => { const [y, m] = ym(i); return D.meta.fyStart === 1 ? y : (m >= D.meta.fyStart ? y + 1 : y); };
const cy = i => ym(i)[0];
const pLabel = (mode, y) => (mode === "FY" ? "FY" : "CY") + String(y).slice(2);
const fH = v => v == null || isNaN(v) ? "–" : Math.round(v).toLocaleString("en-US");
const f1 = v => v == null || isNaN(v) ? "–" : v.toLocaleString("en-US", {maximumFractionDigits: 1, minimumFractionDigits: 1});
const fUSD = v => v == null || isNaN(v) ? "–" : "$" + Math.round(v).toLocaleString("en-US");
const fPct = v => v == null || !isFinite(v) ? "–" : (v > 0 ? "+" : "") + (Math.abs(v) >= 10 ? Math.round(v * 100).toLocaleString("en-US") : (v * 100).toFixed(1)) + "%";
const fK = v => Math.abs(v) >= 1e6 ? "$" + (v / 1e6).toFixed(2) + "M" : Math.abs(v) >= 1e3 ? "$" + (v / 1e3).toFixed(1) + "K" : fUSD(v);
const growth = (a, b) => b > 0 ? a / b - 1 : null;
const cls = v => v == null ? "" : v >= 0 ? "up" : "down";
const sw = c => `<span class="sw" style="background:${c}"></span>`;
const el = id => document.getElementById(id);
const sumRows = (fn) => { let s = 0; for (const r of R) if (fn(r)) s += r[ri_h]; return s; };

// monthly matrix [territory][month]
let MT, monthTotal;
function applyCat(c) {   // c = "All" or a category name (ESG / ET / ...)
  curCat = c;
  const ci = D.categories.indexOf(c);
  R = c === "All" ? ALL : ALL.filter(r => r[ri_cat] === ci);
  MT = T.map(() => new Array(NM).fill(0));
  for (const r of R) MT[r[ri_t]][r[ri_m]] += r[ri_h];
  monthTotal = M.map((_, i) => T.reduce((s, _, t) => s + MT[t][i], 0));
}
applyCat("All");
const inWindow = (i, n, end = LAST) => i > end - n && i <= end;

// kpi(label, value, delta?, deltaText?, sub?) - every card gets a second line so the row lines up
function kpi(label, value, delta, deltaText, sub) {
  const d = delta !== undefined ? `<div class="d"><span class="${cls(delta)}">${fPct(delta)}</span> ${deltaText || ""}</div>`
          : `<div class="d">${sub || "&nbsp;"}</div>`;
  return `<div class="kpi"><div class="l">${label}</div><div class="v">${value}</div>${d}</div>`;
}
function table(head, rows, opts = {}) {
  const th = head.map(h => `<th class="${h.n ? "n" : ""}">${h.t}</th>`).join("");
  const tr = rows.map(r => `<tr class="${r.cls || ""}">` + r.cells.map((c, j) => `<td class="${head[j].n ? "n" : ""}" ${c.style ? `style="${c.style}"` : ""}>${c.v ?? c}</td>`).join("") + "</tr>").join("");
  return `<table><thead><tr>${th}</tr></thead><tbody>${tr}</tbody></table>`;
}
function seg(id, cb) {
  el(id).addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    el(id).querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b));
    cb(b.dataset.v);
  });
}

// ---------- charts
const charts = {};
function axisOpts(stacked, horiz, money) {
  const grid = {color: css("--grid"), drawTicks: false}, ticks = {color: css("--text-2"), padding: 6};
  const val = {stacked, beginAtZero: true, grid, ticks: {...ticks, callback: v => money ? (Math.abs(v) >= 1000 ? "$" + (v / 1000).toLocaleString("en-US") + "K" : "$" + v) : v.toLocaleString("en-US")}, border: {display: false}};
  const cat = {stacked, grid: {display: false}, ticks: {...ticks, maxRotation: 0, autoSkipPadding: 12}, border: {color: css("--border")}};
  return horiz ? {x: val, y: cat} : {x: cat, y: val};
}
function makeChart(id, type, labels, datasets, o = {}) {
  if (charts[id]) charts[id].destroy();
  const horiz = o.horizontal;
  charts[id] = new Chart(el(id), {
    type, data: {labels, datasets},
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: horiz ? "y" : "x",
      interaction: {mode: "index", intersect: false},
      plugins: {
        legend: {display: datasets.length > 1, position: "top", align: "start", labels: {color: css("--text-2"), boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 2}},
        tooltip: {callbacks: {label: c => ` ${c.dataset.label}: ${o.money ? fUSD(c.raw) : f1(c.raw)}`,
                              footer: o.stacked ? items => "Total: " + (o.money ? fUSD : f1)(items.reduce((s, x) => s + (x.raw || 0), 0)) : undefined}},
        callouts: o.callouts || null,
      },
      scales: axisOpts(!!o.stacked, horiz, !!o.money),
      datasets: {bar: {borderRadius: 4, borderSkipped: "start", borderColor: css("--surface"), borderWidth: o.stacked ? {top: 2} : 0, maxBarThickness: 46, categoryPercentage: .78, barPercentage: .9},
                 line: {borderWidth: 2, pointRadius: 0, pointHoverRadius: 5, tension: .25}},
    },
    plugins: [calloutPlugin],
  });
}
// Dashed vertical lines with labels for call-outs
const calloutPlugin = {
  id: "callouts",
  afterDatasetsDraw(chart, _, opts) {
    if (!opts || !opts.items) return;
    const {ctx, chartArea: a, scales: {x}} = chart;
    ctx.save();
    ctx.font = "11px Arial, Helvetica, sans-serif";
    opts.items.forEach((c, k) => {
      const px = x.getPixelForValue(c.idx);
      if (px < a.left || px > a.right) return;
      ctx.strokeStyle = css("--text-3"); ctx.setLineDash([4, 4]); ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(px, a.top + 14 + (k % 2) * 14); ctx.lineTo(px, a.bottom); ctx.stroke();
      ctx.setLineDash([]); ctx.fillStyle = css("--text-2");
      const w = ctx.measureText(c.text).width;
      const tx = Math.min(Math.max(px - w / 2, a.left), a.right - w);
      ctx.fillText(c.text, tx, a.top + 10 + (k % 2) * 14);
    });
    ctx.restore();
  }
};
const barDs = (label, data, color, extra = {}) => ({label, data, backgroundColor: color, ...extra});

// ================= OVERVIEW
let ovMode = "FY", ovRange = 24;
function periods(mode) {
  const f = mode === "FY" ? fy : cy;
  const ys = [...new Set(M.map((_, i) => f(i)))];
  return ys.map(y => {
    const idx = M.map((_, i) => i).filter(i => f(i) === y);
    return {y, idx, partial: idx.length < 12, label: pLabel(mode, y) + (idx.length < 12 && idx[idx.length - 1] === LAST ? " YTD" : "")};
  });
}
function renderOverview() {
  const ps = periods(ovMode);
  const data = T.map((_, t) => ps.map(p => p.idx.reduce((s, i) => s + MT[t][i], 0)));
  makeChart("cPeriod", "bar", ps.map(p => p.label), T.map((n, t) => barDs(n, data[t], tColor(t))));
  const cur = ps[ps.length - 1];
  el("ovPeriodNote").textContent = `${ovMode === "FY" ? "Financial year starts " + MON[D.meta.fyStart - 1] : "Calendar year"}. ${cur.label} covers ${mLabel(cur.idx[0])} to ${mLabel(LAST)}.`;

  // KPIs: current period YTD vs same months last year
  const ytd = cur.idx.reduce((s, i) => s + monthTotal[i], 0);
  const ly = cur.idx.reduce((s, i) => s + (i - 12 >= 0 ? monthTotal[i - 12] : 0), 0);
  const ltm = M.reduce((s, _, i) => s + (inWindow(i, 12) ? monthTotal[i] : 0), 0);
  const pltm = M.reduce((s, _, i) => s + (inWindow(i, 12, LAST - 12) ? monthTotal[i] : 0), 0);
  const active = new Set(R.filter(r => inWindow(r[ri_m], 3)).map(r => r[ri_s])).size;
  const activeT = T.filter((_, t) => MT[t].some((v, i) => v > 0 && inWindow(i, 3))).length;
  el("ovKpis").innerHTML =
    kpi(`${cur.label} hours`, fH(ytd), growth(ytd, ly), "vs same months last year") +
    kpi("Hours (last 12 months)", fH(ltm), growth(ltm, pltm), "vs prior 12 months") +
    kpi("Latest month", fH(monthTotal[LAST]), growth(monthTotal[LAST], monthTotal[LAST - 12]), `${mLabel(LAST)} vs ${mLabel(LAST - 12)}`) +
    kpi("Active people", active, undefined, "", `logged hours ${mLabel(LAST - 2)} to ${mLabel(LAST)}`) +
    kpi("Active territories", `${activeT} of ${T.length}`, undefined, "", "with hours in the last 3 months");

  // table
  const head = [{t: "Territory"}, ...ps.map(p => ({t: p.label, n: 1}))];
  const rows = T.map((n, t) => ({cells: [sw(tColor(t)) + n, ...data[t].map(fH)]}));
  rows.push({cls: "tot", cells: ["Total", ...ps.map((_, j) => fH(data.reduce((s, d) => s + d[j], 0)))]});
  el("ovPeriodTable").innerHTML = table(head, rows);
  renderMonthly();
}
function renderMonthly() {
  const start = ovRange ? Math.max(0, NM - ovRange) : 0;
  const idx = M.map((_, i) => i).slice(start);
  const items = D.callouts.filter(c => mIndex(c.month) >= start).map(c => ({idx: mIndex(c.month) - start, text: mLabel(mIndex(c.month)) + " " + c.event.split(" - ")[0].split(" by ")[0]}));
  makeChart("cMonthly", "bar", idx.map(mLabel), T.map((n, t) => barDs(n, idx.map(i => MT[t][i]), tColor(t))), {stacked: true, callouts: {items}});
}
function renderStatic() {
  el("subtitle").textContent = `Source: ${D.meta.source} · Data ${mLabel(0)} to ${mLabel(LAST)} · ${R.length.toLocaleString()} time entries` + (curCat === "All" ? "" : ` · showing ${curCat} work only`);
  el("foot").textContent = `Generated ${D.meta.generated} from ${D.meta.source}. Hours by reporting month (Source Sheet).`;
  el("legalLeft").innerHTML = legalHtml(D.meta.classification);
  el("signText").innerHTML = signHtml();
  if (D.meta.storyLink) {   // return to the section of the story the viewer came from (passed as #section)
    const from = (location.hash || "").replace("#", "");
    el("backLink").href = D.meta.storyLink + "#" + (/^[a-z]+$/.test(from) ? from : "appendix"); el("backLink").hidden = false;
  }
  el("clsTop").textContent = D.meta.classification; el("title").textContent = D.meta.title; document.title = D.meta.title;
  el("ovCallouts").innerHTML = D.callouts.map(c => `<li><b>${mIndex(c.month) >= 0 ? mLabel(mIndex(c.month)) : c.month}</b><span>${c.event}<br><span class="sub">${c.territory}</span></span></li>`).join("");
  // training table
  const tm = D.meta.trainingMonth ? mIndex(D.meta.trainingMonth) : -1;
  if (tm < 0) { el("trainTable").innerHTML = "<div class='note'>Training month not in data.</div>"; return; }
  const yr = cy(tm);
  const pre = i => cy(i) === yr && i < tm, post = i => cy(i) === yr && i >= tm;
  const ts = T.map((_, t) => t).filter(t => MT[t].some((v, i) => v > 0 && cy(i) === yr));
  const val = (t, f) => MT[t].reduce((s, v, i) => s + (f(i) ? v : 0), 0);
  const tot = f => ts.reduce((s, t) => s + val(t, f), 0);
  const head = [{t: "Hours logged"}, ...ts.map(t => ({t: T[t], n: 1})), {t: ts.map(t => T[t]).join(" + "), n: 1}];
  const preN = M.filter((_, i) => pre(i)).length, postN = M.filter((_, i) => post(i)).length;
  const rows = [
    {cells: [`Pre-training (${preN} months)`, ...ts.map(t => fH(val(t, pre))), fH(tot(pre))]},
    {cells: [`Post-training (${postN} months)`, ...ts.map(t => fH(val(t, post))), fH(tot(post))]},
    {cells: ["Change in monthly average", ...ts.map(t => `<span class="${cls(growth(val(t, post) / postN, val(t, pre) / preN))}">${fPct(growth(val(t, post) / postN, val(t, pre) / preN))}</span>`),
             `<span class="up">${fPct(growth(tot(post) / postN, tot(pre) / preN))}</span>`]},
  ];
  el("trainTitle").textContent = `Training impact, CY${String(yr).slice(2)}`;
  el("trainNote").textContent = `Split at ${mLabel(tm)} (training month). Monthly average compares unequal-length periods fairly.`;
  el("trainTable").innerHTML = table(head, rows);
}

// ================= TERRITORY
let curT = 0;
function renderTerritory() {
  const t = curT, color = tColor(t);
  const first = MT[t].findIndex(v => v > 0);
  if (first < 0) {
    el("terrKpis").innerHTML = kpi("Hours", "0") ; el("terrProjects").innerHTML = el("terrPeople").innerHTML = "<div class='note'>No hours for this selection.</div>";
    ["cTerrMonthly", "cTerrTeam"].forEach(id => charts[id] && charts[id].destroy()); return;
  }
  const idx = M.map((_, i) => i).slice(first);
  makeChart("cTerrMonthly", "bar", idx.map(mLabel), [barDs(T[t], idx.map(i => MT[t][i]), color)]);
  el("terrMonthlyNote").textContent = `Since first hours logged, ${mLabel(first)}`;
  const ltm = MT[t].reduce((s, v, i) => s + (inWindow(i, 12) ? v : 0), 0);
  const pltm = MT[t].reduce((s, v, i) => s + (inWindow(i, 12, LAST - 12) ? v : 0), 0);
  const rT = R.filter(r => r[ri_t] === t), rL = rT.filter(r => inWindow(r[ri_m], 12));
  const share = ltm / M.reduce((s, _, i) => s + (inWindow(i, 12) ? monthTotal[i] : 0), 0);
  const peak = idx.reduce((b, i) => MT[t][i] > MT[t][b] ? i : b, first);
  el("terrKpis").innerHTML =
    kpi("Hours (last 12 months)", fH(ltm), pltm > 0 ? growth(ltm, pltm) : undefined, "vs prior 12 months", "new territory") +
    kpi("Share of all hours", (share * 100).toFixed(1) + "%", undefined, "", "last 12 months") +
    kpi(`Hours in ${mLabel(LAST)}`, fH(MT[t][LAST]), MT[t][LAST - 1] > 0 ? growth(MT[t][LAST], MT[t][LAST - 1]) : undefined, "vs previous month") +
    kpi("Peak month", fH(MT[t][peak]), undefined, "", mLabel(peak)) +
    kpi("People / projects", `${new Set(rL.map(r => r[ri_s])).size} / ${new Set(rL.map(r => r[ri_p])).size}`, undefined, "", "last 12 months");

  // team
  const teamH = D.teams.map((_, k) => rL.filter(r => r[ri_team] === k).reduce((s, r) => s + r[ri_h], 0));
  const order = teamH.map((v, k) => k).filter(k => teamH[k] > 0).sort((a, b) => teamH[b] - teamH[a]);
  makeChart("cTerrTeam", "bar", order.map(k => D.teams[k]), [barDs("Hours (12 mo)", order.map(k => teamH[k]), color)], {horizontal: true});
  el("terrTeamNote").textContent = "Hours in the last 12 months by team";

  // projects
  const pr = {};
  for (const r of rT) {
    const p = pr[r[ri_p]] ||= {all: 0, ltm: 0, staff: new Set(), last: 0};
    p.all += r[ri_h]; if (inWindow(r[ri_m], 12)) p.ltm += r[ri_h]; p.staff.add(r[ri_s]); p.last = Math.max(p.last, r[ri_m]);
  }
  const prow = Object.entries(pr).sort((a, b) => b[1].ltm - a[1].ltm || b[1].all - a[1].all).map(([k, v]) => ({cells: [
    `<b>${P[k].name}</b><div class="sub2">${P[k].client}</div>`, P[k].wrn, P[k].cat, fH(v.ltm), fH(v.all), v.staff.size, mLabel(v.last)]}));
  el("terrProjects").innerHTML = table([{t: "Project / client"}, {t: "WRN"}, {t: "Type"}, {t: "12-mo hrs", n: 1}, {t: "Total hrs", n: 1}, {t: "People", n: 1}, {t: "Last active"}], prow);

  // people
  const ph = {};
  for (const r of rL) ph[r[ri_s]] = (ph[r[ri_s]] || 0) + r[ri_h];
  const prs = Object.entries(ph).sort((a, b) => b[1] - a[1]).map(([s, h]) => ({cells: [S[s].name, S[s].level, S[s].team, fH(h), (h / ltm * 100).toFixed(1) + "%"]}));
  el("terrPeople").innerHTML = table([{t: "Person"}, {t: "Level"}, {t: "Team"}, {t: "Hours", n: 1}, {t: "Share", n: 1}], prs);
}

// ================= US vs NON-US
function renderUS() {
  if (US < 0) { el("usnonus").innerHTML = "<div class='card'>No USA hours in the data.</div>"; return; }
  const idx = M.map((_, i) => i).slice(Math.max(0, NM - 24));
  const us = i => MT[US][i], non = i => monthTotal[i] - MT[US][i];
  makeChart("cUsMonthly", "bar", idx.map(mLabel), [barDs("USA", idx.map(us), tColor(US)), barDs("Non-US", idx.map(non), css("--neutral"))]);
  const ltm = f => M.reduce((s, _, i) => s + (inWindow(i, 12) ? f(i) : 0), 0);
  const pl = f => M.reduce((s, _, i) => s + (inWindow(i, 12, LAST - 12) ? f(i) : 0), 0);
  const u = ltm(us), n = ltm(non);
  el("usKpis").innerHTML = kpi("US hours (last 12 months)", fH(u), growth(u, pl(us)), "vs prior 12 months") +
    kpi("Non-US hours (last 12 months)", fH(n), growth(n, pl(non)), "vs prior 12 months") +
    kpi("US share of hours", (u / (u + n) * 100).toFixed(1) + "%", undefined, "", "last 12 months") +
    kpi(`US hours in ${mLabel(LAST)}`, fH(us(LAST)), us(LAST - 1) > 0 ? growth(us(LAST), us(LAST - 1)) : undefined, "vs previous month");
  // territory-wise by FY
  const ps = periods("FY");
  const tv = (t, p) => p.idx.reduce((s, i) => s + MT[t][i], 0);
  const head = [{t: "Territory"}, ...ps.map(p => ({t: p.label, n: 1}))];
  const order = [US, ...T.map((_, t) => t).filter(t => t !== US)];
  const rows = order.map(t => ({cells: [sw(tColor(t)) + T[t], ...ps.map(p => fH(tv(t, p)))]}));
  rows.splice(1, 0, {cls: "tot", cells: ["<i>Non-US total</i>", ...ps.map(p => fH(p.idx.reduce((s, i) => s + non(i), 0)))]});
  rows.push({cls: "tot", cells: ["US share", ...ps.map(p => { const a = tv(US, p), b = p.idx.reduce((s, i) => s + monthTotal[i], 0); return b ? (a / b * 100).toFixed(1) + "%" : "–"; })]});
  el("usTerrTable").innerHTML = table(head, rows);
  // team-wise
  const th = D.teams.map((_, k) => ({us: 0, non: 0}));
  for (const r of R) if (inWindow(r[ri_m], 12)) th[r[ri_team]][r[ri_t] === US ? "us" : "non"] += r[ri_h];
  const ord = th.map((_, k) => k).filter(k => th[k].us + th[k].non > 0).sort((a, b) => (th[b].us + th[b].non) - (th[a].us + th[a].non));
  makeChart("cUsTeam", "bar", ord.map(k => D.teams[k]), [barDs("USA", ord.map(k => th[k].us), tColor(US)), barDs("Non-US", ord.map(k => th[k].non), css("--neutral"))], {horizontal: true, stacked: true});
  el("usTeamTable").innerHTML = table([{t: "Team"}, {t: "US hrs", n: 1}, {t: "Non-US hrs", n: 1}, {t: "US share", n: 1}],
    ord.map(k => ({cells: [D.teams[k], fH(th[k].us), fH(th[k].non), (th[k].us / (th[k].us + th[k].non) * 100).toFixed(1) + "%"]})));
}

// ================= GROWTH
function renderGrowth() {
  const ps = periods("FY");
  const cur = ps[ps.length - 1];
  const head = [{t: "Territory"}];
  ps.forEach((p, j) => { head.push({t: p.label + " hrs", n: 1}); if (j > 0 && !p.partial) head.push({t: "YoY", n: 1}); });
  head.push({t: `${cur.label} vs same months LY`, n: 1});
  const mk = (name, f) => {
    const cells = [name];
    ps.forEach((p, j) => {
      const v = p.idx.reduce((s, i) => s + f(i), 0); cells.push(fH(v));
      if (j > 0 && !p.partial) { const pv = ps[j - 1].idx.reduce((s, i) => s + f(i), 0); const g = growth(v, pv); cells.push(`<span class="${cls(g)}">${pv ? fPct(g) : "new"}</span>`); }
    });
    const a = cur.idx.reduce((s, i) => s + f(i), 0), b = cur.idx.reduce((s, i) => s + (i >= 12 ? f(i - 12) : 0), 0);
    const g = growth(a, b); cells.push(`<span class="${cls(g)}">${b ? fPct(g) : "new"}</span>`);
    return {cells};
  };
  const rows = T.map((n, t) => mk(sw(tColor(t)) + n, i => MT[t][i]));
  rows.push({...mk("Total", i => monthTotal[i]), cls: "tot"});
  el("growthYoY").innerHTML = table(head, rows);
  el("growthNote").textContent = `YoY is shown for complete financial years. ${cur.label} is compared with the same ${cur.idx.length} months a year earlier.`;

  // rolling 3-mo
  const idx = M.map((_, i) => i).slice(Math.max(2, NM - 24));
  const roll = (t, i) => (MT[t][i] + MT[t][i - 1] + MT[t][i - 2]) / 3;
  makeChart("cRolling", "line", idx.map(mLabel), T.map((n, t) => ({label: n, data: idx.map(i => MT[t].slice(0, i + 1).some(v => v > 0) ? roll(t, i) : null), borderColor: tColor(t), backgroundColor: tColor(t)})));

  // MoM heat table
  const last = M.map((_, i) => i).slice(Math.max(1, NM - 12)).reverse();
  const cell = g => {
    if (g == null) return {v: "–"};
    const a = Math.min(Math.abs(g), 1) * .9 + .1;
    return {v: fPct(g), style: `background:color-mix(in srgb, var(${g >= 0 ? "--pos-bg" : "--neg-bg"}) ${Math.round(a * 100)}%, transparent)`};
  };
  const mrows = last.map(i => ({cells: [mLabel(i), ...T.map((_, t) => cell(MT[t][i - 1] > 0 ? growth(MT[t][i], MT[t][i - 1]) : null)), cell(growth(monthTotal[i], monthTotal[i - 1]))]}));
  el("growthMoM").innerHTML = table([{t: "Month"}, ...T.map(n => ({t: n, n: 1})), {t: "Total", n: 1}], mrows);
}

// ================= RATES
const rateOf = (s, year) => {
  const rs = S[s].rates, ys = Object.keys(rs).map(Number).sort((a, b) => a - b);
  if (!ys.length) return null;
  if (rs[year] != null) return rs[year];
  return year < ys[0] ? rs[ys[0]] : rs[ys.filter(y => y <= year).pop()];
};
function renderRates() {
  const years = D.meta.rateYears.map(Number);
  if (!years.length) { el("rates").innerHTML = "<div class='card'>No 'Sales Price YYYY' columns found in Staff Master.</div>"; return; }
  const val = years.map(() => T.map(() => ({h: 0, v: 0})));
  const per = S.map(() => years.map(() => ({h: 0, v: 0})));
  for (const r of R) {
    const j = years.indexOf(cy(r[ri_m])); if (j < 0) continue;
    const rt = rateOf(r[ri_s], years[j]) || 0;
    val[j][r[ri_t]].h += r[ri_h]; val[j][r[ri_t]].v += r[ri_h] * rt;
    per[r[ri_s]][j].h += r[ri_h]; per[r[ri_s]][j].v += r[ri_h] * rt;
  }
  const seqC = k => years.length === 1 ? css("--seq-hi") : (k === years.length - 1 ? css("--seq-hi") : css("--seq-lo"));
  makeChart("cRateValue", "bar", T, years.map((y, j) => barDs(String(y) + (cy(LAST) === y ? " YTD" : ""), val[j].map(x => x.v), seqC(j))), {money: true});
  const tot = j => val[j].reduce((a, x) => ({h: a.h + x.h, v: a.v + x.v}), {h: 0, v: 0});
  let k = "";
  years.forEach((y, j) => { const t = tot(j); k += kpi(`Billed value ${y}${cy(LAST) === y ? " YTD" : ""}`, fUSD(t.v), undefined, "", `${fH(t.h)} hours`); });
  const a = tot(0), b = tot(years.length - 1);
  const br = x => x.h ? x.v / x.h : null;
  k += kpi(`Blended rate ${years[years.length - 1]}`, br(b) != null ? "$" + br(b).toFixed(2) : "–", years.length > 1 && br(a) && br(b) ? growth(br(b), br(a)) : undefined, `vs ${years[0]} ($${(br(a) || 0).toFixed(2)})`, "hours × sales price ÷ hours");
  el("rateKpis").innerHTML = k;
  el("rateTerr").innerHTML = table([{t: "Territory"}, ...years.flatMap(y => [{t: `${y} hrs`, n: 1}, {t: `${y} $/hr`, n: 1}])],
    T.map((n, t) => ({cells: [sw(tColor(t)) + n, ...years.flatMap((_, j) => [fH(val[j][t].h), val[j][t].h ? "$" + (val[j][t].v / val[j][t].h).toFixed(2) : "–"])]})));
  const head = [{t: "Person"}, {t: "Level"}, {t: "Current tag"}, ...years.map(y => ({t: `Rate ${y}`, n: 1})), ...(years.length > 1 ? [{t: "Change", n: 1}] : []),
    ...years.map(y => ({t: `Hrs ${y}`, n: 1})), ...years.map(y => ({t: `Value ${y}`, n: 1})), {t: "Model FY27 rate", n: 1}, {t: `Model rate vs ${years[years.length - 1]} price`, n: 1}];
  const rows = S.map((s, i) => {
    const rs = years.map(y => s.rates[y]); const mr = D.model[s.level];
    const gap = mr && rs[rs.length - 1] != null ? mr.rate - rs[rs.length - 1] : null;
    return {cells: [s.name, s.level, s.tag || "–", ...rs.map(v => v != null ? "$" + v.toFixed(2) : "–"),
      ...(years.length > 1 ? [fPct(growth(rs[rs.length - 1], rs[0]))] : []),
      ...per[i].map(x => fH(x.h)), ...per[i].map(x => fUSD(x.v)), mr ? "$" + mr.rate.toFixed(2) : "–",
      gap == null ? "–" : gap === 0 ? "same" : `<span class="${gap < 0 ? "up" : "down"}">$${Math.abs(gap).toFixed(2)} ${gap < 0 ? "lower" : "higher"}</span>`]};
  });
  el("rateStaff").innerHTML = table(head, rows);
}

// ================= TAGGING SCENARIO
const PRESETS = D.meta.presets;
const baseRates = () => JSON.parse(JSON.stringify(D.model));
let SC = {...PRESETS.find(p => p.name === "Base") || PRESETS[0], rates: baseRates(), over: {}, custom: false};

function scenario(st) {
  const N = st.window, ann = 12 / N, g = 1 + st.growth, thr = st.threshold;
  const ppl = S.map((s, i) => ({s, i, h: T.map(() => 0), tot: 0, usCur: 0}));
  for (const r of R) {
    if (!inWindow(r[ri_m], N)) continue;
    const p = ppl[r[ri_s]]; p.h[r[ri_t]] += r[ri_h]; p.tot += r[ri_h];
    if (r[ri_t] === US) p.usCur += r[ri_h] * (rateOf(r[ri_s], cy(r[ri_m])) || 0);
  }
  const rows = ppl.filter(p => p.tot > 0).map(p => {
    const prim = p.h.indexOf(Math.max(...p.h)), share = p.h[prim] / p.tot;
    const rec = share >= thr - 1e-9 ? T[prim] : "Shared";
    const tagT = p.s.tagTerritory;
    const status = rec === "Shared" ? "Review: split across territories" : rec === tagT ? "Correctly tagged" : `Tag to ${rec}`;
    const usH = US >= 0 ? p.h[US] : 0;
    const auto = rec === "USA" && tagT !== "USA";
    const eligible = usH > 0 && tagT !== "USA";
    const inc = eligible && (st.over[p.i] ?? auto);
    const rc = st.rates[p.s.level];
    const annH = usH * ann, today = p.usCur * ann, avg = usH ? p.usCur / usH : 0;
    const cost = h => {
      if (!rc) return null;
      const hourly = h * rc.rate;
      return st.method === "best" && rc.cpp > 0 && rc.cpp < hourly ? {v: rc.cpp, via: "Dedicated seat"} : {v: hourly, via: "Hourly"};
    };
    const after = cost(annH), afterN = cost(annH * g);
    return {...p, prim, share, rec, status, usH, annH, today, todayN: annH * g * avg, after, afterN, inc, eligible, auto,
            saving: inc && after ? today - after.v : 0, savingN: inc && afterN ? annH * g * avg - afterN.v : 0};
  }).sort((a, b) => (b.inc - a.inc) || (b.usH - a.usH));
  const inc = rows.filter(r => r.inc), sum = f => inc.reduce((s, x) => s + f(x), 0);
  const tot = {people: inc.length, hours: sum(r => r.annH), today: sum(r => r.today), after: sum(r => r.after ? r.after.v : r.today),
               saving: sum(r => r.saving), savingN: sum(r => r.savingN), todayN: sum(r => r.todayN), seats: inc.filter(r => r.after && r.after.via !== "Hourly").length,
               fixes: rows.filter(r => r.status.startsWith("Tag to")).length, review: rows.filter(r => r.rec === "Shared").length};
  return {rows, inc, tot, N, ann, thr, g};
}

const pctTxt = g => (g > 0 ? "+" : "") + Math.round(g * 100) + "%";
const methodName = m => m === "best" ? "Hourly or seat" : "Hourly rate";
const presetDesc = p => `${Math.round(p.threshold * 100)}% threshold · ${p.window}-month look-back · ${pctTxt(p.growth)} US volume · ${methodName(p.method).toLowerCase()}`;

function syncControls() {
  el("scThr").value = Math.round(SC.threshold * 100); el("scThrV").textContent = Math.round(SC.threshold * 100) + "%";
  el("scGrowth").value = Math.round(SC.growth * 100); el("scGrowthV").textContent = pctTxt(SC.growth);
  el("scWinV").textContent = SC.window + " mo"; el("scMethodV").textContent = SC.method === "best" ? "Best" : "Hourly";
  el("scWin").querySelectorAll("button").forEach(b => b.classList.toggle("on", +b.dataset.v === SC.window));
  el("scMethod").querySelectorAll("button").forEach(b => b.classList.toggle("on", b.dataset.v === SC.method));
}

function renderTagging() {
  syncControls();
  const res = scenario(SC), t = res.tot, N = res.N;
  const pct = t.today ? t.saving / t.today * 100 : 0;
  const winStart = Math.max(0, NM - N);
  const nm = SC.custom ? "Your scenario" : SC.name;
  const ppl = n => `${n} ${n === 1 ? "person" : "people"}`;
  const others = t.fixes - rows_movers(res);
  const avgRate = t.hours ? t.today / t.hours : 0;

  el("scName").textContent = `Showing: ${SC.custom ? "your scenario" : SC.name + " scenario"}`;
  el("tagHeadline").innerHTML = t.people
    ? `Moving ${ppl(t.people)} to the US tag saves about <b>${fK(t.saving)} a year</b>, ${pct.toFixed(0)}% of their US cost`
    : `No one meets the ${Math.round(SC.threshold * 100)}% threshold in this scenario. Lower the threshold or tick people below to include them`;
  el("tagPoints").innerHTML = [
    `<b>Why:</b> ${SC.custom ? "the people in this scenario" : `${ppl(t.people)} spent at least ${Math.round(SC.threshold * 100)}% of their time on US work in the last ${N} months but`} sit under a non-US tag, so US work is charged at their own sales price (about ${fUSD(avgRate)}/hr on average).`,
    `<b>What changes:</b> once US-tagged, their ${fH(t.hours)} US hours a year are charged at ${SC.method === "best" ? "the Global Model FY27 rate, or a dedicated seat where that is cheaper" : "the Global Model FY27 rate for their level"}.`,
    `<b>Next year:</b> with US volume at ${pctTxt(SC.growth)}, the saving would be about <b>${fK(t.savingN)}</b>.`,
    `<b>Also:</b> ${ppl(others)} should be tagged to other territories and ${ppl(t.review)} ${t.review === 1 ? "is" : "are"} split across territories and need a review.`,
  ].map(x => `<li>${x}</li>`).join("");

  el("tagKpis").innerHTML =
    kpi("US cost today", fUSD(t.today), undefined, "", `per year, ${ppl(t.people)}`) +
    kpi("US cost after US tagging", fUSD(t.after), undefined, "", SC.method === "best" ? "hourly or seat" : "at Global Model rates") +
    kpi("Annual saving", fUSD(t.saving), undefined, "", `<span class="up">${pct.toFixed(1)}%</span> lower than today`) +
    kpi("Next-year saving", fUSD(t.savingN), undefined, "", `at ${pctTxt(SC.growth)} US volume`) +
    kpi("US hours moving", fH(t.hours), undefined, "", `per year${t.seats ? ` · ${t.seats} as dedicated seat` : ""}`);

  // rate card
  const changed = (lv, k) => D.model[lv] && D.model[lv][k] !== SC.rates[lv][k] ? ' style="font-weight:700;color:var(--orange)"' : "";
  el("scRates").innerHTML = table([{t: "Level"}, {t: "Rate per hour", n: 1}, {t: "Cost per person", n: 1}, {t: "Seat pays off at", n: 1}],
    Object.entries(SC.rates).map(([lv, r]) => ({cells: [`${r.name}`,
      `<span class="money"><span>$</span><input class="num" type="number" min="0" step="1" data-lv="${lv}" data-k="rate" value="${r.rate}"${changed(lv, "rate")} aria-label="${r.name} rate per hour"></span>`,
      `<span class="money"><span>$</span><input class="num" type="number" min="0" step="100" data-lv="${lv}" data-k="cpp" value="${r.cpp}"${changed(lv, "cpp")} aria-label="${r.name} cost per person"></span>`,
      r.rate ? fH(r.cpp / r.rate) + " hrs" : "–"]})));

  // preset tiles
  el("scTiles").innerHTML = results0().map(x => `<button class="tile ${!SC.custom && SC.name === x.name ? "on" : ""}" data-v="${x.name}">
      <div class="tn">${x.name}</div><div class="td">${presetDesc(x.st)}</div>
      <div class="tv">${fK(x.r.tot.saving)}</div><div class="tl">a year · ${fK(x.r.tot.savingN)} next year</div></button>`).join("") +
    (SC.custom ? `<div class="tile on" style="cursor:default"><div class="tn">Your scenario</div><div class="td">${presetDesc(SC)}${Object.keys(SC.over).length ? ` · ${Object.keys(SC.over).length} manual change(s)` : ""}</div>
      <div class="tv">${fK(t.saving)}</div><div class="tl">a year · ${fK(t.savingN)} next year</div></div>`
      : `<div class="tile ghost"><div class="tn">Your scenario</div><div class="td">Move a slider or tick a person below to build your own</div><div class="tv">–</div><div class="tl">&nbsp;</div></div>`);

  // comparison
  function results0() { return PRESETS.map(p => ({name: p.name, st: {...p, rates: SC.rates, over: {}}})).map(x => ({...x, r: scenario(x.st)})); }
  const list = PRESETS.map(p => ({name: p.name, st: {...p, rates: SC.rates, over: {}}}));
  if (SC.custom) list.push({name: "Your scenario", st: SC});
  const results = list.map(x => ({...x, r: x.name === "Your scenario" ? res : scenario(x.st)}));
  const active = SC.custom ? "Your scenario" : SC.name;
  const hl = n => n === active ? "hl" : "";
  const metric = (label, f) => `<tr><td>${label}</td>` + results.map(x => `<td class="n ${hl(x.name)}">${f(x)}</td>`).join("") + "</tr>";
  el("scCompare").innerHTML = `<table><thead><tr><th></th>${results.map(x => `<th class="n ${hl(x.name)}">${x.name}</th>`).join("")}</tr></thead><tbody>` +
    metric("Tagging threshold", x => Math.round(x.st.threshold * 100) + "%") +
    metric("Look-back period", x => x.st.window + " months") +
    metric("US volume next year", x => pctTxt(x.st.growth)) +
    metric("Costing method", x => methodName(x.st.method)) +
    metric("People moving", x => x.r.tot.people) +
    metric("US hours moving per year", x => fH(x.r.tot.hours)) +
    metric("US cost today", x => fUSD(x.r.tot.today)) +
    metric("US cost after US tagging", x => fUSD(x.r.tot.after)) +
    metric("<b>Annual saving</b>", x => `<b>${fUSD(x.r.tot.saving)}</b>`) +
    metric("Saving %", x => x.r.tot.today ? (x.r.tot.saving / x.r.tot.today * 100).toFixed(1) + "%" : "–") +
    metric("<b>Next-year saving</b>", x => `<b>${fUSD(x.r.tot.savingN)}</b>`) + "</tbody></table>";
  makeChart("cScCompare", "bar", results.map(x => x.name), [
    barDs("Today's volume", results.map(x => x.r.tot.saving), css("--neutral")),
    barDs("Next year (projected)", results.map(x => x.r.tot.savingN), css("--orange"))], {money: true});

  // people table
  el("tagNote").innerHTML = `Hours from ${mLabel(winStart)} to ${mLabel(LAST)}, shown per year. Tick or untick <b>Include</b> to add or remove someone; hover a bar to see their hours by territory.`;
  el("tagLegend").innerHTML = T.map((n, i) => `<span>${sw(tColor(i))}${n}</span>`).join("");
  const mix = p => `<div class="mix">` + T.map((n, i) => p.h[i] > 0 ? `<div title="${n}: ${f1(p.h[i])} hrs (${(p.h[i] / p.tot * 100).toFixed(0)}%)" style="width:${p.h[i] / p.tot * 100}%;background:${tColor(i)}"></div>` : "").join("") + `</div>`;
  const pill = r => {
    if (r.inc && !r.auto) return `<span class="pill man">Added manually</span>`;
    if (!r.inc && r.auto) return `<span class="pill man">Excluded manually</span>`;
    return `<span class="pill ${r.status.startsWith("Tag to") ? "act" : ""}">${r.status}</span>`;
  };
  const head = [{t: "Include"}, {t: "Person"}, {t: "Level"}, {t: "Current tag"}, {t: "Hours mix"}, {t: "Main territory"}, {t: "Share", n: 1},
    {t: "Recommendation"}, {t: "US hours / yr", n: 1}, {t: "US cost today", n: 1}, {t: "After US tagging", n: 1}, {t: "Basis"}, {t: "Saving / yr", n: 1}];
  const trs = res.rows.map(r => ({cells: [
    r.eligible ? `<input type="checkbox" data-i="${r.i}" ${r.inc ? "checked" : ""} aria-label="Include ${r.s.name}">` : `<span class="sub2">${r.s.tagTerritory === "USA" ? "Already US" : "No US hrs"}</span>`,
    r.s.name, r.s.level, r.s.tag || "–", mix(r), T[r.prim], (r.share * 100).toFixed(0) + "%", pill(r),
    r.annH ? fH(r.annH) : "–", r.usH ? fUSD(r.today) : "–", r.inc && r.after ? fUSD(r.after.v) : "–", r.inc && r.after ? r.after.via : "–",
    r.inc ? `<span class="${r.saving >= 0 ? "up" : "down"}">${fUSD(r.saving)}</span>` : "–"]}));
  trs.push({cls: "tot", cells: ["", "Total in scenario", "", "", "", "", "", "", fH(t.hours), fUSD(t.today), fUSD(t.after), "", fUSD(t.saving)]});
  el("tagTable").innerHTML = table(head, trs);

  makeChart("cTagCost", "bar", res.inc.map(r => r.s.name), [
    barDs("Today (sales price)", res.inc.map(r => r.today), css("--neutral")),
    barDs("After US tagging", res.inc.map(r => r.after ? r.after.v : r.today), css("--orange"))], {money: true});

  el("tagAssume").innerHTML = [
    `<b>Current tagging:</b> ${esc(D.meta.taggingNote)}`,
    `<b>Look-back:</b> ${mLabel(winStart)} to ${mLabel(LAST)} (${N} months), annualised ×${res.ann.toFixed(2)}.`,
    `<b>Recommended tag:</b> the territory with at least ${Math.round(SC.threshold * 100)}% of a person's hours; below that the person is marked "Shared" for review. Anyone already US-tagged would not be counted as moving.`,
    `<b>Cost today:</b> US hours × the person's sales price for that year (Staff Master).`,
    `<b>Cost in scenario:</b> US hours × the rate card's Rate/hr for the person's level${SC.method === "best" ? ", or the level's Cost/Person if a dedicated seat is cheaper" : ""}.`,
    `<b>Next year:</b> annual US hours × (1 ${SC.growth >= 0 ? "+" : "−"} ${Math.abs(Math.round(SC.growth * 100))}%), same rates.`,
    `<b>To confirm with Finance:</b> that US-tagged hours are charged at the Global Model rate, and how Cost/Person applies to a dedicated seat.`,
  ].map(x => `<li>${x}</li>`).join("");
  renderCapacity(); renderLong(res);
}
// ---------- one full-time person: hourly today vs 1 dedicated seat + extra Associates
const CAPH = D.meta.capacity || 1800;
let FC = null;
function avgPrice(lv) {
  const ys = D.meta.rateYears.map(Number), yr = ys.length ? Math.max(...ys) : null;
  const rs = S.filter(x => x.level === lv).map(x => yr != null ? x.rates[yr] : undefined).filter(v => v != null);
  return rs.length ? Math.round(rs.reduce((a, b) => a + b, 0) / rs.length) : Math.round(((SC.rates[lv] || {}).rate || 50) * 1.25);
}
function renderCapacity() {
  const RC = SC.rates, lvs = Object.keys(RC).filter(l => l !== "A");
  if (!lvs.length) return;
  const assoc = RC.A || Object.values(RC).sort((a, b) => a.cpp - b.cpp)[0];
  if (!FC || !RC[FC.lv]) FC = {lv: RC.M ? "M" : lvs[0], hrs: CAPH, rate: avgPrice(RC.M ? "M" : lvs[0])};
  el("fcLevel").innerHTML = lvs.map(l => `<button data-v="${l}" class="${l === FC.lv ? "on" : ""}">${esc(RC[l].name)}</button>`).join("");
  el("fcHrs").max = CAPH; el("fcHrs").value = FC.hrs; el("fcHrsV").textContent = fH(FC.hrs) + " hrs";
  if (document.activeElement !== el("fcRate")) el("fcRate").value = FC.rate;
  const name = RC[FC.lv].name, cpp = RC[FC.lv].cpp, todayC = FC.hrs * FC.rate, left = todayC - cpp;
  const nA = left > 0 && assoc.cpp > 0 ? Math.floor(left / assoc.cpp) : 0, rest = left - nA * assoc.cpp;
  const max = Math.max(todayC, cpp + Math.max(left, 0), 1), w = v => (v / max * 100).toFixed(2) + "%";
  el("fcToday").textContent = fUSD(todayC);
  el("fcBarToday").innerHTML = `<div class="seg-b today" style="width:${w(todayC)}">${esc(name)}: ${fK(todayC)}</div>`;
  let parts = `<div class="seg-b mgr" style="width:${w(cpp)}">1 ${esc(name)} seat ${fK(cpp)}</div>`;
  for (let k = 0; k < nA; k++) parts += `<div class="seg-b assoc" style="width:${w(assoc.cpp)}">+ Associate ${fK(assoc.cpp)}</div>`;
  if (rest > 0) parts += `<div class="seg-b save" style="width:${w(rest)}">${rest / max > .08 ? "still saved " + fK(rest) : ""}</div>`;
  el("fcBarTagged").innerHTML = parts;
  el("fcTagged").textContent = fUSD(cpp + nA * assoc.cpp);
  const be = FC.rate ? cpp / FC.rate : 0, beA = FC.rate ? (cpp + assoc.cpp) / FC.rate : 0;
  el("fcVerdict").innerHTML = nA >= 1 ? `The US gets the ${esc(name)} <b>plus ${nA === 1 ? "an extra Associate" : nA + " extra Associates"}</b> and still pays <b>${fK(rest)}</b> less than today.`
    : left > 0 ? `A dedicated seat saves <b>${fK(left)}</b> a year. ${beA <= CAPH ? `From <b>${fH(beA)} hrs</b> the saving also pays for a full Associate.` : ""}`
    : `At ${fH(FC.hrs)} hrs, hourly billing is the better option. A dedicated seat pays off from <b>${fH(be)} hrs</b> a year.`;
  el("fcNote").innerHTML = `${esc(name)} seat = Cost/Person ${fUSD(cpp)}; Associate seat = ${fUSD(assoc.cpp)}. One person, one seat: up to ${fH(CAPH)} hrs a year.`;
}
// ---------- long term: saving for the people in the scenario as US hours grow
function renderLong(res) {
  const assoc = SC.rates.A || Object.values(SC.rates).sort((a, b) => a.cpp - b.cpp)[0];
  const people = res.inc.filter(r => r.annH > 0 && SC.rates[r.s.level]);
  if (!people.length) { el("ltNote").textContent = "No one is included in the scenario, so there is nothing to project."; if (charts.cLong) charts.cLong.destroy(); return; }
  const steps = []; for (let g = 0; g <= 200; g += 10) steps.push(g);
  const at = (m, seats) => people.reduce((s, r) => {
    const rc = SC.rates[r.s.level], h = r.annH * m, cur = h * (r.today / r.annH), hourly = h * rc.rate;
    // full-time people on seats; leftover hours billed hourly, or on one more seat if that is cheaper
    const full = Math.floor(h / CAPH), restH = h - full * CAPH, seatCost = full * rc.cpp + Math.min(restH * rc.rate, rc.cpp);
    return s + cur - (seats ? Math.min(hourly, seatCost) : hourly);
  }, 0);
  const hourly = steps.map(g => at(1 + g / 100, false)), best = steps.map(g => at(1 + g / 100, true));
  const cross = steps.find((g, k) => best[k] >= assoc.cpp);
  el("ltNote").textContent = `${people.length} ${people.length === 1 ? "person" : "people"} in the scenario, hours scaled up. ` +
    (cross != null ? (cross === 0 ? "The saving already covers a full Associate seat." : `The saving passes the cost of a full Associate seat at about +${cross}% US volume.`) : "The saving grows with US volume.") +
    ` Seats: one per ${fH(CAPH)} hrs (Cost/Person).`;
  makeChart("cLong", "line", steps.map(g => g === 0 ? "Today" : "+" + g + "%"), [
    {label: "Saving: cheapest of hourly or dedicated seats", data: best, borderColor: css("--orange"), backgroundColor: css("--orange"), borderWidth: 3},
    {label: "Saving: hourly model rate only", data: hourly, borderColor: css("--charcoal"), backgroundColor: css("--charcoal"), borderDash: [5, 4]},
    {label: "Cost of 1 Associate seat", data: steps.map(() => assoc.cpp), borderColor: css("--tangerine"), backgroundColor: css("--tangerine"), borderWidth: 1.5, borderDash: [2, 3]},
    {label: "Cost of 2 Associate seats", data: steps.map(() => 2 * assoc.cpp), borderColor: css("--neutral"), backgroundColor: css("--neutral"), borderWidth: 1.5, borderDash: [2, 3]}], {money: true});
}
// tag fixes other than the US moves (for the KPI)
const rows_movers = res => res.rows.filter(r => r.status === "Tag to USA").length;

function scChange(patch) { Object.assign(SC, patch, {custom: true}); renderTagging(); }
function initScenarioControls() {
  el("scTiles").addEventListener("click", e => {
    const b = e.target.closest("button.tile"); if (!b) return;
    SC = {...PRESETS.find(p => p.name === b.dataset.v), rates: SC.rates, over: {}, custom: false}; renderTagging();
  });
  el("scThr").addEventListener("input", e => scChange({threshold: +e.target.value / 100}));
  el("scGrowth").addEventListener("input", e => scChange({growth: +e.target.value / 100}));
  el("scWin").addEventListener("click", e => { const b = e.target.closest("button"); if (b) scChange({window: +b.dataset.v}); });
  el("scMethod").addEventListener("click", e => { const b = e.target.closest("button"); if (b) scChange({method: b.dataset.v}); });
  el("scReset").addEventListener("click", () => { SC = {...PRESETS.find(p => p.name === "Base") || PRESETS[0], rates: baseRates(), over: {}, custom: false}; renderTagging(); });
  el("scRates").addEventListener("change", e => {
    const i = e.target.closest("input"); if (!i) return;
    SC.rates[i.dataset.lv][i.dataset.k] = Math.max(0, +i.value || 0); renderTagging();
  });
  el("fcLevel").addEventListener("click", e => { const b = e.target.closest("button"); if (!b) return; FC.lv = b.dataset.v; FC.rate = avgPrice(FC.lv); renderCapacity(); });
  el("fcHrs").addEventListener("input", e => { FC.hrs = +e.target.value; renderCapacity(); });
  el("fcRate").addEventListener("input", e => { FC.rate = Math.max(0, +e.target.value || 0); renderCapacity(); });
  el("tagTable").addEventListener("change", e => {
    const c = e.target.closest("input[type=checkbox]"); if (!c) return;
    SC.over[c.dataset.i] = c.checked; SC.custom = true; renderTagging();
  });
}

// ---------- sign-off + classification (name/email set in build_dashboard.py: CONTACT_NAME / CONTACT_EMAIL)
function signHtml() {
  const c = D.meta.contact || {}, team = c.team || "";
  const ppl = (c.people || []).map(p => !p.email ? esc(p.name) : p.name ? `${esc(p.name)} (<a href="mailto:${esc(p.email)}">${esc(p.email)}</a>)` : `<a href="mailto:${esc(p.email)}">${esc(p.email)}</a>`);
  if (!ppl.length) return "";
  return `For any changes, additions or further analysis, please reach out to ${ppl.join(" or ")}${team ? `, ${esc(team)}` : ""}.`;
}
const legalHtml = cls => esc(cls).replace(/^((?:\w+ )?Confidential)/, "<b>$1</b>");

// ================= TEAM + THEMES (text from team_content.json)
const esc = v => String(v ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
function renderTeam() {
  const c = D.content.team; if (!c) return;
  el("tmSummary").textContent = c.summary;
  el("tmLeaders").innerHTML = c.leaders.map(l => `<li><div class="nm">${esc(l.name)}</div><div class="rl">${esc(l.role)}</div>${l.note ? `<div class="nt">${esc(l.note)}</div>` : ""}</li>`).join("");
  el("tmMembers").textContent = c.members;
  el("tmContacts").innerHTML = c.keyContacts.map(k => {
    const t = T.indexOf(k.territory);
    return `<li><span class="sw" style="background:${t >= 0 ? tColor(t) : css("--orange")}"></span><span class="tt">${esc(k.territory)}</span><span>${esc(k.name)} <span class="rr">· ${esc(k.role)}</span></span></li>`;
  }).join("");
  el("tmCurrent").innerHTML = c.currentServices.map(x => `<li>${esc(x)}</li>`).join("");
  el("tmFuture").innerHTML = c.futureServices.map(x => `<li>${esc(x)}</li>`).join("");
  el("tmFrameworks").innerHTML = c.frameworks.map(x => `<span>${esc(x)}</span>`).join("");
  el("tmDatabases").innerHTML = c.databases.map(x => `<span>${esc(x)}</span>`).join("");
  el("tmTools").innerHTML = c.aiTools.map(x => `<span>${esc(x)}</span>`).join("");
}
function renderThemes() {
  const c = D.content.themes; if (!c) return;
  el("thTitle").textContent = c.title;
  [1, 2, 3, 4].forEach(k => { const show = k <= c.items.length; el("th" + k).style.display = show ? "" : "none";
    const n = document.querySelector(".dn.n" + k); if (n) n.style.visibility = show ? "" : "hidden"; });
  c.items.slice(0, 4).forEach((it, k) => {
    el("th" + (k + 1)).innerHTML = `<span class="num">${String(k + 1).padStart(2, "0")}</span><h3>${esc(it.title)}</h3><div class="tag">${esc(it.tagline)}</div>` +
      `<ul>${it.points.map(([h, d]) => `<li><b>${esc(h)}:</b> ${esc(d)}</li>`).join("")}</ul>`;
  });
  // live link to the data: Energy Transition hours from the ESG/ET column
  const et = D.categories.indexOf("ET");
  if (et < 0) { el("thKpis").innerHTML = ""; return; }
  let h = 0, tot = 0, ph = 0; const pr = new Set(), tr = new Set();
  for (const r of ALL) {
    if (inWindow(r[ri_m], 12)) { tot += r[ri_h]; if (r[ri_cat] === et) { h += r[ri_h]; pr.add(r[ri_p]); tr.add(r[ri_t]); } }
    else if (inWindow(r[ri_m], 12, LAST - 12) && r[ri_cat] === et) ph += r[ri_h];
  }
  el("thKpis").innerHTML =
    kpi("Energy Transition hours", fH(h), ph > 0 ? growth(h, ph) : undefined, "vs prior 12 months", "last 12 months") +
    kpi("Share of all hours", (tot ? h / tot * 100 : 0).toFixed(1) + "%", undefined, "", "last 12 months") +
    kpi("ET projects", pr.size, undefined, "", `across ${tr.size} ${tr.size === 1 ? "territory" : "territories"}`);
}


// ---------- motion: count a number up from 0 (keeps prefix/suffix, decimals and thousands separators)
const REDUCED = matchMedia("(prefers-reduced-motion: reduce)").matches;
function countUp(node, ms = 900) {
  if (REDUCED || !node) return;
  const txt = node.textContent, m = txt.match(/\d[\d,]*(\.\d+)?/); if (!m) return;
  const raw = m[0], target = parseFloat(raw.replace(/,/g, "")); if (!isFinite(target) || target === 0) return;
  const dec = (raw.split(".")[1] || "").length, commas = raw.includes(","), pre = txt.slice(0, m.index), post = txt.slice(m.index + raw.length);
  const fmt = v => commas ? v.toLocaleString("en-US", {minimumFractionDigits: dec, maximumFractionDigits: dec}) : v.toFixed(dec);
  const t0 = performance.now();
  const step = t => { const p = Math.min(1, (t - t0) / ms), e = 1 - Math.pow(1 - p, 3);
    node.textContent = p < 1 ? pre + fmt(target * e) + post : txt; if (p < 1) requestAnimationFrame(step); };
  requestAnimationFrame(step);
}

// ---------- wiring
const RENDER = {team: renderTeam, themes: renderThemes, overview: renderOverview, territory: renderTerritory, usnonus: renderUS, growth: renderGrowth, rates: renderRates, tagging: renderTagging};
let curTab = "overview";
function renderAll() { RENDER[curTab](); }
el("terrPick").innerHTML = T.map((n, t) => `<button data-v="${t}" class="${t === 0 ? "on" : ""}">${n}</button>`).join("");
seg("terrPick", v => { curT = +v; renderTerritory(); });
el("catPick").innerHTML = ["All", ...D.categories].map(c => `<button data-v="${c}" class="${c === "All" ? "on" : ""}">${c === "ET" ? "ET (Energy Transition)" : c}</button>`).join("");
seg("catPick", v => { applyCat(v); renderStatic(); renderAll(); });
seg("ovPeriod", v => { ovMode = v; renderOverview(); });
seg("ovRange", v => { ovRange = +v; renderMonthly(); });
function showTab(name) {
  document.querySelectorAll("#tabs button").forEach(x => x.classList.toggle("on", x.dataset.tab === name));
  document.querySelectorAll("main section").forEach(s => s.classList.toggle("on", s.id === name));
  el("catRow").style.display = name === "team" || name === "themes" ? "none" : "";  // work-type filter only affects data tabs
  el("signText").style.display = name === "tagging" ? "" : "none";   // contact line only on the last tab
  curTab = name; RENDER[curTab]();
  el(name).querySelectorAll(".kpi .v").forEach(v => countUp(v, 800));   // count up once when a tab opens (not on slider moves)
}
el("tabs").addEventListener("click", e => { const b = e.target.closest("button"); if (b) showTab(b.dataset.tab); });
if (!D.content) document.querySelectorAll("#tabs button.ct").forEach(b => b.remove());
else curTab = "team";
Chart.defaults.animation.duration = REDUCED ? 0 : 450; Chart.defaults.animation.easing = "easeOutQuart";
Chart.defaults.font.family = 'Arial, Helvetica, sans-serif';
Chart.defaults.color = css("--text-2");
initScenarioControls(); renderStatic(); showTab(curTab);
matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { Chart.defaults.color = css("--text-2"); renderAll(); });
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
