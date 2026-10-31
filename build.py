"""
ESG Territory Hours - one-file builder
=====================================
Reads one Excel workbook and writes three offline HTML pages:
  ESG_Summary.html           one-page summary (who we are, key points, US vs Non-US hours, value of hours)
  ESG_Hours_Dashboard.html   full dashboard (territories, growth, rates, US tagging scenario planner)
  ESG_US_Story_v2.html       partner story (arrow keys move between sections)

Run:   py build.py                      (or double-click build.py)
       py build.py --excel "D:\\Data\\Hours.xlsx"

Needs Python 3.10+ with:  py -m pip install pandas openpyxl
Files:  build.py (this) + template.html (page design) + your workbook. Missing input tabs are added
to the workbook on the first run (a backup copy is kept).
"""
# ================================ SETTINGS ================================
# Put private values (real workbook path, your name/email) in local_settings.py
# with the same names - that file is ignored by Git.
EXCEL_FILE = "ESG_Hours_Sample.xlsx"   # file name next to build.py, or a full path

CONTACT_NAME = ""                      # sign-off at the end of both pages
CONTACT_EMAIL = ""
CONTACT_NAME_2 = ""                    # optional second contact
CONTACT_EMAIL_2 = ""
CONTACT_TEAM = ""                      # optional, shown after the contacts

LOGO_FILE = "logo.png"                 # optional logo next to build.py (PNG/JPG/SVG)
LOGO_TEXT = ""                         # text shown when there is no logo file
REPORT_TITLE = "ESG Delivery: Territory Hours"      # if the workbook has no 'Report title' assumption
CLASSIFICATION = "Confidential | Internal use only" # if the workbook has no 'Classification' assumption

TEMPLATE_FILE = "template.html"
DASHBOARD_FILE = "ESG_Hours_Dashboard.html"
STORY_FILE = "ESG_US_Story_v2.html"
SUMMARY_FILE = "ESG_Summary.html"      # one-page summary
OPEN_WHEN_DONE = True                  # open the story page at the end
# ==========================================================================
try:
    from local_settings import *       # noqa: F401,F403
except ImportError:
    pass

import argparse
import base64
import json
import os
import re
import shutil
import sys
import webbrowser
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Python packages missing. Install them once with:\n    py -m pip install pandas openpyxl\nthen run again.")

HERE = Path(__file__).resolve().parent
ORANGE, CHARCOAL, PEACH = "D04A02", "2D2D2D", "FDEDE4"

# ------------------------------------------------------------ workbook layout ---
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
SHEET_SETUP = "Setup"           # older layout (still read if there is no Team Sheet)
SHEET_TEAMSHEET = "Team Sheet"  # THE one sheet you edit: names, titles, team, services, story, settings
TS_COLS = ["Section", "Item", "Value", "Detail 1", "Detail 2", "Detail 3", "How to fill / where it appears"]
OLD_INPUT_TABS = ["README", "Setup", "Team", "Services", "ET Themes", "Story", "Scenario Presets", "Tag Mapping", "Assumptions"]
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

# ------------------------------------------------- default text for the input tabs ---
# Used only to pre-fill tabs that are missing from a workbook. After that, edit the text in Excel.
STORY = {
    "title": "Tagging US-focused AC talent to the US",
    "subtitle": "A win-win proposal: the US pays less for the same people, and AC builds a dedicated US bench",
    "meeting": "Discussion with US Partners",
    "presenter": "Your Team, Your Organisation",
    "ask": [
        "Approve tagging to the US of the AC people who already spend most of their time on US work",
        "Agree the Global Model FY27 rate card as the basis for US-tagged hours",
        "Review the tagging every quarter as US volume grows, and add people as they cross the threshold",
    ],
    "winUS": [
        "Lower cost for the same people and the same quality: Global Model rate instead of sales price",
        "The yearly saving can fund extra capacity, such as an additional Associate",
        "A named, dedicated US team that knows US clients, frameworks and ways of working",
        "Hours and cost show up correctly in US reporting",
    ],
    "winAC": [
        "Stable, predictable US demand to plan hiring and skills around",
        "People are tagged to where their work actually is: cleaner utilisation and P&L",
        "Deeper US specialisation and career paths for the team",
        "A model we can repeat for UK, Canada, Australia and Japan",
    ],
    "steps": [
        ["Now", "Tag the people above the threshold to the US", "Confirm the rate card with Finance; update Home Sub LOS 2 for the agreed people"],
        ["Next quarter", "Review shared and growing profiles", "People whose US share passes the threshold move next; track the saving monthly"],
        ["FY28", "Dedicated US pod", "Move full-time US people onto dedicated seats (Cost/Person) where that is cheaper"],
    ],
}

DEFAULT_CONTENT = {'team': {'teamName': 'Your Team Name',
          'summary': 'Team supports global Deals through research, analysis, and due diligence across '
                     'Sustainability/ESG/Energy Transition/EUR/IUR topics',
          'leaders': [{'name': 'Leader One', 'role': 'Managing Director', 'note': 'Practice Leader'},
                      {'name': 'Leader Two', 'role': 'Director', 'note': ''},
                      {'name': 'Leader Three', 'role': 'Manager', 'note': 'Primary point of contact'}],
          'members': '6 Members',
          'keyContacts': [{'territory': 'Australia', 'name': 'Contact Name', 'role': 'Partner'},
                          {'territory': 'USA', 'name': 'Contact Name', 'role': 'Partner'},
                          {'territory': 'Japan', 'name': 'Contact Name', 'role': 'Partner'},
                          {'territory': 'Canada', 'name': 'Contact Name', 'role': 'Director'}],
          'currentServices': ['Company Profiles/Target Snapshot',
                              'Industry ESG trends and analysis',
                              'Industry ESG regulatory overview',
                              'Value chain analysis',
                              'Customer commitments and targets analysis',
                              'Comparable company list',
                              'ESG materiality assessment',
                              'ESG maturity assessment',
                              'ESG metrics benchmarking',
                              'Social media and reputational risk scan',
                              'Industry ESG trends repository',
                              'Industry ESG regulations repository',
                              'Drafting Management interview questions',
                              'Sentiment analysis',
                              'Virtual Data Room data review'],
          'futureServices': ['Drafting Request for information (RFI)',
                             'ESG metrics benchmark repository',
                             'ESG Framework & Disclosures repository',
                             'Red Flags Screening',
                             'ESG Due Diligence (first draft)',
                             'ESG Action Plan (first draft)',
                             'ESG Value Creation Plan',
                             'High level Climate Risk assessment',
                             'Thought Leadership Support',
                             'Proposal Support'],
          'frameworks': ['SASB Standards', 'MSCI - ESG Industry Materiality Map', 'PRI', 'CDP'],
          'databases': ['Pitchbook',
                        'Merger Market',
                        'S&P Capital IQ',
                        'Inframation',
                        'Statista',
                        'Technavio',
                        'Factiva',
                        'Bloomberg',
                        'IBIS World (US)',
                        'Debtwire',
                        'Lexis Nexis (only for US)'],
          'aiTools': ['Web Plot Digitizer',
                      'Copilot',
                      'Cursor',
                      'Claude Design',
                      'Internal GenAI assistant',
                      'ChatGPT Enterprise',
                      'Replit',
                      'Harvey',
                      'Tolt',
                      'Visual Studio Code']},
 'themes': {'title': 'Our Work Across Key Energy Transition Themes',
            'items': [{'title': 'Clean Energy Systems',
                       'tagline': 'Core technologies enabling low-carbon transitions',
                       'points': [['Hydrogen Mobility & Fuel',
                                   'Market trends analysis across Asia; vehicle adoption studies by global OEMs'],
                                  ['EV Ecosystem',
                                   'Market assessment for EV deals; EV charging infrastructure trends and policy '
                                   'mapping'],
                                  ['Battery Energy Storage (BESS)',
                                   'Research on Australian BESS projects (stage, size, distribution)'],
                                  ['Renewable Energy Assets',
                                   'Analysis of 600+ renewable projects (capex, equity ownership, conflict risks)'],
                                  ['Sustainable Aviation Fuel (SAF)',
                                   'R&D on biofuel-based jet fuel pathways vs conventional fuel']]},
                      {'title': 'Carbon Markets, Climate Governance & ESG Frameworks',
                       'tagline': 'Enabling credibility, compliance & climate governance maturity',
                       'points': [['TCFD Framework', 'Slide preparation, peer benchmarking & gap analysis'],
                                  ['Voluntary Carbon Markets (VCM)',
                                   'Countrywide policy analysis; comparison of IC-VCM/VCMI standards'],
                                  ['Biodiversity Credit Systems',
                                   'Research on nature-based credit mechanisms supporting net zero'],
                                  ['ESG Frameworks & Standards', 'Benchmarking ESG disclosures & commitments'],
                                  ['Policy & Regulatory Analysis',
                                   'Governance structures for renewable energy and carbon markets'],
                                  ['Carbon Capture & Storage (CCS)', 'Multi-country project assessments']]},
                      {'title': 'Decarbonisation Pathways & Sector Transitions',
                       'tagline': 'Supporting industries in developing credible net-zero strategies',
                       'points': [['Mine Decarbonisation', 'Benchmarking commitments; energy source transition plans'],
                                  ['Real Estate Transition', 'Mapping ESG KPIs & transition strategies'],
                                  ['Urban Regeneration Projects',
                                   'Analysis of low-carbon infrastructure options (EV charging, green mobility)'],
                                  ['Education Sector Decarbonisation',
                                   'Sustainability journey assessment & net-zero roadmap'],
                                  ['Industrial Emission Reduction',
                                   'Options assessment across infrastructure & built environment'],
                                  ['DERMs & VPP Engineering', 'Global competitor and business model analysis']]},
                      {'title': 'Circular Economy, Resource Efficiency & Environmental Solutions',
                       'tagline': 'Reducing waste, improving resource use & supporting nature-positive outcomes',
                       'points': [['Solar Panel Recycling', 'Market study of obsolete panel recycling in Australia'],
                                  ['Waste Management',
                                   'Company profiling, project analysis & government grant mapping'],
                                  ['Water & Wastewater Treatment',
                                   'Project identification & future potential assessment'],
                                  ['Critical Minerals', 'Screening key projects and gathering intelligence'],
                                  ['Environment, Conservation & Wildlife',
                                   'Strategic ESG guidance for environmental outcomes'],
                                  ['Investor Participation in Renewables',
                                   'Investor landscape analysis & investment outlook']]}]}}


SETUP_PLACEHOLDERS = {"your name", "you@example.com", "your team name", "your organisation", "your team, your organisation",
                      "leader one", "leader two", "leader three", "contact name"}
def read_setup(df):
    if isinstance(df, dict):
        return df
    if df is None or "Field" not in df.columns:
        return {}
    return {clean_str(f): clean_str(v) for sec, f, v in zip(df.get("Section", [""] * len(df)), df["Field"], df.get("Value", [""] * len(df)))
            if clean_str(f) and clean_str(sec).lower() != "also check"}


def apply_setup(payload, setup):
    """Setup tab values win over the other tabs; blank values leave things as they are."""
    global LOGO_FILE, LOGO_TEXT
    g = lambda k: setup.get(k, "")
    m, team, story = payload["meta"], payload["content"]["team"], payload["story"]
    if g("Report title"): m["title"] = g("Report title")
    if g("Classification"): m["classification"] = g("Classification")
    m["summaryTitle"] = g("Summary title") or m.get("summaryTitle") or "ESG Delivery: Summary"
    m["headerSubtitle"] = g("Header subtitle") or m.get("headerSubtitle") or "Sustainability & ESG Delivery"
    if g("Logo file"): LOGO_FILE = g("Logo file")
    if g("Logo text"): LOGO_TEXT = g("Logo text")
    people = [{"name": g(f"Contact name{s}"), "email": g(f"Contact email{s}")} for s in ("", " 2") if g(f"Contact name{s}") or g(f"Contact email{s}")]
    if people: m["contact"]["people"] = people
    for k, key in (("Team name", "teamName"), ("Team summary", "summary"), ("Members", "members")):
        if g(k): team[key] = g(k)
    leaders = [{"name": g(f"Leader {i} - Name"), "role": g(f"Leader {i} - Role"), "note": g(f"Leader {i} - Note")} for i in range(1, 6) if g(f"Leader {i} - Name")]
    if leaders: team["leaders"] = leaders
    contacts = [{"territory": t, "name": g(f"{t} - Name"), "role": g(f"{t} - Role")} for t in TERRITORY_ORDER if g(f"{t} - Name")]
    if contacts: team["keyContacts"] = contacts
    for k, key in (("Meeting", "meeting"), ("Story title", "title"), ("Story subtitle", "subtitle"), ("Presenter", "presenter")):
        if g(k): story[key] = g(k)
    todo = [k for k, v in setup.items() if v.lower() in SETUP_PLACEHOLDERS]
    if not g("Contact name") and not CONTACT_NAME: todo.append("Contact name")
    return todo


# ================================================================ read the data ===
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
              for name in (SHEET_TAG_MAP, SHEET_TEAM, SHEET_SERVICES, SHEET_THEMES, SHEET_STORY, SHEET_PRESETS, SHEET_SETUP, SHEET_TEAMSHEET)
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


def item_key(v):
    s = clean_str(v)
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def read_team_sheet(df, assum):
    """The Team Sheet -> tag map, page content, story text, presets, branding/contact values. Updates assum in place."""
    team = {"teamName": "", "summary": "", "members": "", "leaders": [], "keyContacts": [], "currentServices": [],
            "futureServices": [], "frameworks": [], "databases": [], "aiTools": []}
    themes = {"title": "Our Work Across Key Energy Transition Themes", "items": []}
    story, presets, setup, tmap, theme_map, todo = {}, [], {}, {}, {}, []
    lists = {"ask": "ask", "win for us": "winUS", "win for ac": "winAC"}
    tools = {"esg framework": "frameworks", "database": "databases", "ai tool": "aiTools", "digital & ai tool": "aiTools"}
    settings_num = {"fy start month": int, "scenario window (months)": int, "annual capacity hours": float}
    for _, r in df.iterrows():
        sec, item = clean_str(r.get("Section")).lower(), item_key(r.get("Item"))
        v, d1, d2, d3 = (clean_str(r.get(c)) for c in ("Value", "Detail 1", "Detail 2", "Detail 3"))
        if not sec or sec == "also check":
            continue
        for x in (v, d1, d2):
            if x.lower() in SETUP_PLACEHOLDERS:
                todo.append(f"{r.get('Section')}: {item}")
                break
        il = item.lower()
        if sec in ("branding", "contact"):
            if il == "energy transition page title":
                themes["title"] = v or themes["title"]
            else:
                setup[item] = v
        elif sec == "team":
            key = {"team name": "teamName", "team summary": "summary", "members": "members"}.get(il)
            if key: team[key] = v
        elif sec == "leader" and v:
            team["leaders"].append({"name": v, "role": d1, "note": d2})
        elif sec == "key contact" and v:
            team["keyContacts"].append({"territory": norm_territory(item), "name": v, "role": d1})
        elif sec == "service" and v:
            team["futureServices" if il.startswith("future") else "currentServices"].append(v)
        elif sec == "tool" and v and il in tools:
            team[tools[il]].append(v)
        elif sec == "story":
            if il in ("meeting", "title", "subtitle", "presenter"): story[il] = v
            elif il in lists and v: story.setdefault(lists[il], []).append(v)
            elif il == "next step" and (v or d1): story.setdefault("steps", []).append([v, d1, d2])
        elif sec == "et theme":
            theme_map.setdefault(item, {"title": "", "tagline": "", "points": []}).update(title=v, tagline=d1)
        elif sec == "et work area" and v:
            theme_map.setdefault(item, {"title": "", "tagline": "", "points": []})["points"].append([v, d1])
        elif sec == "setting":
            raw_v = r.get("Value")
            if il == "tagging threshold": assum["Tagging threshold"] = to_share(raw_v, assum.get("Tagging threshold"))
            elif il in settings_num and clean_str(raw_v):
                try: assum[item if item in DEFAULT_ASSUMPTIONS else item] = settings_num[il](float(raw_v))
                except (TypeError, ValueError): pass
            elif clean_str(raw_v):
                assum[item] = raw_v
        elif sec == "scenario preset" and item:
            presets.append({"name": item, "threshold": to_share(r.get("Value")),
                            "window": int(float(r.get("Detail 1"))) if clean_str(r.get("Detail 1")) else None,
                            "growth": to_share(r.get("Detail 2"), 0.0),
                            "method": "best" if "seat" in d3.lower() else "hourly"})
        elif sec == "tag mapping" and item:
            tmap[item.lower()] = norm_territory(v) if v else ""
    themes["items"] = [theme_map[k] for k in sorted(theme_map, key=lambda k: (len(k), k))]
    if setup.get("Report title"): assum["Report title"] = setup["Report title"]
    if setup.get("Classification"): assum["Classification"] = setup["Classification"]
    return {"tagMap": tmap, "content": {"team": team, "themes": themes}, "story": story,
            "presets": presets or None, "setup": setup, "todo": todo}


def read_inputs(inputs, assum):
    """Turn the input tabs into: tag map, page content, story text, scenario presets."""
    if SHEET_TEAMSHEET in inputs:                          # one sheet holds everything
        return read_team_sheet(inputs[SHEET_TEAMSHEET], assum)
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
        "storyLink": STORY_FILE,
    }
    meta["presets"] = [{**pr, "threshold": meta["threshold"] if pr["threshold"] is None else pr["threshold"],
                        "window": meta["window"] if pr["window"] is None else pr["window"]} for pr in (inp["presets"] or SCENARIO_PRESETS)]
    if not any(p["name"].lower() == "base" for p in meta["presets"]) and meta["presets"]:
        meta["presets"][0]["name"] = meta["presets"][0]["name"] or "Base"
    payload = {"meta": meta, "months": [str(p) for p in months], "territories": terrs,
               "staff": [smap[n] for n in staff_names], "teams": teams, "flags": flags, "categories": cats,
               "projects": projects, "rows": rows, "model": model_rates, "callouts": calls,
               "content": inp["content"], "storyInput": inp["story"],
               "setupInput": inp.get("setup") if "setup" in inp else (inputs or {}).get(SHEET_SETUP),   # read by apply_setup, removed before writing
               "todoInput": inp.get("todo")}

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



# ============================================================ input tabs in Excel ===
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



def guess_tag(tag):
    for pat, terr in TAG_KEYWORDS.items():
        if re.search(pat, tag, re.I):
            return terr
    return ""


def effective_config(path):
    """What the pages show today from the workbook's tabs (any older layout) or the built-in defaults."""
    raw, staff, model, callouts, assum, wrn, inputs = load(path)
    a = dict(assum)
    inp = read_inputs(inputs, a)
    content = json.loads(json.dumps(inp["content"] or DEFAULT_CONTENT))
    story = {**STORY, **{k: v for k, v in (inp["story"] or {}).items() if v}}
    presets = inp["presets"] or SCENARIO_PRESETS
    tags = sorted({clean_str(t) for t in staff["Home Sub LOS 2"] if clean_str(t)}) if "Home Sub LOS 2" in staff.columns else []
    meta = {"title": clean_str(a.get("Report title"), REPORT_TITLE), "classification": clean_str(a.get("Classification"), CLASSIFICATION),
            "contact": {"people": [{"name": n, "email": e} for n, e in ((CONTACT_NAME, CONTACT_EMAIL), (CONTACT_NAME_2, CONTACT_EMAIL_2)) if n or e]}}
    setup = read_setup(inputs.get(SHEET_SETUP))
    apply_setup({"meta": meta, "content": content, "story": story}, setup)
    return a, content, story, presets, inp["tagMap"], tags, setup, meta


def team_sheet_rows(path):
    a, content, story, presets, tagmap, tags, setup, meta = effective_config(path)
    t, th, R = content["team"], content["themes"], []
    def add(sec, item, val="", d1="", d2="", d3="", note=""):
        R.append([sec, item, val, d1, d2, d3, note])
    add("Branding", "Report title", meta["title"], note="Title at the top of the dashboard")
    add("Branding", "Summary title", meta.get("summaryTitle", "ESG Delivery: Summary"), note="Title of the one-page summary")
    add("Branding", "Header subtitle", meta.get("headerSubtitle", "Sustainability & ESG Delivery"), note="Next to the logo on the dashboard")
    add("Branding", "Classification", meta["classification"], note="Header and footer of every page, e.g. 'Company Confidential | Internal use only'")
    add("Branding", "Logo file", setup.get("Logo file") or LOGO_FILE, note="Logo file next to build.py, or a full path (PNG, JPG or SVG)")
    add("Branding", "Logo text", setup.get("Logo text") or LOGO_TEXT or "Your Organisation", note="Shown instead of the logo when the file is missing")
    add("Branding", "Energy Transition page title", th.get("title", ""), note="Heading of the Energy Transition tab")
    ppl = meta["contact"]["people"] + [{}, {}]
    add("Contact", "Contact name", ppl[0].get("name") or "Your Name", note="Last line of the pages: 'please reach out to ...'")
    add("Contact", "Contact email", ppl[0].get("email") or "you@example.com", note="Shown in brackets after the name")
    add("Contact", "Contact name 2", ppl[1].get("name", ""), note="Optional second contact")
    add("Contact", "Contact email 2", ppl[1].get("email", ""), note="Optional")
    add("Team", "Team name", t.get("teamName", ""), note="Who we are")
    add("Team", "Team summary", t.get("summary", ""), note="Quote under 'Meet the team'")
    add("Team", "Members", t.get("members", ""), note="e.g. '6 Members'")
    leaders = t.get("leaders", [])
    for i in range(max(5, len(leaders))):
        l = leaders[i] if i < len(leaders) else {}
        add("Leader", i + 1, l.get("name", ""), l.get("role", ""), l.get("note", ""),
            note="Value = name, Detail 1 = role, Detail 2 = note. Blank name = not shown" if i == 0 else "")
    have = {k["territory"]: k for k in t.get("keyContacts", [])}
    for terr in TERRITORY_ORDER + [x for x in have if x not in TERRITORY_ORDER]:
        k = have.get(terr, {})
        add("Key contact", terr, k.get("name", ""), k.get("role", ""), note="Value = name, Detail 1 = role. Blank = not shown" if terr == TERRITORY_ORDER[0] else "")
    for x in t.get("currentServices", []): add("Service", "Current", x)
    for x in t.get("futureServices", []): add("Service", "Future", x)
    for key, label in (("frameworks", "ESG framework"), ("databases", "Database"), ("aiTools", "AI tool")):
        for x in t.get(key, []): add("Tool", label, x)
    for k, label in (("meeting", "Meeting"), ("title", "Title"), ("subtitle", "Subtitle"), ("presenter", "Presenter")):
        add("Story", label, story.get(k, ""), note={"Presenter": "Cover and footer of the story"}.get(label, "Story cover"))
    for x in story.get("ask", []): add("Story", "Ask", x)
    for x in story.get("winUS", []): add("Story", "Win for US", x)
    for x in story.get("winAC", []): add("Story", "Win for AC", x)
    for w, h, d in story.get("steps", []): add("Story", "Next step", w, h, d, note="Value = when, Detail 1 = heading, Detail 2 = detail")
    for i, it in enumerate(th.get("items", []), 1):
        add("ET theme", i, it.get("title", ""), it.get("tagline", ""), note="Value = theme title, Detail 1 = tagline (up to 4 themes)")
        for wa, desc in it.get("points", []): add("ET work area", i, wa, desc, note="Item = theme number, Value = work area, Detail 1 = description")
    for k in ("FY start month", "Tagging threshold", "Scenario window (months)", "Annual capacity hours", "Training month"):
        add("Setting", k, a.get(k, DEFAULT_ASSUMPTIONS.get(k, "")), note={"FY start month": "7 = July (FY27 = Jul-26 to Jun-27); 1 = calendar year",
            "Tagging threshold": "Tag a person to a territory at this share of their hours", "Scenario window (months)": "Look-back for the tagging scenario",
            "Annual capacity hours": "Hours in one full-time seat", "Training month": "Pre/post training split"}[k])
    add("Setting", "Current tagging", a.get("Current tagging", DEFAULT_TAGGING_NOTE), note="Sentence shown in the assumptions on every page")
    for p in presets:
        add("Scenario preset", p["name"], p["threshold"] if p["threshold"] is not None else "", p["window"] if p["window"] is not None else "",
            p["growth"], "Hourly or seat" if p["method"] == "best" else "Hourly",
            note="Value = threshold (blank = Setting), Detail 1 = look-back months, Detail 2 = US volume change, Detail 3 = Hourly / Hourly or seat")
    for tag in tags:
        add("Tag mapping", tag, tagmap.get(tag.lower(), guess_tag(tag)), note="Territory this Home Sub LOS 2 tag means (blank = not a territory)")
    for tab, what in (("Global Model FY27 CPP", "Rate card: your real Rate/hr and Cost/Person by level"),
                      ("Staff Master", "People, levels, sales prices, Home Sub LOS 2 tags (your export)"),
                      ("Project Hours Raw Data", "Time entries (your export)"), ("Callouts", "Milestones on the monthly chart")):
        add("Also check", tab, note=what)
    return R


def add_team_sheet(wb, rows):
    ws = wb.create_sheet(SHEET_TEAMSHEET, 0)
    ws.append(TS_COLS)
    yellow, grey = PatternFill("solid", fgColor="FFF2CC"), PatternFill("solid", fgColor="F2F2F2")
    prev = None
    for r in rows:
        if prev is not None and r[0] != prev:
            ws.append([])                                          # blank line between sections
        ws.append(r); prev = r[0]
        row = ws[ws.max_row]
        row[0].font = Font(bold=True)
        row[6].font = Font(italic=True, color="7D7D7D")
        for c in row[2:6]:
            c.fill = grey if r[0] == "Also check" else yellow
            if clean_str(c.value).lower() in SETUP_PLACEHOLDERS:
                c.font = Font(bold=True, color="C00000")
    style_sheet(ws, [16, 26, 54, 34, 34, 18, 60], ORANGE)
    for row in ws.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
        if row[6].value: row[6].font = Font(italic=True, color="7D7D7D")
        for c in row[2:6]:
            if clean_str(c.value).lower() in SETUP_PLACEHOLDERS: c.font = Font(bold=True, color="C00000")
    ws.sheet_properties.tabColor = "FFC000"
    note = ws.cell(row=1, column=9, value="This is the only sheet you need to edit. Fill the yellow cells; red text = placeholder "
                                          "still to replace. Add rows anywhere (keep the Section name). Data tabs: raw hours, "
                                          "Staff Master, rate card, Callouts, WRN Category.")
    note.font = Font(italic=True, color="7D7D7D"); note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["I"].width = 46


def ensure_input_tabs(path):
    """Add the Team Sheet if missing (pre-filled from the workbook). Saves only then; keeps one backup."""
    wb = load_workbook(path)
    if SHEET_TEAMSHEET in wb.sheetnames:
        old = [x for x in OLD_INPUT_TABS if x in wb.sheetnames]
        print("Team Sheet: found" + (f" (older tabs no longer used, safe to delete: {', '.join(old)})" if old else ""))
        return
    add_team_sheet(wb, team_sheet_rows(path))
    backup = path.with_name(path.stem + " (backup before input tabs)" + path.suffix)
    if not backup.exists():
        shutil.copy2(path, backup)     # keep the original once, in case Excel features need restoring
    try:
        wb.save(path)
    except PermissionError:
        sys.exit(f"Could not save {path.name} - close it in Excel and run again.")
    old = [x for x in OLD_INPUT_TABS if x in wb.sheetnames]
    print(f"Team Sheet added, pre-filled from the workbook (original kept as '{backup.name}')." +
          (f" Older tabs no longer used, safe to delete: {', '.join(old)}" if old else ""))


# ================================================================== write pages ===
def load_templates():
    """template.html holds both pages and the chart library, separated by <!-- @@NAME --> markers."""
    path = HERE / TEMPLATE_FILE
    if not path.exists():
        sys.exit(f"{TEMPLATE_FILE} not found next to build.py.")
    parts = re.split(r"<!-- @@([A-Z]+)[^>]*-->\n?", path.read_text(encoding="utf-8"))
    found = {name: body for name, body in zip(parts[1::2], parts[2::2])}
    for need in ("DASHBOARD", "STORY", "CHARTJS"):
        if need not in found:
            sys.exit(f"{TEMPLATE_FILE} is missing the <!-- @@{need} --> section.")
    return found


def logo_html():
    logo = Path(LOGO_FILE) if Path(LOGO_FILE).is_absolute() else HERE / LOGO_FILE
    if not logo.exists():
        return f'<span class="wordmark">{LOGO_TEXT}</span>'
    mime = {".svg": "image/svg+xml", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}.get(logo.suffix.lower(), "image/png")
    return f'<img class="logo" alt="logo" src="data:{mime};base64,{base64.b64encode(logo.read_bytes()).decode()}">'


def write_page(template, chart_tag, payload, out):
    import html as _html
    sub = _html.escape(payload["meta"].get("headerSubtitle", ""))
    html = (template.replace("/*__CHARTJS__*/", chart_tag).replace("<!--__LOGO__-->", logo_html()).replace("<!--__SUBTITLE__-->", sub)
                    .replace("/*__DATA__*/null", json.dumps(payload, separators=(",", ":"))))
    out.write_text(html, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", default=EXCEL_FILE, help="Path to the hours workbook")
    ap.add_argument("--out-dir", default=str(HERE), help="Folder for the two HTML pages")
    ap.add_argument("--no-open", action="store_true", help="Do not open the story page at the end")
    args = ap.parse_args()
    excel = Path(args.excel).expanduser()
    excel = excel if excel.is_absolute() else HERE / excel
    if not excel.exists():
        sys.exit(f"Excel file not found: {excel}\nSet EXCEL_FILE at the top of build.py (or in local_settings.py).")
    out_dir = Path(args.out_dir).expanduser(); out_dir.mkdir(parents=True, exist_ok=True)
    tpl = load_templates()

    print(f"Workbook: {excel.name}\n\n[1/3] Checking the Team Sheet ...")
    ensure_input_tabs(excel)

    print("\n[2/3] Reading the data ...")
    payload, report = build_payload(*load(excel), source_name=excel.name)
    if payload["content"] is None:
        payload["content"] = DEFAULT_CONTENT
    payload["story"] = {**STORY, **{k: v for k, v in (payload.get("storyInput") or {}).items() if v}}
    payload["meta"].update(dashboard=DASHBOARD_FILE, today=datetime.now().strftime("%B %Y"),
                           summary=SUMMARY_FILE if "SUMMARY" in tpl else None)
    todo_sheet = payload.pop("todoInput", None)
    todo = apply_setup(payload, read_setup(payload.pop("setupInput", None)))
    if todo_sheet is not None:
        todo = todo_sheet
    print("Data check")
    for k, v in report.items():
        print(f"  {k:<45} {v}")
    print(f"  {'Team Sheet: still to fill':<45} {', '.join(todo) if todo else 'nothing - all placeholders replaced'}")

    print("\n[3/3] Writing the pages ...")
    chart_tag = "<script>" + tpl["CHARTJS"].strip() + "\n</script>"
    dash, story = out_dir / DASHBOARD_FILE, out_dir / STORY_FILE
    write_page(tpl["DASHBOARD"], chart_tag, payload, dash)
    write_page(tpl["STORY"], chart_tag, payload, story)
    if "SUMMARY" in tpl:
        summary = out_dir / SUMMARY_FILE
        write_page(tpl["SUMMARY"], chart_tag, payload, summary)
        print(f"  {summary}")
    print(f"  {dash}\n  {story}\n\nDone.")
    if OPEN_WHEN_DONE and not args.no_open:
        webbrowser.open(story.as_uri())


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print(f"\nA file is open in another program ({e.filename}). Close it (e.g. the workbook in Excel) and run again.")
    finally:
        if os.name == "nt" and "PROMPT" not in os.environ:   # double-clicked in Explorer: keep the window open
            input("\nPress Enter to close ...")
