"""
US Partner story (v2)
=====================
Builds a presentation-style page that tells one story to the US partners:
  US work done by AC is growing -> tag the US-focused AC people to the US ->
  the US pays Global Model rates instead of sales price -> over a year the gap pays for extra capacity.

Usage
-----
    python build_story_v2.py
    python build_story_v2.py --excel "D:\\Data\\Hours.xlsx" --out "D:\\Data\\US_Story.html"

It reads the same workbook as build_dashboard.py (and the same settings: Excel path, logo, team_content.json).
In the page: use the arrow keys (or Page Up / Page Down) to move between sections while presenting.
"""
import argparse
import base64
import json
import sys
from datetime import datetime
from pathlib import Path

import build_dashboard as bd

HERE = Path(__file__).resolve().parent
OUTPUT_PATH = HERE / "ESG_US_Story_v2.html"
DASHBOARD_FILE = "ESG_Hours_Dashboard.html"   # linked from the appendix (same folder)

# ------------------------------------------------------------------ STORY TEXT (edit freely) ---
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
# ------------------------------------------------------------------------------------------------


def logo_html():
    logo = HERE / bd.LOGO_FILE
    if not logo.exists():
        return f'<span class="wordmark">{bd.LOGO_TEXT}</span>'
    mime = {".svg": "image/svg+xml", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}.get(logo.suffix.lower(), "image/png")
    return f'<img class="logo" alt="logo" src="data:{mime};base64,{base64.b64encode(logo.read_bytes()).decode()}">'


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--excel", default=str(bd.EXCEL_PATH), help="Path to the hours workbook")
    ap.add_argument("--out", default=str(OUTPUT_PATH), help="Path of the HTML file to write")
    args = ap.parse_args()
    excel = Path(args.excel).expanduser()
    if not excel.exists():
        sys.exit(f"Excel file not found: {excel}")

    payload, report = bd.build_payload(*bd.load(excel), source_name=excel.name)
    content = HERE / bd.CONTENT_FILE
    if payload["content"] is None and content.exists():      # Excel 'Team' / 'Services' / 'ET Themes' tabs win
        payload["content"] = json.loads(content.read_text(encoding="utf-8"))
    payload["story"] = {**STORY, **{k: v for k, v in (payload.get("storyInput") or {}).items() if v}}   # Excel 'Story' tab wins
    payload["meta"]["dashboard"] = DASHBOARD_FILE
    payload["meta"]["today"] = datetime.now().strftime("%B %Y")

    chart_js = HERE / "chart.umd.min.js"
    chart_tag = ("<script>" + chart_js.read_text(encoding="utf-8") + "</script>") if chart_js.exists() else \
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>'

    html = (TEMPLATE.replace("/*__CHARTJS__*/", chart_tag).replace("<!--__LOGO__-->", logo_html())
                    .replace("/*__DATA__*/null", json.dumps(payload, separators=(",", ":"))))
    out = Path(args.out).expanduser()
    out.write_text(html, encoding="utf-8")
    print(f"rows used: {report['rows used']}, total hours: {report['total hours']}")
    print(f"Story page written to: {out}")


TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>US Tagging Proposal</title>
<style>
:root{
  color-scheme:light;
  --orange:#D04A02; --tangerine:#EB8C00; --yellow:#FFB600; --rose:#DB536A; --burgundy:#A32020; --charcoal:#2D2D2D;
  --bg:#FFFFFF; --soft:#F6F4F2; --peach:#FDEDE4; --border:#DEDEDE; --grid:#EBEBEB;
  --text:#2D2D2D; --text-2:#5F5F5F; --text-3:#7D7D7D; --good:#2E7D32;
  --c-canada:#464646; --c-usa:#D04A02; --c-australia:#FFB600; --c-uk:#DB536A; --c-japan:#A32020;
  --neutral:#BDBDBD; --head:Georgia,"Times New Roman",serif;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth;scroll-padding-top:58px}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.5 Arial,Helvetica,sans-serif}
/* top bar */
.topbar{position:sticky;top:0;z-index:10;background:#fff;border-bottom:1px solid var(--border);box-shadow:0 1px 0 rgba(0,0,0,.02)}
.topbar .in{max-width:1280px;margin:0 auto;padding:8px 32px;display:flex;align-items:center;gap:18px}
.topbar .logo{height:30px}
.wordmark{font:700 24px var(--head)}
.chapters{display:flex;gap:2px;overflow-x:auto;margin-left:auto;scrollbar-width:none}
.chapters a{white-space:nowrap;text-decoration:none;color:var(--text-3);font-size:12.5px;padding:8px 10px;border-bottom:2px solid transparent}
.chapters a.on{color:var(--text);border-bottom-color:var(--orange);font-weight:700}
.dashlink{white-space:nowrap;text-decoration:none;font-weight:700;font-size:13px;color:var(--orange);border:1.5px solid var(--orange);border-radius:4px;padding:6px 12px;margin-left:8px}
.dashlink:hover{background:var(--orange);color:#fff}
.progress{position:absolute;left:0;bottom:-1px;height:3px;background:var(--orange);width:0}
/* sections */
section{min-height:calc(100vh - 50px);padding:56px 32px 64px;border-bottom:1px solid var(--grid);display:flex;align-items:center}
section:nth-of-type(even){background:var(--soft)}
.wrap{max-width:1216px;margin:0 auto;width:100%;min-width:0}
.kicker{font-size:12.5px;letter-spacing:.12em;text-transform:uppercase;color:var(--orange);font-weight:700}
h1{font:700 52px/1.08 var(--head);margin:14px 0 14px;letter-spacing:-.01em}
h2{font:700 34px/1.15 var(--head);margin:8px 0 10px;letter-spacing:-.01em;max-width:960px}
h3{font:700 18px/1.3 var(--head);margin:0 0 8px}
.lead{font-size:18px;color:var(--text-2);max-width:880px;margin:0 0 28px}
.lead b{color:var(--text)}
.sowhat{margin-top:24px;border-left:4px solid var(--orange);background:#fff;padding:14px 18px;font-size:16px;max-width:980px}
section:nth-of-type(odd) .sowhat{background:var(--peach)}
.sowhat b{color:var(--orange)}
.grid{display:grid;gap:20px}
.g2{grid-template-columns:1fr 1fr} .g3{grid-template-columns:repeat(3,1fr)} .g4{grid-template-columns:repeat(4,1fr)} .g32{grid-template-columns:3fr 2fr}
.card{background:#fff;border:1px solid var(--border);border-radius:4px;padding:20px 22px;min-width:0}
.stat{border-top:4px solid var(--orange);background:#fff;border-left:1px solid var(--border);border-right:1px solid var(--border);border-bottom:1px solid var(--border);border-radius:0 0 4px 4px;padding:18px 20px}
.stat .v{font:700 40px/1.1 var(--head);color:var(--text)}
.stat .v.o{color:var(--orange)}
.stat .l{font-size:14px;color:var(--text-2);margin-top:6px}
.chart{position:relative;height:340px}
.note{font-size:12.5px;color:var(--text-3)}
/* cover */
#cover{background:linear-gradient(180deg,#fff 0%,#fff 62%,var(--soft) 62%);}
.cover-logo img{height:64px} .cover-logo .wordmark{font-size:44px}
.cover-meta{margin-top:18px;color:var(--text-2);font-size:14px}
.cover-meta b{color:var(--text)}
.rule{width:96px;height:6px;background:var(--orange);margin:28px 0 22px}
/* lists */
.ticks{list-style:none;margin:0;padding:0;display:grid;gap:10px}
.ticks li{padding-left:24px;position:relative}
.ticks li::before{content:"";position:absolute;left:3px;top:8px;width:8px;height:8px;background:var(--orange);transform:rotate(45deg)}
.num-list{counter-reset:n;list-style:none;margin:0;padding:0;display:grid;gap:14px}
.num-list li{counter-increment:n;padding-left:46px;position:relative;font-size:16px}
.num-list li::before{content:counter(n);position:absolute;left:0;top:-3px;width:32px;height:32px;border-radius:50%;background:var(--orange);color:#fff;display:grid;place-items:center;font:700 15px Arial}
/* team */
.leaders{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px}
.leader{border-left:3px solid var(--orange);padding:2px 0 2px 14px}
.leader .nm{font-weight:700;color:var(--orange)} .leader .rl{font-weight:700;font-size:13.5px} .leader .nt{font-size:13px;color:var(--text-2)}
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chips span{border:1px solid var(--border);border-radius:3px;padding:4px 10px;font-size:13px;background:var(--soft)}
.contacts{display:grid;grid-template-columns:repeat(2,1fr);gap:8px 20px;font-size:14px}
.contacts .tt{font-weight:700}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;vertical-align:0}
/* tables */
table{width:100%;border-collapse:collapse;font-size:14px;font-variant-numeric:tabular-nums}
th,td{padding:9px 10px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap}
th{background:var(--charcoal);color:#fff;font-size:12.5px;font-weight:700}
td.n,th.n{text-align:right}
tr.hi td{background:var(--peach)}
tr.tot td{font-weight:700;border-top:2px solid var(--charcoal)}
.mix{display:flex;height:12px;width:170px;border-radius:2px;overflow:hidden;gap:1px}
.pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:12px;border:1px solid var(--border);white-space:nowrap}
.pill.go{border-color:var(--orange);color:var(--orange);font-weight:700}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:12.5px;color:var(--text-2);margin:0 0 10px}
.tw{overflow-x:auto}
/* calculator */
.calc{display:grid;grid-template-columns:340px 1fr;gap:28px;align-items:start}
.calc>*,.ctl>*{min-width:0}
.track>.seg-b{min-width:0;flex-shrink:1}
.ctl{display:grid;gap:18px}
.ctl label{font-size:13px;font-weight:700;display:flex;justify-content:space-between}
.ctl label span{font:700 17px var(--head);color:var(--orange)}
input[type=range]{width:100%;accent-color:var(--orange)}
.seg{display:flex;border:1px solid var(--border);border-radius:4px;overflow:hidden}
.seg button{flex:1;background:#fff;border:0;padding:8px 6px;font:inherit;font-size:13.5px;color:var(--text-2);cursor:pointer}
.seg button+button{border-left:1px solid var(--border)}
.seg button.on{background:var(--orange);color:#fff;font-weight:700}
.money{display:flex;align-items:center;border:1px solid var(--border);border-radius:4px;background:#fff;width:140px}
.money span{padding-left:10px;color:var(--text-3)}
.money input{border:0;width:100%;padding:8px;font:inherit;text-align:right}
.bars{display:grid;gap:22px;grid-template-columns:minmax(0,1fr)}
.barrow .bl{display:flex;justify-content:space-between;font-size:14px;margin-bottom:6px}
.barrow .bl b{font:700 18px var(--head)}
.track{display:flex;height:54px;gap:2px}
.seg-b{display:flex;align-items:center;justify-content:center;color:#fff;font-size:12.5px;font-weight:700;overflow:hidden;white-space:nowrap;padding:0 6px;border-radius:2px;transition:width .35s}
.seg-b.today{background:var(--charcoal)} .seg-b.mgr{background:var(--orange)} .seg-b.assoc{background:var(--tangerine)}
.seg-b.save{background:repeating-linear-gradient(45deg,#E9F3EA,#E9F3EA 6px,#D5E8D7 6px,#D5E8D7 12px);color:var(--good);border:1px dashed var(--good)}
.verdict{font:700 24px/1.3 var(--head);margin:6px 0 0}
.verdict b{color:var(--orange)}
/* steps */
.steps{display:grid;grid-template-columns:repeat(3,1fr);gap:0;position:relative}
.step{padding:22px 24px 22px 0;position:relative}
.step::before{content:"";position:absolute;top:35px;left:28px;right:0;height:2px;background:var(--orange)}
.step:last-child::before{display:none}
.dot{width:28px;height:28px;border-radius:50%;background:var(--orange);position:relative;z-index:1;margin-bottom:14px}
.step .when{font-size:12.5px;letter-spacing:.08em;text-transform:uppercase;color:var(--orange);font-weight:700}
.ask{background:var(--charcoal);color:#fff;border-radius:4px;padding:24px 28px;margin-top:24px}
.ask h3{color:#fff}
.ask .num-list li{font-size:15.5px}
.win h3{display:flex;align-items:center;gap:10px}
.win .badge{background:var(--orange);color:#fff;font:700 12px Arial;padding:3px 9px;border-radius:3px;letter-spacing:.04em}
.plus{display:grid;place-items:center;font:700 42px var(--head);color:var(--orange)}
.appx li{margin-bottom:6px;color:var(--text-2)}
.appx b{color:var(--text)}
.notice{border:1px dashed var(--orange);background:var(--peach);padding:16px 20px;border-radius:4px;font-size:15.5px;max-width:980px;margin-top:8px}
.btn{display:inline-block;background:var(--orange);color:#fff;text-decoration:none;padding:10px 18px;border-radius:4px;font-weight:700}
footer{max-width:1280px;margin:0 auto;padding:18px 32px;color:var(--text-3);font-size:12px}
.legalrow{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}
footer b{color:var(--text)}
.contactline{margin:6px 0 0;color:var(--text-3);font-size:12px}
.contactline a{color:inherit}
/* layout on narrow screens + print (restored) */
@media (max-width:1000px){.calc{grid-template-columns:minmax(0,1fr) !important} .g2,.g3,.g4,.g32,.leaders,.steps{grid-template-columns:1fr} h1{font-size:38px} h2{font-size:28px} section{min-height:auto} .step::before{display:none} .chapters{display:none} #winwin .grid{grid-template-columns:1fr !important} .plus{display:none}}
@media print{.topbar{position:static} section{min-height:auto;page-break-after:always} .chapters,.dashlink,.progress{display:none}}
/* motion: content fades up as each section comes into view */
.reveal{opacity:0;transform:translateY(22px);transition:opacity .7s ease,transform .7s cubic-bezier(.2,.7,.2,1)}
.reveal.in{opacity:1;transform:none}
.stat,.card{transition:transform .25s ease,box-shadow .25s ease}
.stat:hover{transform:translateY(-3px);box-shadow:0 6px 18px rgba(0,0,0,.07)}
@media (prefers-reduced-motion:reduce){.reveal{opacity:1;transform:none;transition:none} html{scroll-behavior:auto} .stat:hover{transform:none}}
@media print{.reveal{opacity:1 !important;transform:none !important}}
</style>
/*__CHARTJS__*/
</head>
<body>
<div class="topbar"><div class="in"><!--__LOGO__--><nav class="chapters" id="chapters"></nav><a class="dashlink" id="dashLink" href="#">Full dashboard &#8594;</a></div><div class="progress" id="progress"></div></div>

<section id="cover" data-title="Cover"><div class="wrap">
  <div class="cover-logo"><!--__LOGO__--></div>
  <div class="rule"></div>
  <div class="kicker" id="cvMeeting"></div>
  <h1 id="cvTitle"></h1>
  <p class="lead" id="cvSub"></p>
  <div class="grid g3" id="cvStats" style="margin-top:36px"></div>
  <div class="cover-meta" id="cvMeta"></div>
</div></section>

<section id="summary" data-title="Summary"><div class="wrap">
  <div class="kicker">Executive summary</div>
  <h2>US work at AC is growing fast, and the US is paying more for it than it needs to</h2>
  <div class="grid g32" style="margin-top:20px">
    <div class="card"><ol class="num-list" id="smPoints"></ol></div>
    <div class="ask" style="margin-top:0"><h3>What we are asking for</h3><ol class="num-list" id="smAsk"></ol></div>
  </div>
</div></section>

<section id="team" data-title="Who we are"><div class="wrap">
  <div class="kicker">Who we are</div>
  <h2 id="tmTitle"></h2>
  <p class="lead" id="tmSummary"></p>
  <div class="grid g32">
    <div class="card">
      <div class="leaders" id="tmLeaders"></div>
      <h3 style="margin-top:6px">What we do for Deals today</h3>
      <div class="chips" id="tmServices"></div>
    </div>
    <div class="card">
      <h3>Territory key contacts</h3>
      <div class="contacts" id="tmContacts"></div>
      <h3 style="margin-top:18px">Tools and databases</h3>
      <div class="chips" id="tmTools"></div>
    </div>
  </div>
</div></section>

<section id="growth" data-title="US demand"><div class="wrap">
  <div class="kicker">1 · US demand is growing</div>
  <h2 id="grTitle"></h2>
  <p class="lead" id="grLead"></p>
  <div class="grid g32">
    <div class="card"><h3>US hours delivered by AC, per month</h3><div class="chart"><canvas id="cUs"></canvas></div></div>
    <div class="grid" id="grStats" style="align-content:start"></div>
  </div>
  <div class="sowhat" id="grSo"></div>
</div></section>

<section id="people" data-title="Who does it"><div class="wrap">
  <div class="kicker">2 · The US work sits with a few people</div>
  <h2 id="ppTitle"></h2>
  <p class="lead" id="ppLead"></p>
  <div class="card"><div class="legend" id="ppLegend"></div><div class="tw" id="ppTable"></div></div>
  <div class="sowhat" id="ppSo"></div>
</div></section>

<section id="cost" data-title="Cost today vs tagged"><div class="wrap">
  <div class="kicker">3 · What the US pays today vs once tagged</div>
  <h2 id="csTitle"></h2>
  <p class="lead" id="csLead"></p>
  <div class="grid g32">
    <div class="card"><h3>US cost per person, per year</h3><div class="chart"><canvas id="cCost"></canvas></div></div>
    <div class="grid" id="csStats" style="align-content:start"></div>
  </div>
  <div class="sowhat" id="csSo"></div>
</div></section>

<section id="free" data-title="A free Associate"><div class="wrap">
  <div class="kicker">4 · Same budget, more capacity (illustration for one full-time US person)</div>
  <h2>A dedicated US Manager plus an extra Associate can cost less than the Manager alone costs today</h2>
  <p class="lead">Section 3 used today's part-time US hours. This shows where it leads once someone works full-time for the US. Pick a level and how many hours a year they work for the US. The bars compare what the US pays today with a dedicated US seat at the Global Model FY27 Cost/Person, and show how many extra Associates the difference pays for.</p>
  <div class="card calc">
    <div class="ctl">
      <div><label>Level of the US-dedicated person</label><div class="seg" id="fcLevel"></div></div>
      <div><label>US hours per year (one person) <span id="fcHrsV"></span></label><input type="range" id="fcHrs" min="400" max="1800" step="50"></div>
      <div><label>Today's sales price per hour</label><div class="money"><span>$</span><input type="number" id="fcRate" min="0" step="1"></div><div class="note" style="margin-top:4px">Pre-filled with this level's average current price</div></div>
      <div class="note" id="fcNote"></div>
    </div>
    <div>
      <div class="bars">
        <div class="barrow"><div class="bl"><span>Today: hours × sales price</span><b id="fcToday"></b></div><div class="track" id="fcBarToday"></div></div>
        <div class="barrow"><div class="bl"><span>US-tagged: dedicated seat + extra Associates</span><b id="fcTagged"></b></div><div class="track" id="fcBarTagged"></div></div>
      </div>
      <p class="verdict" id="fcVerdict"></p>
    </div>
  </div>
  <div class="sowhat" id="fcSo"></div>
</div></section>

<section id="longterm" data-title="Long term"><div class="wrap">
  <div class="kicker">5 · The saving grows with US volume</div>
  <h2 id="ltTitle"></h2>
  <p class="lead" id="ltLead"></p>
  <div class="card"><h3>Annual saving for the US as US hours grow</h3><div class="chart"><canvas id="cLong"></canvas></div><div class="note" id="ltNote" style="margin-top:8px"></div></div>
  <div class="sowhat" id="ltSo"></div>
</div></section>

<section id="winwin" data-title="Win-win"><div class="wrap">
  <div class="kicker">6 · Why it works for both sides</div>
  <h2>A win-win: lower cost for the US, a stable dedicated bench for AC</h2>
  <div class="grid" style="grid-template-columns:1fr 60px 1fr;margin-top:20px">
    <div class="card win"><h3><span class="badge">US</span> What the US practice gains</h3><ul class="ticks" id="wwUS"></ul></div>
    <div class="plus">+</div>
    <div class="card win"><h3><span class="badge">AC</span> What the Acceleration Center gains</h3><ul class="ticks" id="wwAC"></ul></div>
  </div>
</div></section>

<section id="next" data-title="Next steps"><div class="wrap">
  <div class="kicker">7 · Proposed next steps</div>
  <h2>Start with the people who are already working for the US, then grow with demand</h2>
  <div class="steps" id="nsSteps" style="margin-top:20px"></div>
  <div class="ask"><h3>Decisions we need today</h3><ol class="num-list" id="nsAsk"></ol></div>
</div></section>

<section id="appendix" data-title="Appendix"><div class="wrap">
  <div class="kicker">Appendix</div>
  <h2>How the numbers are worked out</h2>
  <div class="grid g2" style="margin-top:16px">
    <div class="card"><ul class="appx" id="apMethod"></ul></div>
    <div class="card"><h3>Rate card used (Global Model FY27)</h3><div class="tw" id="apRates"></div>
      <p style="margin-top:18px">The full interactive dashboard (territories, growth, rates and the scenario planner) has all the detail behind these pages.</p>
      <a class="btn" id="apLink" href="#">Open the full dashboard</a></div>
  </div>
</div></section>

<footer><div class="legalrow"><span id="ftLeft"></span><span id="ftRight"></span></div><div class="contactline" id="signText"></div></footer>

<script>
const D = /*__DATA__*/null;
const ST = D.story, C = D.content || {}, MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const M = D.months, T = D.territories, S = D.staff, R = D.rows, NM = M.length, LAST = NM - 1, US = T.indexOf("USA");
const [ri_m, ri_t, ri_s, ri_p, ri_team, ri_h] = [0, 1, 2, 3, 4, 5];
const el = id => document.getElementById(id);
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const ym = i => M[i].split("-").map(Number);
const mLabel = i => { if (i == null || i < 0 || i >= NM) return "–"; const [y, m] = ym(i); return MON[m - 1] + "-" + String(y).slice(2); };
const cy = i => ym(i)[0];
const fy = i => { const [y, m] = ym(i); return D.meta.fyStart === 1 ? y : (m >= D.meta.fyStart ? y + 1 : y); };
const fH = v => Math.round(v).toLocaleString("en-US");
const fUSD = v => "$" + Math.round(v).toLocaleString("en-US");
const fK = v => Math.abs(v) >= 1e6 ? "$" + (v / 1e6).toFixed(2) + "M" : Math.abs(v) >= 1e3 ? "$" + (v / 1e3).toFixed(1) + "K" : fUSD(v);
const pct = v => v == null || !isFinite(v) ? "–" : (v * 100).toFixed(0) + "%";
const esc = v => String(v ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
const inWin = (i, n, end = LAST) => i > end - n && i <= end;
const TVAR = {canada: "--c-canada", usa: "--c-usa", australia: "--c-australia", uk: "--c-uk", japan: "--c-japan"};
const tColor = i => css(TVAR[String(T[i]).toLowerCase()] || "--text-3");
const rateOf = (s, year) => { const rs = S[s].rates, ys = Object.keys(rs).map(Number).sort((a, b) => a - b); if (!ys.length) return 0;
  if (rs[year] != null) return rs[year]; return year < ys[0] ? rs[ys[0]] : rs[ys.filter(y => y <= year).pop()]; };
const MODEL = D.model, CAP = D.meta.capacity || 1800, THR = D.meta.threshold, WIN = D.meta.window, ANN = 12 / WIN;
const ASSOC = MODEL.A || Object.values(MODEL).sort((a, b) => a.cpp - b.cpp)[0];

// ---------------- core numbers
const usM = new Array(NM).fill(0), allM = new Array(NM).fill(0);
for (const r of R) { allM[r[ri_m]] += r[ri_h]; if (r[ri_t] === US) usM[r[ri_m]] += r[ri_h]; }
const sumWin = (arr, n, end = LAST) => arr.reduce((s, v, i) => s + (inWin(i, n, end) ? v : 0), 0);
const usLTM = sumWin(usM, 12), usPrev = sumWin(usM, 12, LAST - 12), allLTM = sumWin(allM, 12);
const usFirst = usM.findIndex(v => v > 0);
const usLast3 = sumWin(usM, 3) / 3, usFirst3 = usM.slice(usFirst, usFirst + 3).reduce((a, b) => a + b, 0) / 3;

const ppl = S.map((s, i) => ({s, i, h: T.map(() => 0), tot: 0, usCur: 0}));
for (const r of R) {
  if (!inWin(r[ri_m], WIN)) continue;
  const p = ppl[r[ri_s]]; p.h[r[ri_t]] += r[ri_h]; p.tot += r[ri_h];
  if (r[ri_t] === US) p.usCur += r[ri_h] * rateOf(r[ri_s], cy(r[ri_m]));
}
const rows = ppl.filter(p => p.tot > 0).map(p => {
  const usH = p.h[US] || 0, share = p.tot ? usH / p.tot : 0, rc = MODEL[p.s.level];
  const already = p.s.tagTerritory === "USA";
  const move = share >= THR - 1e-9 && !already && !!rc;
  const annH = usH * ANN, today = p.usCur * ANN, tagged = rc ? annH * rc.rate : today;
  return {...p, usH, share, rc, already, move, annH, today, tagged, avg: usH ? p.usCur / usH : 0};
}).sort((a, b) => b.share - a.share);
const movers = rows.filter(r => r.move);
const usWorkers = rows.filter(r => r.usH > 0);
const tot = f => movers.reduce((s, r) => s + f(r), 0);
const today = tot(r => r.today), tagged = tot(r => r.tagged), saving = today - tagged, moverHrs = tot(r => r.annH);
// data situations the story has to handle honestly
const NOUS = US < 0 || usLTM === 0;                    // no US hours in the last 12 months
const NOMOVERS = !NOUS && movers.length === 0;          // nobody passes the threshold (or all already US-tagged)
const NOPRICE = movers.length > 0 && today === 0;      // sales prices missing in Staff Master
const MGR = MODEL.M ? "M" : Object.keys(MODEL).find(l => l !== "A") || Object.keys(MODEL)[0];
function notice(sec, title, html) {   // replace a section's body with an explanation
  const w = document.querySelector(`#${sec} .wrap`); if (!w) return;
  [...w.children].forEach(c => { if (!c.classList.contains("kicker")) c.remove(); });
  w.insertAdjacentHTML("beforeend", `<h2>${title}</h2><div class="notice">${html}</div>`);
}
const NOUS_MSG = "The data has no US hours in the last 12 months, so there is no US work to tag yet. Check that the US territory is spelled USA, US or United States in 'Territory Name'.";
const NOMOVERS_MSG = `No one who works for the US is above the ${Math.round(THR * 100)}% threshold and still on a non-US tag. Either they are already US-tagged, or US work is spread thinly. Lower the 'Tagging threshold' in the Assumptions tab to see who comes closest.`;
const NOPRICE_MSG = "Staff Master has no sales prices, so today's cost (and the saving) cannot be calculated. Add a 'Sales Price' column (or 'Sales Price 2025', 'Sales Price 2026'...) to Staff Master and rebuild.";

// long-term: saving at US volume multiplier m (hourly, and cheapest of hourly or dedicated seats)
function savingAt(m, seats) {
  return movers.reduce((s, r) => {
    const h = r.annH * m, cur = h * r.avg, hourly = h * r.rc.rate;
    const full = Math.floor(h / CAP), restH = h - full * CAP, seat = full * r.rc.cpp + Math.min(restH * r.rc.rate, r.rc.cpp);   // full-time on seats, overflow hourly
    return s + cur - (seats ? Math.min(hourly, seat) : hourly);
  }, 0);
}

// ---------------- chapter nav, progress, keyboard
const secs = [...document.querySelectorAll("section")];
el("chapters").innerHTML = secs.map(s => `<a href="#${s.id}" data-id="${s.id}">${s.dataset.title}</a>`).join("");
const io = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) document.querySelectorAll(".chapters a").forEach(a => a.classList.toggle("on", a.dataset.id === e.target.id)); }), {rootMargin: "-45% 0px -50% 0px"});
secs.forEach(s => io.observe(s));
addEventListener("scroll", () => { const h = document.documentElement; el("progress").style.width = (h.scrollTop / (h.scrollHeight - h.clientHeight) * 100) + "%"; });
addEventListener("keydown", e => {
  if (["INPUT", "TEXTAREA"].includes(document.activeElement.tagName)) return;
  const k = e.key, dir = ["ArrowRight", "ArrowDown", "PageDown", " "].includes(k) ? 1 : ["ArrowLeft", "ArrowUp", "PageUp"].includes(k) ? -1 : 0;
  if (!dir) return; e.preventDefault();
  const y = scrollY + 60, cur = secs.findIndex((s, i) => s.offsetTop <= y && (i === secs.length - 1 || secs[i + 1].offsetTop > y));
  const nx = secs[Math.min(secs.length - 1, Math.max(0, cur + dir))]; nx.scrollIntoView({behavior: "smooth"});
});


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

// ---------------- charts
// charts are created when they scroll into view, so their entry animation is seen
function lazyChart(canvas, cfg) {
  const io = new IntersectionObserver(es => { if (es.some(e => e.isIntersecting)) { io.disconnect(); new Chart(canvas, cfg); } }, {threshold: .25});
  io.observe(canvas);
}
Chart.defaults.font.family = "Arial, Helvetica, sans-serif"; Chart.defaults.color = css("--text-2"); Chart.defaults.animation.duration = REDUCED ? 0 : 1100; Chart.defaults.animation.easing = "easeOutQuart";
const axes = money => ({
  x: {grid: {display: false}, ticks: {maxRotation: 0, autoSkipPadding: 14}, border: {color: css("--border")}},
  y: {beginAtZero: true, grid: {color: css("--grid"), drawTicks: false}, border: {display: false},
      ticks: {padding: 6, callback: v => money ? (v >= 1000 ? "$" + (v / 1000) + "K" : "$" + v) : v.toLocaleString("en-US")}}});
const legend = {position: "top", align: "start", labels: {boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 2}};
const tip = money => ({callbacks: {label: c => ` ${c.dataset.label}: ${money ? fUSD(c.raw) : fH(c.raw)}`}});

// ================= render
function cover() {
  el("cvMeeting").textContent = ST.meeting; el("cvTitle").textContent = ST.title; el("cvSub").textContent = ST.subtitle;
  const best = freeCalc(MGR, CAP, avgRate(MGR));
  el("cvStats").innerHTML = [
    usPrev ? ["+" + pct(usLTM / usPrev - 1), "growth in US hours delivered by AC, last 12 months vs the 12 before"] : [fH(usLTM), "US hours delivered by AC in the last 12 months"],
    NOUS || NOMOVERS ? ["–", NOUS ? "no US hours in the data yet" : "no one to tag yet at the current threshold"] : NOPRICE ? ["–", "add sales prices to Staff Master to calculate the saving"]
      : [fK(saving), `estimated yearly saving for the US by tagging ${movers.length} ${movers.length === 1 ? "person who already works" : "people who already work"} mainly for the US`],
    [best.assoc >= 1 ? `+${best.assoc} Associate${best.assoc > 1 ? "s" : ""}` : fK(best.left), best.assoc >= 1 ? `on top of a full-time US ${MODEL[MGR].name}, for less than that ${MODEL[MGR].name} costs today (illustration, ${fH(CAP)} hrs a year)` : `yearly saving on one full-time US ${MODEL[MGR].name} seat (illustration, ${fH(CAP)} hrs a year)`],
  ].map(([v, l], k) => `<div class="stat"><div class="v ${k ? "o" : ""}">${v}</div><div class="l">${l}</div></div>`).join("");
  el("cvMeta").innerHTML = `<b>${esc(ST.presenter)}</b> · ${D.meta.today} · Data ${mLabel(0)} to ${mLabel(LAST)}`;
}

function summary() {
  const best = freeCalc(MGR, CAP, avgRate(MGR));
  el("smPoints").innerHTML = [
    NOUS ? `<b>There are no US hours in the data yet.</b> ${NOUS_MSG}` :
    `<b>US demand is ${usPrev && usLTM > usPrev ? "growing fast" : "a steady part of our work"}.</b> ${usFirst > 0 ? `US hours went from zero before ${mLabel(usFirst)} to` : "We delivered"} ${fH(usLTM)} US hours in the last 12 months, ${pct(usLTM / allLTM)} of everything we deliver.`,
    NOUS ? "" : NOMOVERS ? `<b>No one is ready to tag yet.</b> ${NOMOVERS_MSG}` :
    `<b>${movers.length} ${movers.length === 1 ? "person already spends" : "people already spend"} at least ${pct(THR)} of their time on US work</b>, but sit${movers.length === 1 ? "s" : ""} under a non-US tag, so the US pays their sales price.`,
    NOUS || NOMOVERS ? "" : NOPRICE ? `<b>The saving cannot be calculated yet.</b> ${NOPRICE_MSG}` :
    `<b>Tagging them to the US cuts the US cost by about ${fK(saving)} a year (${pct(saving / today)})</b> on today's volume, at Global Model FY27 rates.`,
    best.assoc >= 1 ? `<b>As people become full-time on US work, the gap grows: for one full-time US ${MODEL[MGR].name} it is big enough to fund ${best.assoc === 1 ? "an extra Associate" : best.assoc + " extra Associates"}</b>, and the total is still ${fK(best.rest)} below what the ${MODEL[MGR].name} alone costs today.`
                    : best.left > 0 ? `<b>For a full-time US ${MODEL[MGR].name}, a dedicated seat saves ${fK(best.left)} a year</b> against today's hourly cost.` : "",
  ].filter(Boolean).map(x => `<li>${x}</li>`).join("");
  el("smAsk").innerHTML = ST.ask.map(x => `<li>${esc(x)}</li>`).join("");
}

function team() {
  const t = C.team; if (!t) { el("team").remove(); document.querySelector('.chapters a[data-id="team"]')?.remove(); return; }
  el("tmTitle").textContent = `${t.teamName || "Our team"}${t.members ? ": " + t.members.toLowerCase() : ""} supporting Deals across ${T.length} territories`;
  el("tmSummary").textContent = t.summary;
  el("tmLeaders").innerHTML = t.leaders.map(l => `<div class="leader"><div class="nm">${esc(l.name)}</div><div class="rl">${esc(l.role)}</div><div class="nt">${esc(l.note || "")}</div></div>`).join("");
  el("tmServices").innerHTML = t.currentServices.slice(0, 12).map(x => `<span>${esc(x)}</span>`).join("");
  el("tmContacts").innerHTML = t.keyContacts.map(k => { const i = T.indexOf(k.territory); return `<div><span class="sw" style="background:${i >= 0 ? tColor(i) : css("--orange")}"></span><span class="tt">${esc(k.territory)}</span><br>${esc(k.name)}, ${esc(k.role)}</div>`; }).join("");
  el("tmTools").innerHTML = [...t.databases.slice(0, 6), ...t.aiTools.slice(0, 6)].map(x => `<span>${esc(x)}</span>`).join("");
}

function growth() {
  if (NOUS) return notice("growth", "No US hours in the data yet", NOUS_MSG);
  const g = usPrev ? usLTM / usPrev - 1 : null;
  el("grTitle").textContent = g != null ? `US hours have grown ${pct(g)} in a year and now make up ${pct(usLTM / allLTM)} of our work` : `US work now makes up ${pct(usLTM / allLTM)} of our hours`;
  el("grLead").innerHTML = `Since US onboarding in <b>${mLabel(usFirst)}</b>, monthly US hours rose from about <b>${fH(usFirst3)}</b> to <b>${fH(usLast3)}</b> (3-month averages).`;
  const idx = M.map((_, i) => i).slice(Math.max(0, usFirst - 2));
  lazyChart(el("cUs"), {type: "bar", data: {labels: idx.map(mLabel), datasets: [{label: "US hours", data: idx.map(i => usM[i]), backgroundColor: css("--orange"), borderRadius: 3, maxBarThickness: 34}]},
    options: {maintainAspectRatio: false, plugins: {legend: {display: false}, tooltip: tip(false)}, scales: axes(false)}});
  el("grStats").innerHTML = [
    [fH(usLTM), "US hours in the last 12 months"],
    [pct(usLTM / allLTM), "of all AC hours are now US work"],
    [fH(usLast3), `US hours a month now (avg. of the last 3 months)`],
  ].map(([v, l], k) => `<div class="stat"><div class="v ${k === 0 ? "o" : ""}">${v}</div><div class="l">${l}</div></div>`).join("");
  el("grSo").innerHTML = `<b>So what:</b> US work is no longer occasional. It is steady, growing demand that justifies a dedicated, US-tagged team.`;
}

function people() {
  if (NOUS) return notice("people", "No US hours in the data yet", NOUS_MSG);
  if (NOMOVERS && !usWorkers.length) return notice("people", "No one is working on US jobs yet", NOMOVERS_MSG);
  el("ppTitle").textContent = `${movers.length} of ${usWorkers.length} ${usWorkers.length === 1 ? "person" : "people"} doing US work spend at least ${pct(THR)} of their time on it and are not yet US-tagged`;
  el("ppLead").innerHTML = `${esc(D.meta.taggingNote)} Until then, US hours are charged at each person's sales price. Based on the last ${WIN} months of hours.`;
  el("ppLegend").innerHTML = T.map((n, i) => `<span><span class="sw" style="background:${tColor(i)}"></span>${n}</span>`).join("");
  const mix = p => `<div class="mix">` + T.map((n, i) => p.h[i] > 0 ? `<div title="${n}: ${fH(p.h[i])} hrs" style="width:${p.h[i] / p.tot * 100}%;background:${tColor(i)}"></div>` : "").join("") + `</div>`;
  const list = usWorkers.slice(0, 10);
  el("ppTable").innerHTML = `<table><thead><tr><th>Person</th><th>Level</th><th>Current tag</th><th>Where their hours go</th><th class="n">US share</th><th class="n">US hours / yr</th><th>Proposal</th></tr></thead><tbody>` +
    list.map(r => `<tr class="${r.move ? "hi" : ""}"><td><b>${esc(r.s.name)}</b></td><td>${r.s.level}</td><td>${esc(r.s.tag || "–")}</td><td>${mix(r)}</td><td class="n"><b>${pct(r.share)}</b></td><td class="n">${fH(r.annH)}</td>
      <td>${r.move ? `<span class="pill go">Tag to USA</span>` : r.already ? `<span class="pill">Already US-tagged</span>` : `<span class="pill">Keep, review quarterly</span>`}</td></tr>`).join("") + `</tbody></table>`;
  el("ppSo").innerHTML = movers.length ? `<b>So what:</b> ${movers.length === 1 ? "this person already works" : `these ${movers.length} people already work`} as a US team. Tagging them to the US only makes the paperwork match reality.` : `<b>So what:</b> ${NOMOVERS_MSG}`;
}

function cost() {
  if (NOUS) return notice("cost", "No US hours in the data yet", NOUS_MSG);
  if (NOMOVERS) return notice("cost", "No one to tag at the current threshold", NOMOVERS_MSG);
  if (NOPRICE) return notice("cost", "Sales prices are missing", NOPRICE_MSG);
  el("csTitle").textContent = `Tagging them to the US saves the US about ${fK(saving)} a year, ${pct(saving / today)} of what it pays today`;
  el("csLead").innerHTML = `Today the US pays <b>${fUSD(today)}</b> a year for these ${movers.length} people's US hours (sales price). Once US-tagged, the same ${fH(moverHrs)} hours cost <b>${fUSD(tagged)}</b> at Global Model FY27 rates.`;
  lazyChart(el("cCost"), {type: "bar", data: {labels: movers.map(r => r.s.name + " (" + r.s.level + ")"), datasets: [
      {label: "Today (sales price)", data: movers.map(r => r.today), backgroundColor: css("--neutral"), borderRadius: 3, maxBarThickness: 46},
      {label: "US-tagged (Global Model rate)", data: movers.map(r => r.tagged), backgroundColor: css("--orange"), borderRadius: 3, maxBarThickness: 46}]},
    options: {maintainAspectRatio: false, plugins: {legend, tooltip: tip(true)}, scales: axes(true)}});
  el("csStats").innerHTML = [[fUSD(today), "US cost today, per year"], [fUSD(tagged), "US cost once tagged, per year"], [fUSD(saving), "saving for the US, every year"]]
    .map(([v, l], k) => `<div class="stat"><div class="v ${k === 2 ? "o" : ""}">${v}</div><div class="l">${l}</div></div>`).join("");
  el("csSo").innerHTML = `<b>So what:</b> same people, same work, same quality. The only change is the tag, and the US pays ${pct(saving / today)} less.`;
}

// ---------------- free Associate calculator
function avgRate(lv) {
  const yr = D.meta.rateYears.length ? Math.max(...D.meta.rateYears.map(Number)) : null;
  const rs = S.filter(s => s.level === lv).map(s => s.rates[yr]).filter(v => v != null);
  return rs.length ? Math.round(rs.reduce((a, b) => a + b, 0) / rs.length) : Math.round((MODEL[lv]?.rate || 50) * 1.25);
}
function freeCalc(lv, hrs, rate) {
  const rc = MODEL[lv]; const todayC = hrs * rate, seat = rc.cpp;   // one person = one dedicated seat (hours capped at a full year)
  const left = todayC - seat, assoc = left > 0 ? Math.floor(left / ASSOC.cpp) : 0, rest = left - assoc * ASSOC.cpp;
  return {todayC, seat, left, assoc, rest, be: rate ? rc.cpp / rate : 0, beAssoc: rate ? (rc.cpp + ASSOC.cpp) / rate : 0};
}
let FC = {lv: MGR, hrs: CAP, rate: 0};
function freeInit() {
  el("fcHrs").max = CAP;
  const lvs = Object.keys(MODEL).filter(l => l !== "A");
  el("fcLevel").innerHTML = lvs.map(l => `<button data-v="${l}">${MODEL[l].name}</button>`).join("");
  FC.rate = avgRate(FC.lv);
  el("fcLevel").addEventListener("click", e => { const b = e.target.closest("button"); if (!b) return; FC.lv = b.dataset.v; FC.rate = avgRate(FC.lv); freeRender(); });
  el("fcHrs").addEventListener("input", e => { FC.hrs = +e.target.value; freeRender(); });
  el("fcRate").addEventListener("input", e => { FC.rate = Math.max(0, +e.target.value || 0); freeRender(true); });
  freeRender();
}
function freeRender(keepRate) {
  el("fcLevel").querySelectorAll("button").forEach(b => b.classList.toggle("on", b.dataset.v === FC.lv));
  el("fcHrs").value = FC.hrs; el("fcHrsV").textContent = fH(FC.hrs) + " hrs";
  if (!keepRate) el("fcRate").value = FC.rate;
  const name = MODEL[FC.lv].name, r = freeCalc(FC.lv, FC.hrs, FC.rate || 0);
  const max = Math.max(r.todayC, r.seat + Math.max(r.left, 0), 1);
  const w = v => (v / max * 100).toFixed(2) + "%";
  el("fcToday").textContent = fUSD(r.todayC);
  el("fcBarToday").innerHTML = `<div class="seg-b today" style="width:${w(r.todayC)}">${name}: ${fK(r.todayC)}</div>`;
  let parts = `<div class="seg-b mgr" style="width:${w(r.seat)}">1 ${name} seat ${fK(r.seat)}</div>`;
  for (let k = 0; k < r.assoc; k++) parts += `<div class="seg-b assoc" style="width:${w(ASSOC.cpp)}">+ Associate ${fK(ASSOC.cpp)}</div>`;
  if (r.rest > 0) parts += `<div class="seg-b save" style="width:${w(r.rest)}">${r.rest / max > .08 ? "still saved " + fK(r.rest) : ""}</div>`;
  el("fcBarTagged").innerHTML = parts;
  el("fcTagged").textContent = fUSD(r.seat + r.assoc * ASSOC.cpp);
  if (r.assoc >= 1) el("fcVerdict").innerHTML = `The US gets the ${name} <b>plus ${r.assoc === 1 ? "an extra Associate" : r.assoc + " extra Associates"}</b> and still pays <b>${fK(r.rest)}</b> less than today.`;
  else if (r.left > 0) el("fcVerdict").innerHTML = `A dedicated seat saves <b>${fK(r.left)}</b> a year. From <b>${fH(r.beAssoc)} hrs</b> the saving also pays for a full Associate.`;
  else el("fcVerdict").innerHTML = `At ${fH(FC.hrs)} hrs, hourly billing at the Global Model rate is the better option. A dedicated seat pays off from <b>${fH(r.be)} hrs</b> a year.`;
  el("fcNote").innerHTML = `${name} seat = Global Model FY27 Cost/Person (${fUSD(MODEL[FC.lv].cpp)}). Associate seat = ${fUSD(ASSOC.cpp)}. One person, one seat: up to ${fH(CAP)} hrs a year.`;
  // "So what" follows the selection on screen
  el("fcSo").innerHTML = r.left > 0
    ? `<b>So what:</b> a ${name} working ${fH(FC.hrs)} US hours a year costs about <b>${fK(r.todayC)}</b> at today's price. As a dedicated US seat, the same money covers that ${name}${r.assoc ? ` <b>and ${r.assoc === 1 ? "an extra Associate" : r.assoc + " extra Associates"}</b>, with ${fK(r.rest)} to spare` : ` and saves ${fK(r.left)}`}.`
    : `<b>So what:</b> at ${fH(FC.hrs)} US hours a year, keep hourly billing. A dedicated seat pays off once the ${name} works more than ${fH(r.be)} US hours a year.`;
}

function longterm() {
  if (NOUS || NOMOVERS || NOPRICE) return notice("longterm", "The long-term view needs people to tag and their sales prices", NOUS ? NOUS_MSG : NOMOVERS ? NOMOVERS_MSG : NOPRICE_MSG);
  const steps = []; for (let g = 0; g <= 200; g += 10) steps.push(g);
  const hourly = steps.map(g => savingAt(1 + g / 100, false)), best = steps.map(g => savingAt(1 + g / 100, true));
  const cross = steps.find((g, k) => best[k] >= ASSOC.cpp), cross2 = steps.find((g, k) => best[k] >= 2 * ASSOC.cpp);
  el("ltTitle").textContent = cross != null && cross > 0
    ? `As US work grows, the yearly saving passes the cost of a full Associate at about +${cross}% US volume`
    : cross === 0 ? "The yearly saving already covers a full Associate, and it grows with US volume" : "The yearly saving grows in step with US volume";
  el("ltLead").innerHTML = `For the ${movers.length} people proposed today: at today's volume the US saves <b>${fK(best[0])}</b> a year; if US hours double it saves <b>${fK(best[steps.indexOf(100)])}</b>. Dedicated seats make the saving grow faster once people are full-time on US work.`;
  lazyChart(el("cLong"), {type: "line", data: {labels: steps.map(g => g === 0 ? "Today" : "+" + g + "%"), datasets: [
      {label: "Saving: cheapest of hourly or dedicated seat", data: best, borderColor: css("--orange"), backgroundColor: css("--orange"), borderWidth: 3, pointRadius: 0, pointHoverRadius: 5, tension: .2},
      {label: "Saving: hourly Global Model rate only", data: hourly, borderColor: css("--charcoal"), backgroundColor: css("--charcoal"), borderWidth: 2, borderDash: [5, 4], pointRadius: 0, tension: .2},
      {label: "Cost of 1 Associate seat", data: steps.map(() => ASSOC.cpp), borderColor: css("--tangerine"), backgroundColor: css("--tangerine"), borderWidth: 1.5, borderDash: [2, 3], pointRadius: 0},
      {label: "Cost of 2 Associate seats", data: steps.map(() => 2 * ASSOC.cpp), borderColor: css("--neutral"), backgroundColor: css("--neutral"), borderWidth: 1.5, borderDash: [2, 3], pointRadius: 0}]},
    options: {maintainAspectRatio: false, interaction: {mode: "index", intersect: false}, plugins: {legend, tooltip: tip(true)},
      scales: {...axes(true), x: {...axes(true).x, title: {display: true, text: "Change in US hours vs today"}}}}});
  el("ltNote").textContent = `Same ${movers.length} people, hours scaled up; seat = Global Model FY27 Cost/Person, one seat per ${fH(CAP)} hrs. Bringing in more people as they pass the threshold adds to this.`;
  el("ltSo").innerHTML = `<b>So what:</b> this is a long-term arrangement. The more the US grows with AC, the more it saves${cross2 != null ? `: from about +${cross2}% volume the saving covers two Associates` : ""}.`;
}

function winwin() {
  el("wwUS").innerHTML = ST.winUS.map(x => `<li>${esc(x)}</li>`).join("");
  el("wwAC").innerHTML = ST.winAC.map(x => `<li>${esc(x)}</li>`).join("");
}
function nextSteps() {
  el("nsSteps").innerHTML = ST.steps.map(([w, t, d]) => `<div class="step"><div class="dot"></div><div class="when">${esc(w)}</div><h3>${esc(t)}</h3><div style="color:var(--text-2)">${esc(d)}</div></div>`).join("");
  el("nsAsk").innerHTML = ST.ask.map(x => `<li>${esc(x)}</li>`).join("");
}
function appendix() {
  el("apMethod").innerHTML = [
    `<b>Data:</b> ${esc(D.meta.source)}, ${fH(R.length)} time entries, ${mLabel(0)} to ${mLabel(LAST)}. Hours by reporting month (Source Sheet).`,
    `<b>Current tagging:</b> ${esc(D.meta.taggingNote)}`,
    `<b>Who is proposed:</b> people with at least ${pct(THR)} of their hours on US work in the last ${WIN} months, not already US-tagged.`,
    `<b>Cost today:</b> US hours × the person's sales price for that year (Staff Master), annualised.`,
    `<b>Cost once tagged:</b> the same hours × Global Model FY27 Rate/hr for the person's level.`,
    `<b>Dedicated seat:</b> Global Model FY27 Cost/Person for a full year, covering up to ${fH(CAP)} hours.`,
    `<b>Extra Associate:</b> one Associate seat at ${fUSD(ASSOC.cpp)} a year.`,
    `<b>To confirm with Finance:</b> that US-tagged hours are charged at the Global Model rate, and how Cost/Person applies to a dedicated seat.`,
  ].map(x => `<li>${x}</li>`).join("");
  el("apRates").innerHTML = `<table><thead><tr><th>Level</th><th class="n">Rate / hr</th><th class="n">Cost / person / yr</th><th class="n">Seat pays off at</th></tr></thead><tbody>` +
    Object.entries(MODEL).map(([lv, r]) => `<tr><td>${esc(r.name)}</td><td class="n">${fUSD(r.rate)}</td><td class="n">${fUSD(r.cpp)}</td><td class="n">${r.rate ? fH(r.cpp / (avgRate(lv) || r.rate)) + " hrs" : "–"}</td></tr>`).join("") + `</tbody></table>`;
  // both links carry the current section, so the dashboard's "Back" button returns here
  const current = () => { const y = scrollY + innerHeight * 0.4; let id = secs[0].id; for (const s of secs) if (s.offsetTop <= y) id = s.id; return id; };
  [el("apLink"), el("dashLink")].forEach(a => { a.href = D.meta.dashboard; a.addEventListener("click", () => { a.href = D.meta.dashboard + "#" + current(); }); });
  el("ftLeft").innerHTML = esc(D.meta.classification).replace(/^((?:\w+ )?Confidential)/, "<b>$1</b>") + ` · ${esc(ST.presenter)}`;
  const c = D.meta.contact || {}, team = c.team || "";
  const ppl = (c.people || []).map(p => !p.email ? esc(p.name) : p.name ? `${esc(p.name)} (<a href="mailto:${esc(p.email)}">${esc(p.email)}</a>)` : `<a href="mailto:${esc(p.email)}">${esc(p.email)}</a>`);
  el("signText").innerHTML = ppl.length ? `For any changes, additions or further analysis, please reach out to ${ppl.join(" or ")}${team ? `, ${esc(team)}` : ""}.` : "";
  el("ftRight").textContent = `Generated ${D.meta.generated}`;
}

cover(); summary(); team(); growth(); people(); cost(); freeInit(); longterm(); winwin(); nextSteps(); appendix();

// ---------- motion: reveal each section's blocks one after another; count up the big numbers
(function reveal() {
  if (REDUCED) return;
  const blocks = [];
  document.querySelectorAll("section .wrap").forEach(w => [...w.children].forEach((c, k) => {
    c.classList.add("reveal"); c.style.transitionDelay = Math.min(k * 110, 550) + "ms"; blocks.push(c);
  }));
  const io = new IntersectionObserver(es => es.forEach(e => {
    if (!e.isIntersecting) return;
    e.target.classList.add("in"); io.unobserve(e.target);
    const delay = parseInt(e.target.style.transitionDelay) || 0;
    setTimeout(() => e.target.querySelectorAll(".stat .v").forEach(v => countUp(v, 1000)), delay + 150);
  }), {threshold: .12});
  blocks.forEach(b => io.observe(b));
})();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
