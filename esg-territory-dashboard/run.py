"""
ESG Dashboard v2 - one step build
=================================
Double-click  'Build Dashboards.bat'   (or run:  python run.py)

It will:
  1. add any missing input tabs to the workbook (only the first time; a backup is kept)
  2. build ESG_Hours_Dashboard.html   (full dashboard)
  3. build ESG_US_Story_v2.html       (partner story)
  4. open the story page

Close the workbook in Excel before running.
"""
# ============================== SETTINGS ==============================
# Defaults for the sample. Put YOUR values in local_settings.py (same names) -
# that file is ignored by Git, so your real file paths, names and email are never committed.
EXCEL_FILE = "ESG_Hours_Sample.xlsx"   # file name in this folder, or a full path like r"D:\Data\Hours.xlsx"

CONTACT_NAME = ""                      # shown at the end of both pages
CONTACT_EMAIL = ""
CONTACT_NAME_2 = ""                    # optional second contact
CONTACT_EMAIL_2 = ""

LOGO_FILE = "logo.png"                 # file in engine/ (ignored by Git)
LOGO_TEXT = ""                         # text shown instead when there is no logo file
CLASSIFICATION = "Confidential | Internal use only"   # used when the workbook has no 'Classification' assumption

OPEN_WHEN_DONE = True                  # open the story page in the browser at the end
# ======================================================================
try:
    from local_settings import *       # noqa: F401,F403  (your private overrides)
except ImportError:
    pass

import sys
import traceback
import webbrowser
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "engine"))


def run(module, *args):
    """Call a module's main() with command-line style arguments."""
    old = sys.argv
    sys.argv = [module.__name__, *map(str, args)]
    try:
        module.main()
    finally:
        sys.argv = old


def main():
    try:
        import pandas, openpyxl  # noqa: F401
    except ImportError:
        sys.exit("Python packages missing. Install them once with:\n    py -m pip install -r requirements.txt\nthen run again.")
    import add_input_tabs, build_dashboard as bd, build_story_v2 as sv

    excel = Path(EXCEL_FILE)
    excel = excel if excel.is_absolute() else HERE / excel
    if not excel.exists():
        sys.exit(f"Excel file not found: {excel}\nSet EXCEL_FILE at the top of run.py.")

    bd.CONTACT_NAME, bd.CONTACT_EMAIL = CONTACT_NAME, CONTACT_EMAIL
    bd.CONTACT_NAME_2, bd.CONTACT_EMAIL_2 = CONTACT_NAME_2, CONTACT_EMAIL_2
    bd.LOGO_FILE, bd.LOGO_TEXT, bd.CLASSIFICATION = LOGO_FILE, LOGO_TEXT, CLASSIFICATION
    dash, story = HERE / "ESG_Hours_Dashboard.html", HERE / "ESG_US_Story_v2.html"

    print(f"Workbook: {excel.name}\n")
    print("[1/3] Checking input tabs ...")
    run(add_input_tabs, "--excel", excel)
    print("\n[2/3] Building the dashboard ...")
    run(bd, "--excel", excel, "--out", dash)
    print("\n[3/3] Building the partner story ...")
    run(sv, "--excel", excel, "--out", story)
    print(f"\nDone. Open: {story.name}")
    if OPEN_WHEN_DONE:
        webbrowser.open(story.as_uri())


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print(f"\nA file is open in another program ({e.filename}). Close it (e.g. the workbook in Excel) and run again.")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        print("\nSomething went wrong - see the message above.")
        sys.exit(1)
