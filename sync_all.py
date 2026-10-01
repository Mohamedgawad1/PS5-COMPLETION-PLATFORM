"""sync_all.py - one-touch full sync for the PS5 platform.

Pipeline:
  1) rebuild_data.py            -> index.html from the newest DPR SUMMARY + ITR/Punch
  2) make_platform_excel.py     -> golden PS5 PLATFORM.xlsx from the platform data
  3) make_platform_excel.py --pages --local -> PAGES/*.xlsx (all 9 download pages)
  4) copy PAGES/* -> EXCEL/*  (this is what the download buttons serve)
  5) refresh EXCEL/7 - COMPLETE.xlsm (34 rows) while preserving the (+ ROW) macro
  6) git add/commit/push the changed files (EXCEL/*, PS5 PLATFORM.xlsx,
     index.html, rebuild_data.py, make_platform_excel.py)

Run from the repo folder:  python sync_all.py
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(HERE, "EXCEL")
PAGES = os.path.join(HERE, "PAGES")
GOLDEN = os.path.join(HERE, "PS5 PLATFORM.xlsx")


def run(cmd):
    print("\n>>>", " ".join(cmd))
    r = subprocess.run(cmd, cwd=HERE)
    if r.returncode != 0:
        print("!! FAILED:", cmd, "exit", r.returncode)
        sys.exit(r.returncode)


def refresh_xlsm():
    import openpyxl
    new = os.path.join(PAGES, "7 - COMPLETE.xlsx")
    old = os.path.join(EXCEL, "7 - COMPLETE.xlsm")
    if not (os.path.exists(new) and os.path.exists(old)):
        print("  (xlsm refresh skipped - files missing)")
        return
    wb_new = openpyxl.load_workbook(new, data_only=True)
    rows = [list(r) for r in wb_new["7 \u00b7 COMPLETE"].iter_rows(min_row=2, values_only=True)]
    wb_new.close()
    wb = openpyxl.load_workbook(old, keep_vba=True, data_only=True)
    ws = wb["7 \u00b7 COMPLETE"]
    for i in range(ws.max_row, 1, -1):
        ws.delete_rows(i)
    for ri, r in enumerate(rows, start=2):
        for ci, v in enumerate(r, start=1):
            ws.cell(row=ri, column=ci, value=v)
    wb.save(old)
    wb.close()
    print(f"  [xlsm] 7 - COMPLETE.xlsm refreshed -> {len(rows)} rows, macro kept")


def copy_pages():
    os.makedirs(EXCEL, exist_ok=True)
    for f in sorted(os.listdir(PAGES)):
        if f.lower().endswith(".xlsx"):
            shutil.copy2(os.path.join(PAGES, f), os.path.join(EXCEL, f))
    print("  [files] PAGES/* -> EXCEL/*")


def main():
    run([sys.executable, "rebuild_data.py"])
    run([sys.executable, "make_platform_excel.py"])
    run([sys.executable, "make_platform_excel.py", "--pages", "--local"])
    copy_pages()
    refresh_xlsm()

    print("\n>>> git commit / pull --rebase / push")
    publish()
    print("\nDONE - platform online + all downloads synced.")


def publish():
    """Commit the build, rebase it onto origin/main, then push.

    Order matters. `git pull --rebase` refuses to run on a dirty tree
    ("cannot pull with rebase: You have unstaged changes", exit 128) and the
    old order pulled BEFORE committing, so every cycle aborted here and the
    live platform never received a single update. The pull is wrapped in a
    stash + pop so a leftover untracked/modified file can never block the
    push either.
    """
    tracked = ["EXCEL", GOLDEN, "index.html", "rebuild_data.py",
               "make_platform_excel.py", "sync_all.py",
               "sync_cloud_to_excel.py"]
    subprocess.run(["git", "add", "--"] + tracked, cwd=HERE, check=True)

    staged = subprocess.run(["git", "diff", "--cached", "--quiet"],
                            cwd=HERE).returncode
    committed = False
    if staged != 0:
        subprocess.run(["git", "commit", "-m",
                        "Full sync: platform + all Excel downloads"],
                       cwd=HERE, check=True)
        committed = True
        print("  committed the new build")

    # Rebase onto origin/main with the tree clean. `git stash -u` also parks
    # untracked leftovers, which is what used to make the pull impossible.
    stash = subprocess.run(["git", "stash", "push", "-u", "-m",
                            "auto-sync: untracked leftovers"],
                           cwd=HERE, capture_output=True, text=True)
    parked = 'No local changes to save' not in (stash.stdout or '')
    try:
        r = subprocess.run(["git", "pull", "--rebase", "origin", "main"],
                           cwd=HERE, capture_output=True, text=True)
        if r.returncode != 0:
            # A half-finished rebase would block every future run.
            subprocess.run(["git", "rebase", "--abort"], cwd=HERE,
                           capture_output=True)
            print("  pull --rebase failed (%s) - pushing our commit anyway"
                  % r.returncode)
            print((r.stderr or '').strip()[-400:])
        else:
            print("  rebased on origin/main")
    finally:
        if parked:
            subprocess.run(["git", "stash", "pop"], cwd=HERE,
                           capture_output=True)

    if not committed:
        print("  nothing staged to commit - already in sync")

    for attempt in (1, 2):
        r = subprocess.run(["git", "push", "origin", "main"], cwd=HERE,
                           capture_output=True, text=True)
        if r.returncode == 0:
            print("  pushed to origin/main")
            return
        # Someone pushed between our pull and our push - rebase and retry once.
        print("  push rejected (%s) - rebasing and retrying" % r.returncode)
        subprocess.run(["git", "fetch", "origin", "main"], cwd=HERE,
                       capture_output=True)
        subprocess.run(["git", "rebase", "origin/main"], cwd=HERE,
                       capture_output=True)
    print((r.stderr or '').strip()[-400:])
    raise SystemExit("push failed - the live platform was NOT updated")


if __name__ == "__main__":
    main()