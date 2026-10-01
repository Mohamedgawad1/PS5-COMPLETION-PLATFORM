# -*- coding: utf-8 -*-
"""
dpr_file.py - ONE resolver for "the current PS5 DPR workbook".

Every script in the pipeline (rebuild_data.py, sync_cloud_to_excel.py,
punch_itr_explorer.py, dpr_dashboard.py, apply_platform_edits_to_summery.py)
used to search for the DPR file on its own. They disagreed on two things:

  1. SPELLING - the daily file is "DPR SUMMARY" (correct) while the older
     files are "DPR SUMMERY". A plain 'completions dpr summery' substring test
     only matched the stale files, so the pipeline silently built from an
     August workbook instead of the current one.

  2. LOCATION - two files with the IDENTICAL name existed:
        Downloads\\PS-5 COMPLETIONS DPR SUMMARY -30-SEPTEMBER-2026.xlsx
        Downloads\\PS5 - CPP ... Dashboard_files\\<same name>
     sync_cloud_to_excel.py wrote platform edits into the first one while
     rebuild_data.py read the second one, so every edit was written into a file
     nobody rebuilt from. That is why the platform never updated.

Both finders now delegate here, so there is exactly one answer.
"""
import glob
import os
import re

HOME = os.path.join(os.path.expanduser('~'), 'Downloads')

# The one place the daily DPR workbook lives. Every read AND every write must
# use this exact file.
CANON_DIR = os.path.join(
    HOME, 'PS5 - CPP AGI Completion Progress Dashboard_files')

# Only used to discover a file the first time, or if the canonical copy is
# missing. Never written to.
FALLBACK_DIRS = (HOME,)

STEM = r'PS-?5\s*COMPLETIONS\s*DPR\s*SUMM(?:ARY|ERY)'
DPR_RE = re.compile(STEM, re.I)

DATE_NUM = re.compile(STEM + r'\s*-\s*(\d{1,2})-(\d{1,2})-(\d{2,4})', re.I)
DATE_MON = re.compile(STEM + r'\s*-\s*(\d{1,2})-([A-Za-z]+)-(\d{2,4})', re.I)

MONTHS = {'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05',
          'JUN': '06', 'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10',
          'NOV': '11', 'DEC': '12'}


def is_dpr(name):
    """True for the DPR workbook under either spelling, backups excluded."""
    if name.startswith('~$'):
        return False
    if 'BACKUP' in name.upper():
        return False
    return bool(DPR_RE.search(name))


def file_date(name):
    """(year, month, day) from a '-DD-MM-YY' / '-DD-MONTH-YYYY' suffix,
    or None when the name carries no date."""
    m = DATE_NUM.search(name)
    if m:
        dd, mm, yy = m.group(1), m.group(2), m.group(3)
        return (yy if len(yy) == 4 else '20' + yy, mm.zfill(2), dd.zfill(2))
    m = DATE_MON.search(name)
    if m and m.group(2)[:3].upper() in MONTHS:
        dd, yy = m.group(1), m.group(3)
        return (yy if len(yy) == 4 else '20' + yy,
                MONTHS[m.group(2)[:3].upper()], dd.zfill(2))
    return None


def _scan(dirs):
    out = []
    for folder in dirs:
        if not os.path.isdir(folder):
            continue
        for path in glob.glob(os.path.join(folder, '*.xlsx')):
            name = os.path.basename(path)
            if is_dpr(name):
                out.append((file_date(name) or ('', '', ''),
                            os.path.getmtime(path), path))
    return out


def find_dpr():
    """The single canonical DPR workbook, or None.

    Prefers CANON_DIR and, when several files sit there, the newest one by the
    date in its filename (not by mtime - a file can be touched without being
    newer).
    """
    in_canon = _scan((CANON_DIR,))
    if in_canon:
        return max(in_canon)[2]
    found = _scan(FALLBACK_DIRS)
    return max(found)[2] if found else None


def warn_if_shadow_copies(path):
    """If a same-named DPR also exists outside CANON_DIR, say so.

    Shadow copies are how the writes and the reads drifted apart, so this is
    called before every write.
    """
    if not path:
        return []
    real = os.path.realpath(path)
    shadow = []
    for _, _, other in _scan(FALLBACK_DIRS):
        if os.path.basename(other) == os.path.basename(path) \
                and os.path.realpath(other) != real:
            shadow.append(other)
    if shadow:
        print('[dpr] WARNING: a second copy with the same name exists:')
        for s in shadow:
            print('[dpr]   %s' % s)
        print('[dpr]   the pipeline only reads/writes %s' % path)
    return shadow