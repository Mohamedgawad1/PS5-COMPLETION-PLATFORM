"""
ps5data.py - pack/unpack the PS5 platform datasets.

The dashboard data used to live as one 6.3 MB JavaScript literal inside
index.html, which made every page load download and parse all of it before
anything appeared on screen.

The data now lives in small d_<name>.js files that index.html pulls in on
demand. To keep those files small, list-of-list tables are stored column-wise
with a per-column dictionary whenever repeating values make that smaller
(ITR drops from 4.3 MB to 1.6 MB this way). PS5UNPACK() in index.html rebuilds
the exact original arrays, so no rendering code had to change.

Format of a packed table:
    {"k": [dictionary column indexes], "D": [dictionary values], "R": [columns]}
Non-list values (objects such as SUBS, or scalars such as DPRSUM) are stored
as plain JSON.
"""

import json

# Column dictionaries only pay off when a column repeats enough.
DICT_MAX_RATIO = 0.92


def _columns(rows):
    ncol = max(len(r) for r in rows) if rows else 0
    cols = [[] for _ in range(ncol)]
    for r in rows:
        for c in range(ncol):
            cols[c].append(r[c] if c < len(r) else '')
    return cols


def pack(value):
    """Return (is_packed, payload) for one dataset."""
    if not isinstance(value, list) or not value or any(isinstance(r, dict) for r in value):
        return False, json.dumps(value, ensure_ascii=False, separators=(',', ':'))

    cols = _columns(value)
    ks, ds, rs = [], [], []
    for col in cols:
        idx, uniq = {}, []
        for v in col:
            if v not in idx:
                idx[v] = len(uniq)
                uniq.append(v)
        i_cost = len(json.dumps([idx[v] for v in col], separators=(',', ':')))
        d_cost = len(json.dumps(uniq, ensure_ascii=False, separators=(',', ':')))
        r_cost = len(json.dumps(col, ensure_ascii=False, separators=(',', ':')))
        if d_cost + i_cost < r_cost * DICT_MAX_RATIO:
            ks.append(len(rs))
            ds.append(uniq)
            rs.append([idx[v] for v in col])
        else:
            rs.append(col)
    return True, json.dumps({'k': ks, 'D': ds, 'R': rs}, ensure_ascii=False, separators=(',', ':'))


def unpack_rows(payload):
    """Python mirror of PS5UNPACK - used by the Excel exporter to read the data."""
    p = json.loads(payload)
    if not isinstance(p, dict) or 'R' not in p:
        return p
    rs, ks, ds = p['R'], p['k'], p['D']
    # ks holds the column position; D is indexed by the order dictionaries were
    # added, which is not the same thing once any column is left as plain values.
    for di, c in enumerate(ks):
        d = ds[di]
        rs[c] = [d[i] for i in rs[c]]
    n = len(rs[0]) if rs else 0
    return [[rs[c][i] for c in range(len(rs))] for i in range(n)]


def build_js(groups, banner=''):
    """groups: ordered list of (dataset_name, value). Returns the file text.

    Each dataset is assigned straight onto the window as a global, because the
    rendering code refers to bare SUBS / PUNCH / ITR / ... names. index.html
    pre-declares all of them as empty arrays so a lazy file that has not arrived
    yet can never cause a ReferenceError.
    """
    out = []
    if banner:
        out.append('/* %s */' % banner)
    out.append('(function(){"use strict";')
    for name, value in groups:
        packed, payload = pack(value)
        out.append('window.%s=%s%s%s;' % (name, 'PS5UNPACK(' if packed else '', payload,
                                         ')' if packed else ''))
    out.append('})();')
    return '\n'.join(out) + '\n'


# Which datasets go in which file, and which pages need them.
# core is small enough to always be there, so it is the only blocking one.
# The d_ prefix keeps these apart from the repo's ps5_config.js / live_itr.js.
PREFIX = 'd_'
CORE = ['SUBS', 'RFC', 'RFCK', 'DPRSUM', 'MILES', 'ITRT', 'PUNT', 'NOTES', 'PCOL', 'OVR', 'KCNT']
FILES = {
    'core': CORE,
    'punch': ['PUNCH'],
    'itr': ['ITR'],
    'cables': ['CABLES'],
    'wire': ['WIRE'],
    'cbl': ['CBL'],
    'cs': ['CS'],
    'fg': ['FG', 'FZ'],
    'io': ['IO'],
    'trm': ['TRM'],
}


def roundtrip(value):
    """Parse a d_*.js payload back into the original value (test helper)."""
    p = json.loads(value)
    if isinstance(p, dict) and 'R' in p:
        return unpack_rows(value)
    return p
