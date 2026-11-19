#!/usr/bin/env python3
"""model_map.py - group the variables of an equation model and draw a map you can unfold.

Give it the model's equations (EViews style: Y = formula). It finds groups of variables that are
densely linked inside and loosely linked to the rest (Louvain community detection on the equation
links), splits each group again until every view fits on one screen, finds each group's key variables
(the fewest variables every link between it and the rest passes through), and writes one HTML page:

  * a map of all variables that starts folded; click a group to unfold it in place, down to single
    variables, while other groups stay folded
  * the main chain through the model, zoomable step by step
  * a table of every variable, folded by the same groups
  * a search that unfolds the map down to a variable, and a panel with its equation and links

Scenario data is optional: with --a and --b the boxes are coloured by the gap B - A.
Standard-library Python only (no pip install).

  python model_map.py --eq model.txt -o map.html
  python model_map.py --eq model.txt --a base.csv --b scenario.xlsx -o map.html
  python model_map.py --eq model.txt --groups            # print the group tree in the terminal
"""
import argparse
import csv
import datetime
import json
import math
import os
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from array import array
from collections import Counter, defaultdict, deque

NAN = float('nan')
NA_STRINGS = {'NA', 'N/A', '#N/A', 'NAN', 'ND', 'NULL', 'NONE', '-', '.', '..',
              '#VALUE!', '#DIV/0!', '#REF!', '#NUM!'}

try:
    csv.field_size_limit(min(sys.maxsize, 2 ** 31 - 1))
except Exception:  # pragma: no cover
    pass


def _num(cell):
    """Parse a numeric cell; missing / text -> NaN."""
    if cell is None:
        return NAN
    if isinstance(cell, (int, float)):
        v = float(cell)
        return v if math.isfinite(v) else NAN
    t = str(cell).strip()
    if not t or t.upper() in NA_STRINGS:
        return NAN
    t = t.replace(',', '')
    try:
        v = float(t)
    except ValueError:
        return NAN
    return v if math.isfinite(v) else NAN


def _clean(c):
    return '' if c is None else str(c).strip()


def _isnan(v):
    return v != v


# --------------------------------------------------------------------------
# Periods
# --------------------------------------------------------------------------
MONTHS = {m: i + 1 for i, m in enumerate(
    'JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split())}
_RE_Q1 = re.compile(r'^(\d{4})\s*[-\s]?\s*[Qq]\s*([1-4])$')
_RE_Q2 = re.compile(r'^[Qq]([1-4])[\s\-/:]*(\d{4})$')
_RE_M1 = re.compile(r'^(\d{4})\s*[Mm]\s*(\d{1,2})$')
_RE_YM = re.compile(r'^(\d{4})-(\d{1,2})$')
_RE_COLON = re.compile(r'^(\d{4}):(\d{1,2})$')
_RE_ISO = re.compile(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T].*)?$')
_RE_US = re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:\s.*)?$')
_RE_Y = re.compile(r'^(\d{4})\s*[Aa]?$')
_RE_MON = re.compile(r'^([A-Za-z]{3})[a-z]*[\s\-]+(\d{2}|\d{4})$')


def _date_key(y, mo, d):
    try:
        datetime.date(y, mo, d)
    except ValueError:
        return None
    return ((y, 5, mo * 100 + d), '%04d-%02d-%02d' % (y, mo, d))


def parse_period(cell, allow_serial=False):
    """Return (sortkey, label) if the cell looks like a date/period, else None."""
    if cell is None:
        return None
    t = str(cell).strip()
    if not t or len(t) > 30:
        return None
    m = _RE_Q1.match(t)
    if m:
        y, q = int(m.group(1)), int(m.group(2))
        return ((y, 2, q), '%dQ%d' % (y, q))
    m = _RE_Q2.match(t)
    if m:
        q, y = int(m.group(1)), int(m.group(2))
        return ((y, 2, q), '%dQ%d' % (y, q))
    m = _RE_M1.match(t) or _RE_YM.match(t)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return ((y, 3, mo), '%dM%02d' % (y, mo))
        return None
    m = _RE_COLON.match(t)
    if m:
        y, s = int(m.group(1)), int(m.group(2))
        if 1 <= s <= 12:
            return ((y, 4, s), '%d:%d' % (y, s))
        return None
    m = _RE_ISO.match(t)
    if m:
        return _date_key(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = _RE_US.match(t)
    if m:
        mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000 if y < 50 else 1900
        return _date_key(y, mo, d)
    m = _RE_MON.match(t)
    if m and m.group(1).upper() in MONTHS:
        y = int(m.group(2))
        if y < 100:
            y += 2000 if y < 50 else 1900
        mo = MONTHS[m.group(1).upper()]
        return ((y, 3, mo), '%dM%02d' % (y, mo))
    m = _RE_Y.match(t)
    if m:
        y = int(m.group(1))
        if 1800 <= y <= 2200:
            return ((y, 1, 0), str(y))
        return None
    if allow_serial:  # Excel date serial number (xlsx headers)
        try:
            v = float(t)
        except ValueError:
            return None
        if 15000 < v < 80000 and abs(v - round(v)) < 1e-9:
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(round(v)))
            return _date_key(d.year, d.month, d.day)
    return None


def infer_ppy(keys):
    """Periods per year, used by @PCHY/@PCY."""
    if not keys:
        return 4
    kind = Counter(k[1] for k in keys).most_common(1)[0][0]
    if kind == 2:
        return 4
    if kind == 3:
        return 12
    if kind == 1:
        return 1
    if kind == 4:
        return 12 if max(k[2] for k in keys if k[1] == 4) > 4 else 4
    ds = sorted(datetime.date(k[0], k[2] // 100, k[2] % 100) for k in keys if k[1] == 5)
    if len(ds) < 2:
        return 4
    gaps = sorted((b - a).days for a, b in zip(ds, ds[1:]))
    g = gaps[len(gaps) // 2]
    if g < 45:
        return 12
    if g < 200:
        return 4
    return 1


# --------------------------------------------------------------------------
# Reading files (CSV / TSV / XLSX) with the standard library
# --------------------------------------------------------------------------
_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n - 1


def read_xlsx(path, sheet=None):
    """Minimal .xlsx reader (zipfile + xml). Returns list of rows of strings."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if 'xl/sharedStrings.xml' in names:
            for _, el in ET.iterparse(z.open('xl/sharedStrings.xml')):
                if el.tag == _NS + 'si':
                    shared.append(''.join(t.text or '' for t in el.iter(_NS + 't')))
                    el.clear()
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'): r.get('Target') for r in rels}
        sheets = []
        for s in wb.iter(_NS + 'sheet'):
            tgt = (relmap.get(s.get(_RNS + 'id')) or '').lstrip('/')
            if not tgt.startswith('xl/'):
                tgt = 'xl/' + tgt
            sheets.append((s.get('name'), tgt))
        if not sheets:
            raise ValueError('%s: no worksheets found' % path)
        target = sheets[0][1]
        if sheet is not None:
            found = None
            for i, (nm, tg) in enumerate(sheets):
                if nm.lower() == str(sheet).lower() or str(i + 1) == str(sheet):
                    found = tg
            if not found:
                raise ValueError('%s: sheet %r not found; sheets are %s'
                                 % (path, sheet, [s[0] for s in sheets]))
            target = found
        grid = []
        for _, el in ET.iterparse(z.open(target)):
            if el.tag != _NS + 'row':
                continue
            row = []
            for c in el.findall(_NS + 'c'):
                ref = c.get('r')
                j = _col_index(ref) if ref else len(row)
                t = c.get('t')
                val = ''
                if t == 's':
                    v = c.find(_NS + 'v')
                    if v is not None and v.text is not None:
                        val = shared[int(v.text)]
                elif t == 'inlineStr':
                    is_ = c.find(_NS + 'is')
                    if is_ is not None:
                        val = ''.join(x.text or '' for x in is_.iter(_NS + 't'))
                else:
                    v = c.find(_NS + 'v')
                    if v is not None and v.text is not None:
                        val = v.text
                        if t not in ('str', 'e', 'b'):
                            try:
                                f = float(val)
                                val = str(int(f)) if f.is_integer() and abs(f) < 1e15 else repr(f)
                            except ValueError:
                                pass
                if j >= len(row):
                    row.extend([''] * (j - len(row) + 1))
                row[j] = val
            r = el.get('r')
            ri = int(r) - 1 if r else len(grid)
            while len(grid) < ri:
                grid.append([])
            grid.append(row)
            el.clear()
    return grid


def xlsx_sheets(path):
    """Names of the worksheets of an .xlsx file, in order."""
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        return [s.get('name') for s in wb.iter(_NS + 'sheet')]


def read_grid(path, sheet=None):
    """Return (rows, is_xlsx)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        return read_xlsx(path, sheet), True
    if ext == '.xls':
        raise SystemExit('%s: old .xls format is not supported - save it as .xlsx or .csv' % path)
    with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
        sample = f.read(65536)
        f.seek(0)
        delim = ','
        if ext in ('.tsv', '.tab'):
            delim = '\t'
        else:
            try:
                delim = csv.Sniffer().sniff(sample, delimiters=',\t;|').delimiter
            except csv.Error:
                pass
        return [row for row in csv.reader(f, delimiter=delim)], False


# --------------------------------------------------------------------------
# Scenario data
# --------------------------------------------------------------------------
NAME_HDRS = {'mnemonic', 'mnemonics', 'series', 'variable', 'name', 'code', 'ticker',
             'id', 'series id', 'series_id', 'mnem', 'series name', 'var'}


def _is_desc_hdr(h):
    h = h.lower().rstrip(':').strip()
    return ('desc' in h or 'title' in h or h in ('concept', 'label', 'long name', 'longname'))


class Scenario:
    def __init__(self, label):
        self.label = label
        self.blocks = []        # list of lists of period sortkeys
        self.series = {}        # NAME -> (block id, array of floats)
        self.desc = {}
        self.keys = {}          # sortkey -> label
        self.notes = []
        self.dups = 0
        self.keep_empty = False

    def put(self, name, bid, arr, desc=''):
        name = name.strip().upper()
        if not name:
            return
        if name in self.series:
            self.dups += 1
        self.series[name] = (bid, arr)
        if desc and name not in self.desc:
            self.desc[name] = desc


def _pick_name_col(grid, start, candidates, hdr):
    for j in candidates:
        if j < len(hdr) and hdr[j] in NAME_HDRS:
            return j
    best, bestscore = None, -1
    sample = grid[start:start + 60]
    for j in candidates:
        score = 0
        for r in sample:
            c = _clean(r[j]) if j < len(r) else ''
            if c and _isnan(_num(c)):
                score += 1
        if score > bestscore:
            best, bestscore = j, score
    return best


def _increasing(cands):
    """Keep period cells that form an increasing run (rejects numbers that merely look like dates)."""
    out, last = [], None
    for j, p in cands:
        if not p:
            continue
        if last is None or p[0] > last:
            out.append((j, p))
            last = p[0]
    # a real date axis has (almost) no out-of-order cells
    n = sum(1 for _, p in cands if p)
    return out if n and len(out) >= 0.9 * n else []


def _ingest(sc, grid, serial, path):
    nonempty = [i for i, r in enumerate(grid) if any(_clean(c) for c in r)]
    if not nonempty:
        raise SystemExit('%s: file is empty' % path)
    h = nonempty[0]
    hdr = [_clean(c).lower() for c in grid[h]]

    def find(cands):
        for j, c in enumerate(hdr):
            if c in cands:
                return j
        return -1
    vcol = find({'value', 'val', 'data', 'obs_value', 'amount'})
    dcol = find({'date', 'period', 'time', 'obs', 'datetime', 'quarter', 'month', 'year'})
    scol = find(NAME_HDRS)
    if vcol >= 0 and dcol >= 0 and scol >= 0 and len([c for c in hdr if c]) <= 10:
        return _ingest_long(sc, grid, h, scol, dcol, vcol, serial, path)

    rows_hdr = None
    for i in nonempty[:25]:
        cols = _increasing([(j, parse_period(c, serial)) for j, c in enumerate(grid[i])])
        if len(cols) >= 3:
            rows_hdr = (i, cols)
            break
    cols_best = None
    for c in range(4):
        idx = _increasing([(i, parse_period(r[c], serial) if c < len(r) else None)
                           for i, r in enumerate(grid)])
        if len(idx) >= 3 and (cols_best is None or len(idx) > len(cols_best[1])):
            cols_best = (c, idx)
    nr = len(rows_hdr[1]) if rows_hdr else 0
    nc = len(cols_best[1]) if cols_best else 0
    if nr == 0 and nc == 0:
        raise SystemExit('%s: could not find a row or column of dates/periods '
                         '(e.g. 2026Q1, 2026-03-31, 2026:1).' % path)
    if nr >= nc:
        _ingest_rows(sc, grid, rows_hdr[0], rows_hdr[1], path)
    else:
        _ingest_cols(sc, grid, cols_best[0], cols_best[1], path)


def _ingest_rows(sc, grid, i, cols, path):
    for _, p in cols:
        sc.keys[p[0]] = p[1]
    datecols = set(j for j, _ in cols)
    hdr = [_clean(c).lower() for c in grid[i]]
    width = max(len(r) for r in grid[i:i + 50])
    nondate = [j for j in range(width) if j not in datecols]
    if not nondate:
        raise SystemExit('%s: no column with series names found' % path)
    namecol = _pick_name_col(grid, i + 1, nondate, hdr)
    desccol = next((j for j in nondate if j < len(hdr) and _is_desc_hdr(hdr[j])), None)
    bid = len(sc.blocks)
    sc.blocks.append([p[0] for _, p in cols])
    n = 0
    jlist = [j for j, _ in cols]
    for r in grid[i + 1:]:
        name = _clean(r[namecol]) if namecol < len(r) else ''
        if not name:
            continue
        L = len(r)
        arr = array('d', (_num(r[j]) if j < L else NAN for j in jlist))
        if all(_isnan(v) for v in arr) and not sc.keep_empty:
            continue
        desc = _clean(r[desccol]) if desccol is not None and desccol < L else ''
        sc.put(name, bid, arr, desc)
        n += 1
    sc.notes.append('%s: series in rows, %d series x %d periods (%s .. %s), names from column %d'
                    % (os.path.basename(path), n, len(cols), cols[0][1][1], cols[-1][1][1], namecol + 1))


def _ingest_cols(sc, grid, c, idx, path):
    for _, p in idx:
        sc.keys[p[0]] = p[1]
    first = idx[0][0]
    pre = [i for i in range(first) if any(_clean(x) for x in grid[i])]
    if not pre:
        raise SystemExit('%s: dates are in column %d but there is no header row with series names'
                         % (path, c + 1))
    namerow, descrow = None, None
    for i in pre:
        lab = _clean(grid[i][c] if c < len(grid[i]) else '').lower().rstrip(':').strip()
        if not lab and c > 0:
            lab = _clean(grid[i][0]).lower().rstrip(':').strip()
        if lab in NAME_HDRS and namerow is None:
            namerow = i
        elif _is_desc_hdr(lab) and descrow is None:
            descrow = i
    if namerow is None:
        best = -1
        for i in pre:
            sc_ = sum(1 for j, x in enumerate(grid[i]) if j > c and _clean(x) and _isnan(_num(x)))
            if sc_ > best:
                best, namerow = sc_, i
    names = grid[namerow]
    descs = grid[descrow] if descrow is not None else []
    bid = len(sc.blocks)
    sc.blocks.append([p[0] for _, p in idx])
    rows = [grid[i] for i, _ in idx]
    n = 0
    for j in range(c + 1, len(names)):
        name = _clean(names[j])
        if not name:
            continue
        arr = array('d', (_num(r[j]) if j < len(r) else NAN for r in rows))
        if all(_isnan(v) for v in arr) and not sc.keep_empty:
            continue
        sc.put(name, bid, arr, _clean(descs[j]) if j < len(descs) else '')
        n += 1
    sc.notes.append('%s: series in columns, %d series x %d periods (%s .. %s), names from row %d'
                    % (os.path.basename(path), n, len(idx), idx[0][1][1], idx[-1][1][1], namerow + 1))


def _ingest_long(sc, grid, h, scol, dcol, vcol, serial, path):
    data = defaultdict(dict)
    for r in grid[h + 1:]:
        if max(scol, dcol, vcol) >= len(r):
            continue
        p = parse_period(r[dcol], serial)
        name = _clean(r[scol])
        if not p or not name:
            continue
        sc.keys[p[0]] = p[1]
        data[name][p[0]] = _num(r[vcol])
    keys = sorted(set(k for d in data.values() for k in d))
    bid = len(sc.blocks)
    sc.blocks.append(keys)
    for name, d in data.items():
        sc.put(name, bid, array('d', (d.get(k, NAN) for k in keys)))
    sc.notes.append('%s: long format (series, date, value), %d series' % (os.path.basename(path), len(data)))


def load_scenario(paths, label, sheet=None, keep_empty=True):
    """Read scenario data. For an .xlsx with no sheet named, every sheet with a row or column of dates is
    read; sheets without dates (notes, metadata) are skipped and listed in sc.notes."""
    sc = Scenario(label)
    sc.keep_empty = keep_empty
    for p in paths:
        if sheet is None and os.path.splitext(p)[1].lower() in ('.xlsx', '.xlsm'):
            names = xlsx_sheets(p)
            read = 0
            for k, nm in enumerate(names):
                grid, serial = read_xlsx(p, k + 1), True
                if not any(any(_clean(c) for c in r) for r in grid):
                    continue
                try:
                    _ingest(sc, grid, serial, '%s [%s]' % (p, nm))
                    read += 1
                except (SystemExit, ValueError) as e:
                    sc.notes.append('%s: sheet %r skipped (%s)' % (os.path.basename(p), nm, str(e).split(': ', 1)[-1][:80]))
            if not read:
                raise SystemExit('%s: no sheet with a row or column of dates' % p)
        else:
            grid, serial = read_grid(p, sheet)
            _ingest(sc, grid, serial, p)
    if sc.dups:
        sc.notes.append('%d duplicate series names (last one kept)' % sc.dups)
    return sc


def load_meta(path):
    grid, _ = read_grid(path)
    hdr = [_clean(c).lower() for c in grid[0]]
    nc = next((j for j, h in enumerate(hdr) if h in NAME_HDRS), 0)
    dc = next((j for j, h in enumerate(hdr) if _is_desc_hdr(h)), 1)
    out = {}
    for r in grid[1:]:
        if nc < len(r) and dc < len(r) and _clean(r[nc]):
            out[_clean(r[nc]).upper()] = _clean(r[dc])
    return out


# --------------------------------------------------------------------------
# EViews-style equations: tokenizer, parser, evaluator
# --------------------------------------------------------------------------
class ParseError(Exception):
    pass


class EvalError(Exception):
    pass


_TOKRE = re.compile(r'''\s*(?:
    (?P<str>"[^"]*")
  | (?P<num>(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?)
  | (?P<name>@?[A-Za-z_][A-Za-z0-9_$.%#]*)
  | (?P<op>\*\*|<=|>=|<>|[-+*/^(),=<>])
)''', re.X)

FUNCS = {'LOG', 'LN', 'EXP', 'ABS', 'SQR', 'SQRT', 'D', 'DLOG', 'LOG10', 'INV', 'LAG'}
EXCLUDE = {'C', 'AR', 'MA', 'SAR', 'SMA', 'PDL', 'NA'}


def tokenize(s):
    toks, pos, n = [], 0, len(s)
    while pos < n:
        if s[pos].isspace():
            pos += 1
            continue
        m = _TOKRE.match(s, pos)
        if not m or m.end() == pos:
            raise ParseError('unexpected character %r at "%s"' % (s[pos], s[max(0, pos - 15):pos + 10]))
        if m.group('str') is not None:
            toks.append(('str', m.group('str')[1:-1]))
        elif m.group('num') is not None:
            toks.append(('num', m.group('num')))
        elif m.group('name') is not None:
            toks.append(('name', m.group('name')))
        else:
            op = m.group('op')
            toks.append(('op', '^' if op == '**' else op))
        pos = m.end()
    return toks


class _Parser:
    def __init__(self, toks):
        self.t = toks
        self.i = 0

    def peek(self, k=0):
        j = self.i + k
        return self.t[j] if j < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, v):
        if self.take()[1] != v:
            raise ParseError('expected %r' % v)

    def parse(self):
        n = self.full()
        if self.i != len(self.t):
            raise ParseError('unexpected %r' % (self.peek()[1],))
        return n

    def full(self):   # a or b
        n = self.andx()
        while self.peek()[0] == 'name' and self.peek()[1].upper() == 'OR':
            self.take()
            n = ('bin', 'OR', n, self.andx())
        return n

    def andx(self):   # a and b
        n = self.cmp()
        while self.peek()[0] == 'name' and self.peek()[1].upper() == 'AND':
            self.take()
            n = ('bin', 'AND', n, self.cmp())
        return n

    def cmp(self):    # a < b, a = b, ...
        n = self.expr()
        while self.peek()[1] in ('<', '>', '<=', '>=', '=', '<>') and self.peek()[0] == 'op':
            op = self.take()[1]
            n = ('bin', op, n, self.expr())
        return n

    def expr(self):
        n = self.term()
        while self.peek()[1] in ('+', '-'):
            op = self.take()[1]
            n = ('bin', op, n, self.term())
        return n

    def term(self):
        n = self.unary()
        while self.peek()[1] in ('*', '/'):
            op = self.take()[1]
            n = ('bin', op, n, self.unary())
        return n

    def unary(self):
        if self.peek()[1] in ('-', '+'):
            op = self.take()[1]
            u = self.unary()
            return ('neg', u) if op == '-' else u
        return self.power()

    def power(self):
        b = self.atom()
        if self.peek()[1] == '^':
            self.take()
            return ('bin', '^', b, self.unary())
        return b

    def _try_lag(self):
        # NAME ( [+-] int )
        t0, t1, t2, t3 = self.peek(0), self.peek(1), self.peek(2), self.peek(3)
        if t0[1] != '(':
            return None
        if t1[0] == 'num' and t2[1] == ')' and re.fullmatch(r'\d+', t1[1]):
            self.i += 3
            return int(t1[1])
        if t1[1] in ('-', '+') and t2[0] == 'num' and t3[1] == ')' and re.fullmatch(r'\d+', t2[1]):
            self.i += 4
            return -int(t2[1]) if t1[1] == '-' else int(t2[1])
        return None

    def atom(self):
        kind, val = self.take()
        if kind == 'num':
            return ('num', float(val))
        if kind == 'str':
            return ('str', val)
        if kind == 'op' and val == '(':
            n = self.full()
            self.expect(')')
            return n
        if kind == 'name':
            up = val.upper()
            is_func = up in FUNCS or up.startswith('@')
            if self.peek()[1] == '(':
                if not is_func:
                    lag = self._try_lag()
                    if lag is not None:
                        return ('coef', lag) if up == 'C' else ('var', up, lag)
                self.take()
                args = []
                if self.peek()[1] != ')':
                    args.append(self.full())
                    while self.peek()[1] == ',':
                        self.take()
                        args.append(self.full())
                self.expect(')')
                return ('call', up, args)
            if up.startswith('@'):
                return ('call', up, [])
            return ('var', up, 0)
        raise ParseError('unexpected %r' % (val,))


def parse_expr(toks):
    if not toks:
        raise ParseError('empty expression')
    p = _Parser(toks)
    try:
        return p.parse()
    except ParseError as e:
        near = ' '.join(v for _, v in toks[max(0, p.i - 4):p.i + 2])
        raise ParseError('%s, near "%s"' % (e, near))


def ast_vars(node, out=None):
    """Ordered unique variable names in an expression."""
    if out is None:
        out = []
    k = node[0]
    if k == 'var':
        if node[1] not in out and node[1] not in EXCLUDE:
            out.append(node[1])
    elif k == 'neg':
        ast_vars(node[1], out)
    elif k == 'bin':
        ast_vars(node[2], out)
        ast_vars(node[3], out)
    elif k == 'call':
        for a in node[2]:
            ast_vars(a, out)
    return out


def ast_has(node, kinds):
    k = node[0]
    if k in kinds:
        return True
    if k == 'neg':
        return ast_has(node[1], kinds)
    if k == 'bin':
        return ast_has(node[2], kinds) or ast_has(node[3], kinds)
    if k == 'call':
        return any(ast_has(a, kinds) for a in node[2])
    return False



def ast_text(node):
    k = node[0]
    if k == 'num':
        v = node[1]
        return ('%d' % v) if v.is_integer() else ('%g' % v)
    if k == 'var':
        return node[1] if node[2] == 0 else '%s(%d)' % (node[1], node[2])
    if k == 'coef':
        return 'C(%d)' % node[1]
    if k == 'neg':
        return '-' + ast_text(node[1])
    if k == 'bin':
        return '%s%s%s' % (ast_text(node[2]), node[1], ast_text(node[3]))
    if k == 'call':
        return '%s(%s)' % (node[1], ','.join(ast_text(a) for a in node[2]))
    if k == 'str':
        return '"%s"' % node[1]
    return '?'


class Equation:
    __slots__ = ('target', 'lhs', 'rhs', 'lhs_vars', 'rhs_vars', 'drivers', 'text',
                 'kind', 'evaluable', 'source', 'why')

    def __init__(self, target, drivers, text, kind, source, lhs=None, rhs=None,
                 lhs_vars=(), rhs_vars=(), evaluable=False, why=''):
        self.target = target
        self.drivers = list(drivers)
        self.text = text
        self.kind = kind
        self.source = source
        self.lhs, self.rhs = lhs, rhs
        self.lhs_vars, self.rhs_vars = list(lhs_vars), list(rhs_vars)
        self.evaluable = evaluable
        self.why = why


_NAME_SCAN = re.compile(r'(?<![\w$.@%#])([A-Za-z_][A-Za-z0-9_$.%#]*)')


def _scan_names(s):
    out = []
    for m in _NAME_SCAN.finditer(s):
        nm = m.group(1).upper()
        if nm in FUNCS or nm in EXCLUDE or nm in out:
            continue
        out.append(nm)
    return out


def _split_terms(s):
    terms, depth, cur = [], 0, ''
    for ch in s:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        if ch.isspace() and depth == 0:
            if cur:
                terms.append(cur)
            cur = ''
        else:
            cur += ch
    if cur:
        terms.append(cur)
    return terms


_SKIP_START = re.compile(
    r'^(for|next|if|endif|else|while|wend|smpl|include|call|subroutine|endsub|string|scalar|'
    r'vector|matrix|show|freeze|delete|wfopen|wfsave|pageselect|assign|solve|output|close|exit|'
    r'stop|logmode|model|copy|fetch|store|group|graph|table|text|coef|sample|pagecreate|wfcreate|'
    r'open|save|print|setmaxerrs|mode|tic|toc|statusline)\b', re.I)
_SPEC = re.compile(r'^(?:equation\s+)?(?:[A-Za-z_][\w$]*\.)?(ls|tsls|gmm|nls|arch|binary|logit|probit|'
                   r'glm|ardl|qreg|robustls|cointreg)\b(?:\([^)]*\))?\s*', re.I)


def _strip_comment(s):
    """Drop a trailing EViews comment (' ...) that is not inside a quoted string."""
    inq = False
    for i, ch in enumerate(s):
        if ch == '"':
            inq = not inq
        elif ch == "'" and not inq:
            return s[:i].rstrip()
    return s


def _drop_error_terms(s):
    """Remove EViews error-term options "[ar(1)=0.8, uncond, ...]" and the + that joins them to the formula.
    Scans character by character: brackets inside quotes and a missing closing ] are handled.
    Returns (formula without them, list of the removed [..] parts)."""
    out, found, i, n = [], [], 0, len(s)
    while i < n:
        if s[i] != '[':
            out.append(s[i]); i += 1; continue
        j, q = i + 1, None                      # find the matching ]
        while j < n:
            c = s[j]
            if q:
                if c == q:
                    q = None
            elif c in '"\'':
                q = c
            elif c == ']':
                break
            j += 1
        found.append(s[i:j + 1].strip() if j < n else s[i:].strip() + ']')
        head = ''.join(out).rstrip()
        if head.endswith('+'):                   # "... + [..]"
            head = head[:-1].rstrip()
            out = [head + ' ']
            i = j + 1
        else:                                    # "[..] + ..." or "= [..]"
            out = [head + ' ']
            k = j + 1
            while k < n and s[k].isspace():
                k += 1
            i = k + 1 if k < n and s[k] == '+' else j + 1
    return re.sub(r'\s{2,}', ' ', ''.join(out)).strip(), found


def parse_equation_line(raw, source):
    """Parse one logical line of EViews text. Returns Equation, ('addlink', y, af) or None."""
    s = raw.strip()
    if not s or s[0] in "'#;!%{" or s.startswith('//'):
        return None
    # characters pasted from Word / Excel
    s = (s.replace('\u2212', '-').replace('\u2013', '-').replace('\u2014', '-').replace('\u00a0', ' ')
          .replace('\u201c', '"').replace('\u201d', '"').replace('\t', ' ')
          .replace('\uff0b', '+').replace('\uff3b', '[').replace('\uff3d', ']').replace('\uff08', '(')
          .replace('\uff09', ')').replace('\uff1d', '=').replace('\uff0c', ',').replace('\u3000', ' '))
    s = re.sub('[\u200b\u200c\u200d\u2060\ufeff\u00ad]', '', s)    # invisible characters
    # scientific notation written with spaces or a D: "1.5E -05", "1.5 E-05", "1.5D-05" -> 1.5E-05
    s = re.sub(r'(?<![A-Za-z0-9_$.%#])(\d+\.?\d*|\.\d+)\s*[eEdD]\s*([-+]?)\s*(\d+)(?![A-Za-z0-9_$.%#(])',
               lambda m: '%sE%s%s' % (m.group(1), m.group(2), m.group(3)), s)
    s = _strip_comment(s)
    shown = s.strip()                       # what the user sees, error-term options included
    # error terms written as an added term: "... + [ar(1)=0.13, ...]" or "[ar(1)=.1] + ..."
    s, _br = _drop_error_terms(s)
    s = re.sub(r'[+\-*/^]\s*$', '', s.strip()).strip()   # a dangling operator left at the end
    s = re.sub(r'=\s*\+', '= ', s)
    s = re.sub(r"\s{2,}", " ", s)
    if not s or s.startswith(':'):
        return None
    m = re.match(r'^@ADD\w*\s*(?:\([^)]*\))?\s+(\S+)\s+(\S+)', s, re.I)
    if m:
        return ('addlink', m.group(1).upper(), m.group(2).upper())
    if _SKIP_START.match(s):
        return None
    kind = 'equation'
    m = re.match(r'^@IDENTITY\s+', s, re.I)
    if m:
        kind, s = 'identity', s[m.end():]
    m = re.match(r'^(series|genr|frml)\s+', s, re.I)
    if m:
        kind, s = 'identity', s[m.end():]
    spec = False
    m = _SPEC.match(s)
    if m:
        spec, s = True, s[m.end():]
        if s.lower().startswith('equation '):
            return None
    if s.startswith('@'):
        return None
    text = s.strip()
    br = _drop_error_terms(shown)[1]
    if br:
        text = text + '   ' + ' '.join(b.strip() for b in br)
    why = ''
    try:
        toks = tokenize(s)
    except ParseError as e:
        toks = None
        why = str(e)
    if toks is not None:
        depth, eqi = 0, -1
        for i, (kd, v) in enumerate(toks):
            if kd != 'op':
                continue
            if v == '(':
                depth += 1
            elif v == ')':
                depth -= 1
            elif v == '=' and depth == 0:
                eqi = i
                break
        if eqi > 0:
            try:
                lhs = parse_expr(toks[:eqi])
                rhs = parse_expr(toks[eqi + 1:])
            except ParseError as e:
                lhs = rhs = None
                why = str(e)
                if eqi + 1 >= len(toks):
                    why = 'nothing after "="'

            if lhs is not None:
                lv, rv = ast_vars(lhs), ast_vars(rhs)
                if not lv:
                    return None
                tgt = lv[0]
                drivers = rv + [v for v in lv[1:] if v not in rv]
                coef = ast_has(rhs, ('coef',)) or ast_has(lhs, ('coef',))
                return Equation(tgt, drivers, text, kind, source, lhs, rhs, lv, rv,
                                evaluable=not coef,
                                why='has unsubstituted coefficients C(i)' if coef else '')
        elif eqi < 0 and spec:
            terms = _split_terms(s)
            names = []
            for t in terms:
                try:
                    names.append(ast_vars(parse_expr(tokenize(t))))
                except ParseError:
                    names.append(_scan_names(t))
            if not names or not names[0]:
                return None
            tgt = names[0][0]
            drivers = []
            for lst in names[1:] + [names[0][1:]]:
                for v in lst:
                    if v not in drivers:
                        drivers.append(v)
            return Equation(tgt, drivers, text, 'spec', source,
                            why='specification only (no coefficients)')
    # fallback: dependencies only
    if '=' not in s:
        return None
    left, right = s.split('=', 1)
    lv, rv = _scan_names(left), _scan_names(right)
    if not lv:
        return None
    return Equation(lv[0], rv + [v for v in lv[1:] if v not in rv], text, kind, source,
                    why='could not parse: %s' % (why or 'unknown problem'))


def _eq_sig(e):
    """What counts as 'the same equation': same parsed formula (spacing, case and error-term options ignored)."""
    if e.lhs is not None and e.rhs is not None:
        return ast_text(e.lhs) + '=' + ast_text(e.rhs)
    return re.sub(r'\s+', '', _drop_error_terms(e.text)[0]).upper()


def _logical_lines(text):
    out, buf = [], ''
    for line in text.splitlines():
        st = line.strip()
        if not st:
            if buf:
                out.append(buf)
                buf = ''
            continue
        if buf and (buf.count('(') > buf.count(')') or buf.count('[') > buf.count(']') or re.search(r'[-+*/^(,=]\s*$', buf)
                    or re.match(r'^[-+*/^]', st)) and not st.startswith(("'", ':', '@')):
            buf += ' ' + st
        else:
            if buf:
                out.append(buf)
            buf = st
    if buf:
        out.append(buf)
    return out


EQ_EXT = {'.txt', '.prg', '.eq', '.eqs', '.mod', '.model', '.eviews', '.text'}


def load_equations(paths):
    eqs, addlinks, stats, fails = {}, {}, Counter(), []
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, fs in os.walk(p):
                for f in sorted(fs):
                    if os.path.splitext(f)[1].lower() in EQ_EXT:
                        files.append(os.path.join(root, f))
        else:
            files.append(p)
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext in ('.csv', '.tsv', '.xlsx', '.xlsm'):
            grid, _ = read_grid(f)
            lines = [str(c) for r in grid for c in r if c and ('=' in str(c) or _SPEC.match(str(c).strip()))]
        else:
            with open(f, 'r', encoding='utf-8-sig', errors='replace') as fh:
                lines = _logical_lines(fh.read())
        stats['files'] += 1
        for ln in lines:
            stats['lines'] += 1
            r = parse_equation_line(ln, os.path.basename(f))
            if r is None:
                if '=' in ln and len(fails) < 30 and not _SKIP_START.match(ln.strip()):
                    fails.append(ln[:160])
                continue
            if isinstance(r, tuple):
                addlinks[r[1]] = r[2]
                continue
            stats['parsed'] += 1
            prev = eqs.get(r.target)
            if prev is not None:
                stats['duplicate targets'] += 1
                if prev.evaluable and not r.evaluable:
                    continue
            eqs[r.target] = r
    stats['equations'] = len(eqs)
    stats['evaluable'] = sum(1 for e in eqs.values() if e.evaluable)
    return eqs, addlinks, stats, fails


# --------------------------------------------------------------------------
# The comparison engine
# --------------------------------------------------------------------------

def _scc_order(nodes, preds):
    """Strongly connected components of the graph (edges pred -> node), in dependency order."""
    idx, low, onst, st, comps, cnt = {}, {}, set(), [], [], [0]
    succ = defaultdict(list)
    for v in nodes:
        for u in preds.get(v, ()):
            if u in preds or u in nodes:
                succ[u].append(v)
    nodeset = set(nodes)
    for root in nodes:
        if root in idx:
            continue
        work = [(root, iter([w for w in preds.get(root, ()) if w in nodeset]))]
        idx[root] = low[root] = cnt[0]; cnt[0] += 1; st.append(root); onst.add(root)
        while work:
            v, it = work[-1]
            adv = False
            for w in it:
                if w not in idx:
                    idx[w] = low[w] = cnt[0]; cnt[0] += 1; st.append(w); onst.add(w)
                    work.append((w, iter([z for z in preds.get(w, ()) if z in nodeset])))
                    adv = True
                    break
                elif w in onst:
                    low[v] = min(low[v], idx[w])
            if adv:
                continue
            work.pop()
            if work:
                low[work[-1][0]] = min(low[work[-1][0]], low[v])
            if low[v] == idx[v]:
                c = []
                while True:
                    w = st.pop(); onst.discard(w); c.append(w)
                    if w == v:
                        break
                comps.append(c)
    return comps   # predecessors' components come first



def compute_layers(parents, area=None):
    """Longest chain of equations from each variable downstream ("layers it affects").
    parents: variable -> variables its equation uses (any lag). A loop counts as one layer, since inside
    a loop the longest chain is endless. Returns per variable:
      d    longest chain to any variable (0 = feeds nothing)
      da   longest chain that ends in `area` (None = cannot reach the area; 0 = in the area, feeds no other)
      nx / nxa   next variable on that chain;  lp  size of the loop it sits in (1 = no loop)."""
    nodes = set(parents)
    for ps in parents.values():
        nodes.update(ps)
    nodes = sorted(nodes)
    comps = _scc_order(nodes, parents)          # upstream components first
    comp = {}
    for ci, c in enumerate(comps):
        for n in c:
            comp[n] = ci
    out_edges = defaultdict(list)               # component -> [(u, v)] edges leaving it
    for v, ps in parents.items():
        cv = comp[v]
        for u in ps:
            if comp[u] != cv:
                out_edges[comp[u]].append((u, v))
    area = set(area or ())
    nc = len(comps)
    d, da, nx, nxa = [0] * nc, [None] * nc, [None] * nc, [None] * nc
    for ci in range(nc - 1, -1, -1):            # downstream components first
        if any(n in area for n in comps[ci]):
            da[ci] = 0
        for u, v in out_edges.get(ci, ()):
            cv = comp[v]
            if d[cv] + 1 > d[ci]:
                d[ci], nx[ci] = d[cv] + 1, (u, v)
            if da[cv] is not None and (da[ci] is None or da[cv] + 1 > da[ci]):
                da[ci], nxa[ci] = da[cv] + 1, (u, v)
    res = {}
    for n in nodes:
        ci = comp[n]
        res[n] = (d[ci], da[ci], nx[ci][1] if nx[ci] else None, nxa[ci][1] if nxa[ci] else None, len(comps[ci]))
    return res



def layer_chain(L, n, area=False, limit=60):
    """The longest chain from n as a list of names (loops shown by one member each)."""
    out, k = [n], 3 if area else 2
    while len(out) < limit:
        nxt = L.get(out[-1], (0, None, None, None, 1))[k]
        if not nxt or nxt in out:
            break
        out.append(nxt)
    return out



def _louvain(n, edges, gamma=1.0):
    """Louvain community detection (modularity), deterministic. edges: (i, j, w) undirected, i != j.
    Returns a community label per node 0..n-1."""
    adj = [dict() for _ in range(n)]
    for i, j, w in edges:
        if i != j:
            adj[i][j] = adj[i].get(j, 0.0) + w
            adj[j][i] = adj[j].get(i, 0.0) + w
    loops = [0.0] * n
    labels = list(range(n))
    for _level in range(12):
        k = [sum(a.values()) + loops[i] for i, a in enumerate(adj)]
        m2 = sum(k)
        if m2 <= 0:
            break
        comm, tot, improved = list(range(len(adj))), k[:], False
        for _pass in range(25):
            moved = False
            for i in range(len(adj)):
                ci, ki = comm[i], k[i]
                wc = {}
                for j, w in adj[i].items():
                    cj = comm[j]
                    wc[cj] = wc.get(cj, 0.0) + w
                tot[ci] -= ki
                best, bg = ci, wc.get(ci, 0.0) - gamma * tot[ci] * ki / m2
                for c, w in wc.items():
                    g = w - gamma * tot[c] * ki / m2
                    if g > bg + 1e-12:
                        best, bg = c, g
                tot[best] += ki
                if best != ci:
                    comm[i], moved, improved = best, True, True
            if not moved:
                break
        if not improved:
            break
        ids = {c: x for x, c in enumerate(sorted(set(comm)))}
        comm = [ids[c] for c in comm]
        labels = [comm[l] for l in labels]
        nadj, nloops = [dict() for _ in ids], [0.0] * len(ids)
        for i, a in enumerate(adj):
            ci = comm[i]
            nloops[ci] += loops[i]
            for j, w in a.items():
                cj = comm[j]
                if ci == cj:
                    nloops[ci] += w
                else:
                    nadj[ci][cj] = nadj[ci].get(cj, 0.0) + w
        adj, loops = nadj, nloops
    ids = {c: x for x, c in enumerate(sorted(set(labels)))}
    return [ids[c] for c in labels]



def _min_cover(cut):
    """Minimum vertex cover of a bipartite set of links (inside u, outside v), by Konig's theorem:
    the fewest variables that every link between a group and the rest passes through."""
    left = sorted({u for u, _ in cut})
    nb = defaultdict(list)
    for u, v in cut:
        nb[u].append(v)
    mate_l, mate_r = {}, {}
    for u in left:                                   # Kuhn's algorithm, iterative augmenting paths
        seen, stack, parent = {u}, [(u, iter(nb[u]))], {}
        found = None
        while stack and found is None:
            x, it = stack[-1]
            for v in it:
                if v in parent:
                    continue
                parent[v] = x
                if v not in mate_r:
                    found = v
                    break
                y = mate_r[v]
                if y not in seen:
                    seen.add(y)
                    stack.append((y, iter(nb[y])))
                    break
            else:
                stack.pop()
        if found is not None:
            v = found
            while True:
                x = parent[v]
                nxt = mate_l.get(x)
                mate_l[x], mate_r[v] = v, x
                if x == u:
                    break
                v = nxt
    zl, zr = set(), set()
    q = [u for u in left if u not in mate_l]
    zl.update(q)
    while q:
        x = q.pop()
        for v in nb[x]:
            if v not in zr and mate_l.get(x) != v:
                zr.add(v)
                y = mate_r.get(v)
                if y is not None and y not in zl:
                    zl.add(y)
                    q.append(y)
    return sorted(set(left) - zl) + sorted(zr)



def compute_groups(nodes, par, geo=None, boxes=40, by_geo=False, desc=None):
    """Split the equation graph into a tree of groups, so that no view shows more than `boxes` boxes.
    Each split uses Louvain community detection (dense inside, loose links to the rest), with its
    resolution tuned to give 2 .. ~2/3 boxes parts; too many parts are grouped again; a set with no
    structure is cut along the order of its equations. A group with at most `boxes` variables is a leaf.
    Every group gets its key variables: the fewest variables all links between it and the rest of the
    model pass through (minimum vertex cover of its boundary links).
    nodes: names; par: list of parent index lists. Returns {'tree': [...top groups...], 'boxes': boxes}."""
    n = len(nodes)
    boxes = max(8, int(boxes))
    maxp = max(2, boxes * 2 // 3)
    und = set()
    for v, ps in enumerate(par):
        for u in ps:
            if u != v:
                und.add((u, v) if u < v else (v, u))
    nbr = [[] for _ in range(n)]
    for a, b in und:
        nbr[a].append(b)
        nbr[b].append(a)
    order = {}
    for x, c in enumerate(_scc_order(list(range(n)), {v: par[v] for v in range(n)})):
        for v in c:
            order[v] = x

    def fam(i):                                                # concept family: geography and numbers removed
        n = nodes[i]
        g = geo(n) if geo else ''
        st = n[:-(len(g) + 1)] if g else n
        return re.sub(r'\d+$', '', st) or st

    def dfam(i):                                               # the same from the description, if any
        d = (desc(nodes[i]) if desc else '') or ''
        g = geo(nodes[i]) if geo else ''
        if g:
            d = re.sub(r'[\s,;:(\-]*%s\)?\s*$' % re.escape(g), '', d, flags=re.I)
        return re.sub(r'\s*\d+\s*$', '', d).strip(' ,;:-')

    def label(mem, pgeo=None):
        """What is in the group: its main concepts, then its regions, e.g. 'FEMP, FOUT · ICA',
        'House price index · 13 regions', 'FY · IUSA + 50 regions'. Returns (label, geography part)."""
        n = len(mem)
        gpart = ''
        if geo:
            c = Counter(geo(nodes[i]) for i in mem)
            codes = [g for g in c if g]
            g1, k1 = c.most_common(1)[0]
            if g1 and k1 >= 0.6 * n:
                gpart = g1
            elif codes:
                top, cov = [], 0
                for g2, k2 in c.most_common(3):
                    if not g2:
                        break
                    top.append(g2)
                    cov += k2
                    if cov >= 0.8 * n:
                        break
                if top and cov >= 0.8 * n:
                    gpart = ' + '.join(top)
                else:
                    big = [g for g, k in c.most_common(1) if g and k >= 0.2 * n]
                    gpart = ('%s + %d regions' % (big[0], len(codes) - 1)) if big else '%d regions' % len(codes)
        fc = Counter(fam(i) for i in mem)
        fams, cov = [], 0
        for f, k in fc.most_common(3):
            fams.append(f)
            cov += k
            if cov >= 0.7 * n:
                break

        def fname(f):
            if desc:
                ds = Counter(dfam(i) for i in mem if fam(i) == f)
                d, k = ds.most_common(1)[0]
                if d and k >= 0.5 * sum(ds.values()):
                    return d
            return f
        if cov >= 0.7 * n:
            cpart = ', '.join(fname(f) for f in fams)
        else:
            cpart = '%d concepts: %s, …' % (len(fc), ', '.join(fname(f) for f in fams[:2]))
        if len(cpart) > 44:
            cpart = cpart[:42].rstrip(', ') + '…'
        show_g = gpart and gpart != pgeo
        return (cpart + (' · ' + gpart if show_g else '')), gpart

    def keys_of(mem):
        ms = set(mem)
        cut = [(a, b) for a in mem for b in nbr[a] if b not in ms]
        if not cut:
            return []
        k_in, k_out = _min_cover(cut), _min_cover([(b, a) for a, b in cut])
        return k_out if len(k_out) <= len(k_in) else k_in

    def louvain_parts(members, gamma):
        pos = {v: x for x, v in enumerate(members)}
        e = [(pos[a], pos[b], 1.0) for a in members for b in nbr[a] if b in pos and a < b]
        lab = _louvain(len(members), e, gamma)
        out = defaultdict(list)
        for x, v in enumerate(members):
            out[lab[x]].append(v)
        return list(out.values())

    def chunk(items, k, size=lambda x: 1):
        """k bins, consecutive in the order given, of about equal total size."""
        tot = sum(size(x) for x in items)
        bins, cur, acc = [], [], 0
        for x in items:
            cur.append(x)
            acc += size(x)
            if acc >= tot * (len(bins) + 1) / k and len(bins) < k - 1:
                bins.append(cur)
                cur = []
        if cur:
            bins.append(cur)
        return bins

    def regroup(parts):
        """Too many parts: bundle parts that hang off the same neighbouring parts; then cut into blocks."""
        where = {}
        for x, m in enumerate(parts):
            for v in m:
                where[v] = x
        sig = defaultdict(list)
        for x, m in enumerate(parts):
            nb = frozenset(where[b] for a in m for b in nbr[a] if b in where and where[b] != x)
            sig[nb].append(x)
        groups = sorted(sig.values(), key=lambda xs: (-sum(len(parts[x]) for x in xs), xs[0]))
        if 1 < len(groups) <= maxp:
            return [sum((parts[x] for x in xs), []) for xs in groups]
        ps = sorted(parts, key=lambda m: min(order[v] for v in m))
        return [sum(b, []) for b in chunk(ps, maxp, size=len)]

    def split(members):
        best = None
        for g in (1.0, 2.0, 4.0, 8.0, 16.0, 32.0):            # tune the resolution: prefer balanced splits
            parts = louvain_parts(members, g)
            if len(parts) < 2:
                continue
            big = max(len(m) for m in parts) / len(members)
            if best is None or big < best[0] - 0.05:
                best = (big, parts)
            if big <= 0.5 or len(parts) > maxp:
                break
        parts = best[1] if best else []
        if len(parts) > maxp:
            parts = regroup(parts)
        if len(parts) < 2:                                     # no structure at all: cut along equation order
            ms = sorted(members, key=lambda v: (order[v], nodes[v]))
            parts = chunk(ms, min(maxp, max(2, -(-len(ms) // boxes))))
        return [sorted(m) for m in parts if m]

    def build(members, depth, pgeo=None):
        out, seen = [], Counter()
        for mem in sorted(split(members), key=lambda m: (-len(m), nodes[m[0]])):
            lab, gp = label(mem, pgeo)
            seen[lab] += 1
            if seen[lab] > 1:                                  # keep names in a view distinct
                lab = '%s (%d)' % (lab, seen[lab])
            node = {'m': mem, 'k': keys_of(mem), 'lab': lab}
            if len(mem) > boxes and depth < 30 and len(mem) < len(members):
                node['ch'] = build(mem, depth + 1, gp or pgeo)
            out.append(node)
        return out

    linked = [i for i in range(n) if nbr[i]]                   # variables with no links at all are left out
    if by_geo and geo and len(linked) > boxes:                 # --map-geo: split by geography first
        gl = {i: (geo(nodes[i]) or '') for i in linked}
        byg = defaultdict(list)
        for i in linked:
            byg[gl[i]].append(i)
        coded = sum(len(m) for g, m in byg.items() if g)
        if len([g for g in byg if g]) >= 3 and coded >= 0.5 * len(linked):
            users = defaultdict(set)                           # geography -> other geographies that use it
            for i in linked:
                for u in par[i]:
                    if u in gl and gl[u] != gl[i]:
                        users[gl[u]].add(gl[i])
            gs = sorted(byg)
            hubs = [g for g in gs if g == '' or len(users[g]) >= max(2, len(gs) / 3)]
            rest = [g for g in gs if g not in hubs]

            def gnode(g):
                mem = sorted(byg[g])
                node = {'m': mem, 'k': keys_of(mem), 'lab': g or '(no geography)'}
                if len(mem) > boxes:
                    node['ch'] = build(mem, 1, g)
                return node
            tree = [gnode(g) for g in hubs]
            room = max(2, maxp - len(tree))
            if len(rest) <= room:
                tree += [gnode(g) for g in rest]
            else:
                per = -(-len(rest) // (-(-len(rest) // maxp)))
                for b in range(0, len(rest), per):
                    blk = rest[b:b + per]
                    ch = [gnode(g) for g in blk]
                    mem = sorted(v for c in ch for v in c['m'])
                    tree.append({'m': mem, 'k': keys_of(mem), 'lab': '%s … %s (%d regions)' % (blk[0] or '-', blk[-1], len(blk)),
                                 'ch': ch})
            return {'tree': tree, 'boxes': boxes, 'by': 'geography'}
    tree = build(linked, 0) if len(linked) > boxes else (
        [{'m': linked, 'k': [], 'lab': 'whole model'}] if linked else [])
    return {'tree': tree, 'boxes': boxes}



def load_exog(items):
    names = []
    for it in items or ():
        if os.path.isfile(it):
            if os.path.splitext(it)[1].lower() in ('.xlsx', '.xlsm'):
                text = ' '.join(str(c) for r in read_xlsx(it) for c in r)
            else:
                with open(it, 'r', encoding='utf-8-sig', errors='replace') as f:
                    text = f.read()
        else:
            text = it
        for tok in re.split(r'[\s,;]+', text):
            tok = tok.strip().strip('"\'').upper()
            if tok and not tok.startswith(("'", '#')):
                names.append(tok)
    return names



# ---------------------------------------------------------------- model map

def detect_geo(names, sep='auto'):
    """Geography codes from the names: a name's last token (after '.' or '_') counts as a geography
    only if it ends the names of at least 3 different concepts (FGDP_ICAN, FLBR_ICAN, FHPI_ICAN).
    Used only to name groups. Returns a function name -> code or ''."""
    best = None
    for sp in (['.', '_'] if sep in (None, '', 'auto') else [sep]):
        stems = defaultdict(set)
        for n in names:
            if sp in n:
                st, g = n.rsplit(sp, 1)
                if g and st:
                    stems[g].add(st)
        codes = {g for g, ss in stems.items() if len(ss) >= 3}
        cover = sum(1 for n in names if sp in n and n.rsplit(sp, 1)[1] in codes)
        if best is None or cover > best[2]:
            best = (sp, codes, cover)
    if not best or not best[2]:
        return None, set()
    sp, codes = best[0], best[1]

    def geo(n):
        if sp in n:
            g = n.rsplit(sp, 1)[1]
            if g in codes:
                return g
        return ''
    return geo, codes


def gaps(A, B, names):
    """Gap B - A in % of A: in the last period both have, and the largest in absolute value."""
    out = {}
    for n in names:
        if n not in A.series or n not in B.series:
            continue
        (ba, xa), (bb, xb) = A.series[n], B.series[n]
        va = {k: v for k, v in zip(A.blocks[ba], xa) if v == v}
        vb = {k: v for k, v in zip(B.blocks[bb], xb) if v == v}
        ks = sorted(set(va) & set(vb))
        if not ks:
            continue
        g = [((vb[k] - va[k]) / abs(va[k]) * 100) if va[k] else (0.0 if vb[k] == va[k] else None) for k in ks]
        g = [x for x in g if x is not None]
        if not g:
            continue
        mx = max(g, key=abs)
        if abs(mx) < 1e-9:
            continue
        out[n] = (round(g[-1], 4), round(mx, 4))
    return out


def _read_pairs(path, sheet=None):
    """Rows of a csv/xlsx file as lists of strings (header included)."""
    grid, _ = read_grid(path, sheet)
    return [[_clean(c) if c is not None else '' for c in row] for row in grid]


class ModelMap:
    """The links of a model (which variables each variable's equation uses), its groups and its map.

    Build it from equations, from a list of links, or from a dict, then ask questions or write the page:

        import model_map as mm
        m = mm.ModelMap.from_equations(['model.txt'])
        m = mm.ModelMap.from_links([('FOIL', 'FCPI'), ('FCPI', 'FGDP')])
        m = mm.ModelMap({'FGDP': ['FC', 'FI'], 'FC': ['FYPDPI']})       # variable -> what it uses
        m.add_scenarios(['base.csv'], ['scen.xlsx'])                    # optional: gaps for colours
        m.group_path('FGDP$.IUSA'); m.upstream('FHPI.ICA', 2); m.link_path('FOIL', 'FHPI.ICA')
        m.to_html('map.html'); m.to_json('map.json')
    """

    def __init__(self, parents, equations=None, descriptions=None, exog=(), boxes=40, by_geo=False,
                 geo_sep='auto', geo_names=True, af_suffix='_A'):
        exog = {str(x).upper() for x in exog}
        par = defaultdict(set)
        for n, ps in dict(parents).items():
            n = str(n).upper()
            par[n]                                             # keep variables with no inputs
            if n in exog:
                continue
            par[n].update(str(p).upper() for p in ps if str(p).upper() != n)
        names = set(par)
        for ps in par.values():
            names.update(ps)
        self._par = par
        self._extra = set()
        self.names = names
        self.equations = {str(k).upper(): v for k, v in (equations or {}).items()}
        self.descriptions = {str(k).upper(): v for k, v in (descriptions or {}).items()}
        self.exog = exog
        self.gap = {}                                          # name -> (gap at end %, largest gap %)
        self.data_files = ([], [])
        self.boxes, self.by_geo, self.geo_sep, self.geo_names = boxes, by_geo, geo_sep, geo_names
        self.af_suffix = [x.strip().upper() for x in (af_suffix or '').split(',') if x.strip()]
        self.n_equations = len(self.equations) or sum(1 for n in par if par[n])
        self.unread = []
        self._cache = {}

    # ---- constructors
    @classmethod
    def from_equations(cls, paths, extra=None, exog=None, **kw):
        """EViews-style equation files or folders. `extra` files override equations of the same variable."""
        paths = [paths] if isinstance(paths, str) else list(paths)
        eqs, addl, stats, fails = load_equations(paths)
        for e in ([extra] if isinstance(extra, str) else (extra or [])):
            e2, a2, _, f2 = load_equations([e])
            eqs.update(e2)
            addl.update(a2)
            fails += f2
        ex = set(load_exog(exog if isinstance(exog, (list, tuple)) else ([exog] if exog else [])))
        m = cls({n: e.drivers for n, e in eqs.items()}, equations={n: e.text for n, e in eqs.items()}, exog=ex, **kw)
        m.n_equations = len(eqs)
        m.unread = fails
        m.sources = [os.path.basename(p) for p in paths]
        return m

    @classmethod
    def from_links(cls, links, **kw):
        """Links as (from, to) pairs: `from` is used in the equation of `to`."""
        par = defaultdict(set)
        for a, b in links:
            par[str(b).upper()].add(str(a).upper())
            par[str(a).upper()]
        return cls(par, **kw)

    @classmethod
    def from_links_csv(cls, path, src=None, dst=None, sheet=None, **kw):
        """A csv/xlsx table of links. Columns named from/to, source/target, driver/variable or input/output
        (or give src= and dst=); otherwise the first two columns. A third column 'description' is ignored."""
        rows = [r for r in _read_pairs(path, sheet) if any(r)]
        if not rows:
            raise ValueError('%s: empty' % path)
        h = [c.strip().lower() for c in rows[0]]
        pairs = [('from', 'to'), ('source', 'target'), ('driver', 'variable'), ('input', 'output'), ('x', 'y')]
        i = j = None
        if src and dst:
            i, j = h.index(src.lower()), h.index(dst.lower())
        else:
            for a, b in pairs:
                if a in h and b in h:
                    i, j = h.index(a), h.index(b)
                    break
        body = rows[1:] if i is not None else rows
        if i is None:
            i, j = 0, 1
            if h[0] in ('from', 'source') or not any(ch.isdigit() or ch.isalpha() for ch in rows[0][0]):
                body = rows[1:]
        m = cls.from_links([(r[i], r[j]) for r in body if len(r) > max(i, j) and r[i] and r[j]], **kw)
        m.sources = [os.path.basename(path)]
        return m

    # ---- optional extras
    def add_scenarios(self, a_paths, b_paths):
        """Read scenario A and B (csv/xlsx) and keep each variable's gap B - A in % of A."""
        a_paths = [a_paths] if isinstance(a_paths, str) else list(a_paths)
        b_paths = [b_paths] if isinstance(b_paths, str) else list(b_paths)
        A, B = load_scenario(a_paths, 'A'), load_scenario(b_paths, 'B')
        self._extra.update(A.series)
        self._extra.update(B.series)
        for sc in (A, B):
            for n, d in sc.desc.items():
                self.descriptions.setdefault(n, d)
        self.gap = gaps(A, B, set(A.series) | set(B.series))
        self.data_files = ([os.path.basename(p) for p in a_paths], [os.path.basename(p) for p in b_paths])
        self._cache.clear()
        return self

    def set_gaps(self, mapping):
        """Your own numbers to colour the boxes: name -> value (or (value at end, largest))."""
        self.gap = {str(k).upper(): (tuple(v) if isinstance(v, (list, tuple)) else (float(v), float(v)))
                    for k, v in mapping.items() if v is not None}
        self._cache.clear()
        return self

    def set_descriptions(self, mapping_or_path):
        """name -> description, or a csv/xlsx file with name and description columns."""
        if isinstance(mapping_or_path, str):
            rows = [r for r in _read_pairs(mapping_or_path) if len(r) >= 2 and r[0]]
            if rows and rows[0][0].strip().lower() in ('name', 'variable', 'mnemonic', 'series'):
                rows = rows[1:]
            mapping_or_path = {r[0]: r[1] for r in rows}
        self.descriptions.update({str(k).upper(): str(v) for k, v in mapping_or_path.items() if v})
        return self

    # ---- the graph
    @property
    def nodes(self):
        if 'nodes' not in self._cache:
            allv = self.names | self._extra
            af = self.add_factors_of(allv)
            self._cache['af'] = af
            drop = {a for v in af.values() for a in v}
            self._cache['nodes'] = sorted(allv - drop)
        return self._cache['nodes']

    def add_factors_of(self, allv=None):
        """Add-factor series (e.g. FGDP.IUSA_A or FGDP_A.IUSA for FGDP.IUSA) that no equation uses:
        {variable: [its add-factor series]}. They are left out of the map and shown with their variable."""
        if allv is None:
            return self._cache.get('af') or (self.nodes and self._cache.get('af', {}))
        used = set()
        for ps in self._par.values():
            used.update(ps)
        out = defaultdict(list)
        for n in sorted(allv):
            if n in used or self._par.get(n):
                continue
            for sf in self.af_suffix:
                bases = [n[:-len(sf)]] if n.endswith(sf) else []
                for sep in ('.', '_'):
                    if sf + sep in n:
                        bases.append(n.replace(sf + sep, sep, 1))
                b = next((x for x in bases if x and x in allv and x != n), None)
                if b:
                    out[b].append(n)
                    break
        return dict(out)

    def _ix(self):
        if 'ix' not in self._cache:
            nodes = self.nodes
            idx = {n: i for i, n in enumerate(nodes)}
            P = [sorted(idx[p] for p in self._par.get(n, ()) if p in idx) for n in nodes]
            C = [[] for _ in nodes]
            for v, ps in enumerate(P):
                for u in ps:
                    C[u].append(v)
            self._cache['ix'] = (idx, P, C)
        return self._cache['ix']

    def parents(self, name):
        """Variables that `name`'s equation uses."""
        return sorted(self._par.get(name.upper(), ()))

    def children(self, name):
        """Variables whose equations use `name`."""
        idx, P, C = self._ix()
        i = idx.get(name.upper())
        return [] if i is None else [self.nodes[v] for v in C[i]]

    def upstream(self, name, depth=None):
        """Everything that feeds `name`, within `depth` steps (None = all): {variable: steps}."""
        return self._walk(name, depth, up=True)

    def downstream(self, name, depth=None):
        """Everything `name` feeds, within `depth` steps (None = all): {variable: steps}."""
        return self._walk(name, depth, up=False)

    def _walk(self, name, depth, up):
        idx, P, C = self._ix()
        i = idx.get(name.upper())
        if i is None:
            raise KeyError(name)
        nb = P if up else C
        seen, q = {i: 0}, deque([i])
        while q:
            u = q.popleft()
            if depth is not None and seen[u] >= depth:
                continue
            for v in nb[u]:
                if v not in seen:
                    seen[v] = seen[u] + 1
                    q.append(v)
        del seen[i]
        return {self.nodes[v]: d for v, d in sorted(seen.items(), key=lambda kv: (kv[1], self.nodes[kv[0]]))}

    def link_path(self, src, dst):
        """Shortest chain of equations from `src` to `dst` (list of names), or None if `src` cannot reach it."""
        idx, P, C = self._ix()
        a, b = idx.get(src.upper()), idx.get(dst.upper())
        if a is None or b is None:
            raise KeyError(src if a is None else dst)
        prev, q = {a: None}, deque([a])
        while q:
            u = q.popleft()
            if u == b:
                out = []
                while u is not None:
                    out.append(self.nodes[u])
                    u = prev[u]
                return out[::-1]
            for v in C[u]:
                if v not in prev:
                    prev[v] = u
                    q.append(v)
        return None

    def layers(self):
        """name -> longest chain of equations downstream (a loop counts as one layer)."""
        if 'layers' not in self._cache:
            L = compute_layers({n: list(self._par.get(n, ())) for n in self.nodes})
            self._cache['layers'] = L
        return {n: v[0] for n, v in self._cache['layers'].items()}

    def _geo(self):
        if 'geo' not in self._cache:
            self._cache['geo'] = detect_geo(self.nodes, self.geo_sep) if self.geo_names else (None, set())
        return self._cache['geo']

    def _tree(self):
        key = ('tree', self.boxes, self.by_geo)
        if key not in self._cache:
            idx, P, C = self._ix()
            self._cache[key] = compute_groups(self.nodes, P, geo=self._geo()[0], boxes=self.boxes, by_geo=self.by_geo,
                                              desc=(self.descriptions.get if self.descriptions else None))
        return self._cache[key]

    def groups(self):
        """The group tree with names: [{'name', 'size', 'keys', 'members', 'groups': [...]}, ...].
        Variables in no equation at all are in a last group '(not linked)'."""
        nodes = self.nodes

        def conv(g):
            out = {'name': g['lab'], 'size': len(g['m']), 'keys': [nodes[i] for i in g['k']],
                   'members': [nodes[i] for i in g['m']]}
            if g.get('ch'):
                out['groups'] = [conv(c) for c in g['ch']]
            return out
        T = self._tree()['tree']
        res = [conv(g) for g in T]
        seen = {i for g in T for i in g['m']}
        rest = [n for i, n in enumerate(nodes) if i not in seen]
        if rest:
            res.append({'name': '(not linked)', 'size': len(rest), 'keys': [], 'members': rest})
        return res

    def group_path(self, name):
        """The groups a variable sits in, from the top: ['IAK … IMS (26 regions)', 'ICA', ...]."""
        name = name.upper()

        def find(gs, acc):
            for g in gs:
                if name in g['members']:
                    a = acc + [g['name']]
                    return find(g.get('groups') or [], a) or a
            return None
        return find(self.groups(), []) or []

    def key_variables(self, group_name):
        """Key variables of the first group with this name (the fewest variables all its outside links pass through)."""
        def find(gs):
            for g in gs:
                if g['name'] == group_name:
                    return g['keys']
                r = find(g.get('groups') or [])
                if r is not None:
                    return r
        return find(self.groups())

    def table(self):
        """One row per variable: dict with level_1.., variable, description, layers, gap_end, key."""
        rows, L = [], self.layers()

        def walk(g, path):
            p = path + [g['name']]
            if g.get('groups'):
                for c in g['groups']:
                    walk(c, p)
            else:
                ks = set(g['keys'])
                for n in g['members']:
                    r = {'level_%d' % (k + 1): x for k, x in enumerate(p)}
                    r.update(variable=n, description=self.descriptions.get(n, ''), layers=L.get(n, 0),
                             gap_end=self.gap[n][0] if n in self.gap else None, key=n in ks)
                    rows.append(r)
        for g in self.groups():
            walk(g, [])
        return rows

    # ---- output
    def summary(self):
        G = self.groups()

        def depth(gs):
            return 1 + max((depth(g['groups']) for g in gs if g.get('groups')), default=0) if gs else 0
        L = self.layers()
        return {'variables': len(self.nodes), 'equations': self.n_equations, 'top_groups': len(G),
                'levels': depth(G), 'boxes_per_view': self.boxes, 'longest_chain': max(L.values(), default=0),
                'changed': len(self.gap), 'unread_lines': len(self.unread),
                'add_factor_series_left_out': sum(len(v) for v in (self.add_factors_of() or {}).values())}

    def to_json(self, path=None):
        """Groups, layers, links and gaps as JSON (written to `path` if given; the dict is returned)."""
        out = {'summary': self.summary(), 'groups': self.groups(), 'layers': self.layers(),
               'links': {n: self.parents(n) for n in self.nodes if self._par.get(n)},
               'gaps': {n: list(v) for n, v in self.gap.items()}, 'descriptions': self.descriptions}
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(out, f, indent=1)
        return out

    def to_csv(self, path):
        """Every variable with its group path (level_1, level_2, ...), layers, gap and key flag."""
        rows = self.table()
        nl = max((sum(1 for k in r if k.startswith('level_')) for r in rows), default=0)
        cols = ['level_%d' % (k + 1) for k in range(nl)] + ['variable', 'description', 'layers', 'gap_end', 'key']
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(cols)
            for r in rows:
                w.writerow(['' if r.get(c) is None else r.get(c, '') for c in cols])
        return path

    def payload(self, title=None):
        nodes = self.nodes
        idx, P, C = self._ix()
        self.layers()
        L = self._cache['layers']
        comps = _scc_order(list(range(len(nodes))), {v: P[v] for v in range(len(nodes))})
        cid = [0] * len(nodes)
        for ci, c in enumerate(comps):
            for v in c:
                cid[v] = ci
        T = self._tree()
        has = bool(self.gap)
        return {
            'meta': {'title': title or ', '.join(getattr(self, 'sources', []) or ['model']),
                     'generated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
                     'eq_files': getattr(self, 'sources', []), 'has_data': has,
                     'a': self.data_files[0], 'b': self.data_files[1],
                     'n_eq': self.n_equations, 'n_fail': len(self.unread), 'exog': sorted(self.exog)},
            'nodes': nodes, 'par': P,
            'eq': [str(self.equations.get(n, ''))[:600] for n in nodes],
            'desc': [self.descriptions.get(n, '') for n in nodes] if self.descriptions else None,
            'ex': [1 if n in self.exog else 0 for n in nodes],
            'af': [self._cache.get('af', {}).get(n) for n in nodes] if self._cache.get('af') else None,
            'gap': [list(self.gap[n]) if n in self.gap else None for n in nodes] if has else None,
            'lay': [L.get(n, (0,))[0] for n in nodes], 'lp': [L.get(n, (0, 0, 0, 0, 1))[4] for n in nodes],
            'cid': cid, 'tree': T['tree'], 'boxes': T['boxes'],
        }

    def to_html(self, path, title=None):
        """Write the interactive map page."""
        data = self.payload(title)
        html = HTML.replace('__TITLE__', _esc(data['meta']['title'])).replace(
            '__DATA__', json.dumps(data, separators=(',', ':')).replace('</', '<\\/'))
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)
        return path


def find_variable(name, eq_paths=(), a_paths=(), b_paths=(), link_paths=(), out=print):
    """Explain where a variable is: in which equation lines, in which data file/sheet/row/column, whether it
    was loaded, and the closest names if it is nowhere. For "why is X not in my map?"."""
    import difflib
    q = name.strip().upper()
    out('Looking for %s' % q)
    loaded = set()
    # equations
    if eq_paths:
        eqs, _, _, fails = load_equations(list(eq_paths))
        loaded.update(eqs)
        hits = 0
        files = []
        for pth in eq_paths:
            if os.path.isdir(pth):
                for root, _, fs in os.walk(pth):
                    files += [os.path.join(root, f) for f in sorted(fs) if os.path.splitext(f)[1].lower() in EQ_EXT]
            else:
                files.append(pth)
        rx = re.compile(r'(?<![\w$.@%#])' + re.escape(q) + r'(?![\w$.%#])', re.I)
        for f in files:
            if os.path.splitext(f)[1].lower() in ('.xlsx', '.xlsm', '.csv'):
                continue
            with open(f, 'r', encoding='utf-8-sig', errors='replace') as fh:
                for k, ln in enumerate(fh, 1):
                    if rx.search(ln):
                        hits += 1
                        if hits <= 12:
                            out('  equations  %s line %d: %s' % (os.path.basename(f), k, ln.strip()[:150]))
        if q in eqs:
            e = eqs[q]
            out('  -> has an equation%s; it uses %d variables' % ('' if e.evaluable else ' (could not be fully read: %s)' % e.why,
                                                              len(e.drivers)))
        users = [n for n, e in eqs.items() if q in e.drivers]
        for n, e in eqs.items():
            loaded.update(e.drivers)
        if users:
            out('  -> used in the equations of %d variables, e.g. %s' % (len(users), ', '.join(sorted(users)[:6])))
        if not hits:
            out('  not in the equation files')
        elif q not in eqs and not users:
            out('  -> appears in the equation files but not as a variable the reader picked up '
                '(check that line with: python model_map.py --eq ... --unread)')
    # data
    for label, paths in (('A', a_paths), ('B', b_paths)):
        for pth in paths or ():
            ext = os.path.splitext(pth)[1].lower()
            sheets = list(enumerate(xlsx_sheets(pth), 1)) if ext in ('.xlsx', '.xlsm') else [(None, None)]
            found = 0
            for k, nm in sheets:
                grid = read_xlsx(pth, k) if k else read_grid(pth)[0]
                for r, row in enumerate(grid, 1):
                    for c, cell in enumerate(row, 1):
                        v = _clean(cell) if cell is not None else ''
                        if v and v.strip().upper() == q:
                            found += 1
                            vals = [x for x in row[c:c + 400] if x not in (None, '') and not _isnan(_num(x))]
                            out('  data %s  %s%s row %d column %d: %r%s, %d numbers on that row' % (
                                label, os.path.basename(pth), (' [sheet %s]' % nm) if nm else '', r, c, v,
                                '' if v == q else ' (spaces or case differ)', len(vals)))
            if not found:
                out('  data %s  not in %s' % (label, os.path.basename(pth)))
        if paths:
            sc = load_scenario(list(paths), label)
            loaded.update(sc.series)
            if q in sc.series:
                out('  -> loaded from scenario %s' % label)
            else:
                out('  -> NOT loaded from scenario %s. The reader notes:' % label)
                for nt in sc.notes:
                    out('       ' + nt)
    for pth in link_paths or ():
        m = ModelMap.from_links_csv(pth)
        loaded.update(m.names)
        out('  links  %s: %s' % (os.path.basename(pth), 'found' if q in m.names else 'not found'))
    if q in loaded:
        out('=> %s is in the map. In the page, type it in the search box.' % q)
    else:
        close = difflib.get_close_matches(q, sorted(loaded), n=6, cutoff=0.75)
        out('=> %s is not in the map.%s' % (q, (' Closest names: ' + ', '.join(close)) if close else ''))
    return q in loaded


def build(args):
    """Command line: build a ModelMap from the arguments and write what was asked for."""
    say = (lambda m: None) if args.quiet else (lambda m: print(m, file=sys.stderr))
    t0 = time.time()
    kw = dict(boxes=args.map_boxes, by_geo=args.map_geo, geo_sep=args.geo_sep, geo_names=not args.no_geo_names,
              af_suffix=args.af_suffix)
    if args.links:
        say('Reading links ...')
        m = ModelMap.from_links_csv(args.links[0], **kw)
        for extra in args.links[1:]:
            m2 = ModelMap.from_links_csv(extra)
            for n, ps in m2._par.items():
                m._par[n].update(ps)
            m.names |= m2.names
            m._cache.clear()
    else:
        say('Reading equations ...')
        m = ModelMap.from_equations(args.eq, extra=args.eq_extra, exog=args.exog, **kw)
    if args.a:
        say('Reading scenarios ...')
        m.add_scenarios(args.a, args.b)
    if args.desc:
        m.set_descriptions(args.desc)
    S = m.summary()
    say('%d variables, %d equations, %d groups at the top, %d levels, at most %d boxes per view, longest chain %d%s'
        % (S['variables'], S['equations'], S['top_groups'], S['levels'], S['boxes_per_view'], S['longest_chain'],
           ('; %d add-factor series shown with their variables' % S['add_factor_series_left_out'])
           if S['add_factor_series_left_out'] else ''))
    if args.json:
        m.to_json(args.json)
        say('JSON written: %s' % args.json)
    if args.csv:
        m.to_csv(args.csv)
        say('CSV written: %s' % args.csv)
    if args.out or not (args.json or args.csv or args.groups):
        out = args.out or 'model_map.html'
        m.to_html(out, args.title)
        say('Map written: %s (%.1f MB) in %.1fs' % (out, os.path.getsize(out) / 1e6, time.time() - t0))
    if args.groups:
        print_groups(m.groups(), args.depth)
    return m


def print_groups(groups, depth):
    print('Groups. Key variables = the fewest variables that every link between the group and the rest of '
          'the model passes through.')

    def walk(gs, ind, d):
        for g in gs:
            ks = g['keys']
            print('%s%-44s %6d variables  %d key%s: %s%s' % ('  ' * ind, g['name'][:44], g['size'], len(ks),
                  '' if len(ks) == 1 else 's', ', '.join(ks[:6]), ', ...' if len(ks) > 6 else ''))
            if g.get('groups') and d < depth:
                walk(g['groups'], ind + 1, d + 1)
    walk(groups, 0, 1)


def _esc(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


HTML = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Model map: __TITLE__</title>
<style>
:root{--bg:#fff;--fg:#1d2433;--mut:#667085;--line:#e4e7ec;--card:#f7f8fa;--a:#2563eb;--b:#d9480f;--pos:#0f9d58;--neg:#d92d20;--hl:#fff5d6;--nf:#ececff;--ns:#9370db;--cf:#fffde8;--cs:#c9c26a;--ln:#4b5563}
@media (prefers-color-scheme:dark){:root{--bg:#0f1218;--fg:#e6e9ef;--mut:#98a2b3;--line:#2a2f3a;--card:#171b23;--a:#6c9cff;--b:#ff8a4c;--pos:#3ccb7f;--neg:#ff6b5e;--hl:#3a3216;--nf:#262446;--ns:#8f7ad6;--cf:#25230f;--cs:#7d7637;--ln:#aab2c0}}
*{box-sizing:border-box}body{margin:0;font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--fg)}
header{padding:14px 20px;border-bottom:1px solid var(--line)}h1{font-size:19px;margin:0}h2{font-size:16px;margin:22px 0 6px}
main{padding:12px 20px 40px;max-width:1500px;margin:auto}.hint{color:var(--mut);font-size:12.5px;margin:4px 0}.small{font-size:12.5px}
.ctl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:8px 0}.btn{border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:7px;padding:4px 10px;cursor:pointer;font:inherit;font-size:13px}
input[type=text],select{border:1px solid var(--line);background:var(--bg);color:var(--fg);border-radius:7px;padding:5px 8px;font:inherit;font-size:13px}
.badge{display:inline-block;border:1px solid var(--line);background:var(--card);border-radius:999px;padding:1px 9px;font-size:12px;margin:2px}
a{color:var(--a);text-decoration:none}a:hover{text-decoration:underline}.pos{color:var(--pos)}.neg{color:var(--neg)}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{border-bottom:1px solid var(--line);padding:4px 7px;text-align:left;vertical-align:top}th{font-weight:600;color:var(--mut);font-size:12px}td.n,th.n{text-align:right;white-space:nowrap}
.tw{overflow:auto;max-height:620px;border:1px solid var(--line);border-radius:10px}
#panel{position:fixed;right:14px;top:14px;bottom:14px;width:min(430px,92vw);background:var(--bg);border:1px solid var(--line);border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,.18);padding:14px 16px;overflow:auto;display:none;z-index:5}
#panel pre{white-space:pre-wrap;word-break:break-word;background:var(--card);border-radius:8px;padding:8px;font-size:12px}
.x{float:right;cursor:pointer;font-size:18px;color:var(--mut)}
</style></head><body>
<header><h1>Model map: __TITLE__</h1><div class="hint" id="sub"></div></header>
<main>
<div class="ctl"><label>View <select id="mby"><option value="fold">all variables, unfold step by step</option><option value="route">main chain, zoomable</option><option value="trace">around a variable</option></select></label>
 <span id="mcf"><button class="btn" id="ffold">fold all</button> <button class="btn" id="fone">unfold one more level</button> <label class="small"><input type="checkbox" id="fall"> all links</label> <button class="btn" id="fmer">copy as Mermaid</button></span>
 <span id="mct">feeding it: <select id="tu"><option>0</option><option>1</option><option selected>2</option><option>3</option><option>4</option></select> steps · fed by it: <select id="td"><option>0</option><option>1</option><option selected>2</option><option>3</option><option>4</option></select> steps</span>
 <span style="position:relative"><input type="text" id="find" placeholder="find a variable by name or description" style="min-width:290px" autocomplete="off"><div id="fl" style="position:absolute;z-index:4;top:32px;left:0;min-width:420px;max-height:340px;overflow:auto;background:var(--bg);border:1px solid var(--line);border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,.15);display:none"></div></span> <span id="findn" class="hint"></span>
 <span class="zb"><button class="btn" data-z="in">+</button><button class="btn" data-z="out">−</button><button class="btn" data-z="reset">reset view</button></span></div>
<div class="hint" id="hint"></div>
<div class="small" id="crumb" style="margin:6px 0"></div>
<div class="hint" id="note"></div>
<div id="map"></div>
<div class="small" id="chain" style="margin:8px 0"></div>
<h2>All variables, by group</h2><div class="hint">Every variable, folded by the same groups as the map. Click ▸ to open a group.</div>
<div class="ctl"><button class="btn" id="tx1">open one level</button> <button class="btn" id="tx0">fold all</button></div>
<div class="tw"><table><thead><tr><th>Group / variable</th><th class="n">Variables</th><th class="n" id="thch">Changed</th><th class="n" id="thg">Gap at end</th><th class="n">Key variables</th><th class="n">Layers</th></tr></thead><tbody id="ttb"></tbody></table></div>
<p class="hint" id="foot"></p>
</main>
<div id="panel"></div>
<script>
const D=__DATA__;
const $=s=>document.querySelector(s),esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const N=D.nodes.length,NI={};D.nodes.forEach((n,i)=>NI[n]=i);const CH=D.nodes.map(()=>[]);D.par.forEach((ps,v)=>ps.forEach(u=>CH[u].push(v)));
const HAS=!!D.gap,gapOf=i=>HAS&&D.gap[i]?D.gap[i][0]:null,pct=v=>v==null?'–':(v>0?'+':'')+(Math.abs(v)>=100?v.toFixed(0):v.toFixed(2))+'%',cls=v=>v==null?'':v<0?'neg':v>0?'pos':'';
const descOf=i=>D.desc?D.desc[i]||'':'';
const vlink=i=>`<a href="#" data-vi="${i}" title="${esc(descOf(i))}">${esc(D.nodes[i])}</a>`;
const fillOf=g=>{if(g==null)return'var(--card)';const t=Math.min(1,Math.abs(g)/3);return g<0?`rgba(217,45,32,${.12+.55*t})`:`rgba(15,157,88,${.12+.55*t})`};
const M=D.meta;$('#sub').textContent=`${N.toLocaleString()} variables, ${M.n_eq.toLocaleString()} equations from ${M.eq_files.join(', ')}${HAS?` · colours: gap B − A at the end (B: ${M.b.join(', ')}; A: ${M.a.join(', ')})`:''}${M.n_fail?` · ${M.n_fail} lines could not be read`:''} · made ${M.generated}`;
if(!HAS){$('#thch').style.display='none';$('#thg').style.display='none'}
$('#foot').textContent='Groups: Louvain community detection on the equation links, split again until each view has at most '+D.boxes+' boxes. Key variables: minimum vertex cover of a group\'s boundary links (König). Layers: longest chain of equations downstream (a loop counts as one). Main chain: search path count, key route at the top and widest path inside each step.';
// ---------- the group tree
const T=D.tree||[],inT=new Set();T.forEach(g=>g.m.forEach(i=>inT.add(i)));const rest=[];for(let i=0;i<N;i++)if(!inT.has(i))rest.push(i);
const TOP=T.slice();if(rest.length)TOP.push({m:rest,k:[],lab:'(not linked)'});
let fid=0;const PARENT=new Map(),BYF=new Map();const w0=(g,top,up)=>{g.__f='f'+(fid++);g.__top=top;PARENT.set(g,up);BYF.set(g.__f,g);(g.ch||[]).forEach(c=>w0(c,top,g))};TOP.forEach((g,x)=>w0(g,x,null));
const PATH=new Map();const pv=(g,a)=>{const b=[...a,g];if(g.ch)g.ch.forEach(c=>pv(c,b));else g.m.forEach(v=>PATH.set(v,b))};TOP.forEach(g=>pv(g,[]));
const changed=g=>HAS?g.m.filter(i=>D.gap[i]).length:null;
// ---------- pan and zoom
const VS={v:null,full:null};
function panZoom(sv){const set=()=>sv.setAttribute('viewBox',VS.v.join(' '));
 const zoomAt=(f,cx,cy)=>{const [x,y,w,h]=VS.v,nw=Math.min(VS.full[2]*2,Math.max(VS.full[2]/40,w*f)),k=nw/w;VS.v=[cx-(cx-x)*k,cy-(cy-y)*k,nw,h*k];set()};
 const toSvg=(ex,ey)=>{const r=sv.getBoundingClientRect(),[x,y,w,h]=VS.v,sc=Math.max(w/r.width,h/r.height),ox=(r.width*sc-w)/2,oy=(r.height*sc-h)/2;return[x-ox+(ex-r.left)*sc,y-oy+(ey-r.top)*sc,sc]};
 sv.addEventListener('wheel',e=>{e.preventDefault();const [cx,cy]=toSvg(e.clientX,e.clientY);zoomAt(e.deltaY>0?1.15:1/1.15,cx,cy)},{passive:false});
 let drag=null,moved=false;sv.addEventListener('pointerdown',e=>{drag=[e.clientX,e.clientY,VS.v.slice()];moved=false});
 window.addEventListener('pointermove',e=>{if(!drag)return;const sc=toSvg(0,0)[2],dx=(e.clientX-drag[0])*sc,dy=(e.clientY-drag[1])*sc;if(Math.abs(dx)+Math.abs(dy)>2)moved=true;VS.v=[drag[2][0]-dx,drag[2][1]-dy,drag[2][2],drag[2][3]];set()});
 window.addEventListener('pointerup',()=>{drag=null});sv.addEventListener('click',e=>{if(moved){e.stopPropagation();e.preventDefault();moved=false}},true);
 document.querySelectorAll('[data-z]').forEach(b=>b.onclick=()=>{const [x,y,w,h]=VS.v,z=b.dataset.z;if(z==='reset'){VS.v=VS.full.slice();set()}else zoomAt(z==='in'?1/1.3:1.3,x+w/2,y+h/2)})}
// ---------- layered layout helpers
function scc(ids,adj){let ix=0;const I=new Map(),low=new Map(),on=new Set(),st=[],comp=new Map();let nc=0;
 function sc(v){I.set(v,ix);low.set(v,ix);ix++;st.push(v);on.add(v);for(const w of adj.get(v)||[]){if(!I.has(w)){sc(w);low.set(v,Math.min(low.get(v),low.get(w)))}else if(on.has(w))low.set(v,Math.min(low.get(v),I.get(w)))}
  if(low.get(v)===I.get(v)){let w;do{w=st.pop();on.delete(w);comp.set(w,nc)}while(w!==v);nc++}}
 ids.forEach(v=>{if(!I.has(v))sc(v)});return comp}
function layerCols(ids,edges){const adj=new Map();edges.forEach(([u,v])=>{if(!adj.has(u))adj.set(u,[]);adj.get(u).push(v)});
 const comp=scc(ids,adj),nc=Math.max(-1,...comp.values())+1,cin=Array.from({length:nc},()=>new Set());
 edges.forEach(([u,v])=>{const a=comp.get(u),b=comp.get(v);if(a!==b)cin[b].add(a)});const L=new Map();
 for(let c=nc-1;c>=0;c--){let m=0;cin[c].forEach(a=>{m=Math.max(m,(L.get(a)??0)+1)});L.set(c,m)}
 const col=new Map();ids.forEach(v=>col.set(v,L.get(comp.get(v))||0));return col}
// ---------- all variables, unfold step by step
const FOLD=new Set();if(TOP.length===1&&(TOP[0].ch||TOP[0].m.length>1))FOLD.add(TOP[0]);let HL=null;
const hue=t=>t<0?'var(--mut)':`hsl(${(t*67)%360},62%,48%)`;
// ---------- all variables as a flowchart with nested, foldable frames (Mermaid-style)
function flowOrder(K,E){const outA=Array.from({length:K},()=>[]),inA=Array.from({length:K},()=>[]),oW=new Float64Array(K),iW=new Float64Array(K);
 E.forEach(([a,b,w])=>{outA[a].push([b,w]);inA[b].push([a,w]);oW[a]+=w;iW[b]+=w});
 const alive=new Uint8Array(K).fill(1),left=[],right=[];let n=K;const drop=x=>{alive[x]=0;n--;outA[x].forEach(([b,w])=>{iW[b]-=w});inA[x].forEach(([a,w])=>{oW[a]-=w})};
 while(n>0){let ch=true;while(ch){ch=false;for(let x=0;x<K;x++)if(alive[x]&&oW[x]<=1e-9){right.push(x);drop(x);ch=true}for(let x=0;x<K;x++)if(alive[x]&&iW[x]<=1e-9){left.push(x);drop(x);ch=true}}
  if(n<=0)break;let b=-1,bv=-Infinity;for(let x=0;x<K;x++)if(alive[x]&&oW[x]-iW[x]>bv){bv=oW[x]-iW[x];b=x}left.push(b);drop(b)}
 const order=[...left,...right.reverse()],rank=new Int32Array(K);order.forEach((x,r)=>rank[x]=r);
 const lay=new Int32Array(K);order.forEach(x=>{inA[x].forEach(([a])=>{if(rank[a]<rank[x])lay[x]=Math.max(lay[x],lay[a]+1)})});
 return {order,rank,lay,inA,outA}}
const NB={w:206,h:58},VB={w:176,h:30},PAD=14,TITLE=28,GAPX=56,GAPY=16;
function elemsOf(list){return list.map(g=>FOLD.has(g)?{k:'c',g,kids:g.ch?elemsOf(g.ch):g.m.map(i=>({k:'v',i}))}:{k:'g',g})}
function membersOf(e){return e.k==='v'?[e.i]:e.g.m}
function layout(cont){ // cont: {kids}; sets e.w,e.h on kids and e.x,e.y relative to cont's inner origin; returns size
 cont.kids.forEach(e=>{if(e.k==='c'){const sz=layout(e);e.w=sz.w+2*PAD;e.h=sz.h+2*PAD+TITLE}else{const b=e.k==='g'?NB:VB;e.w=b.w;e.h=b.h}});
 const K=cont.kids.length,own=new Map();cont.kids.forEach((e,x)=>membersOf(e).forEach(i=>own.set(i,x)));
 const W=new Map();own.forEach((x,v)=>{for(const u of D.par[v]){const y=own.get(u);if(y==null||y===x)continue;const k=y*K+x;W.set(k,(W.get(k)||0)+1)}});
 const E=[...W.entries()].map(([k,w])=>[Math.floor(k/K),k%K,w]),F=flowOrder(K,E);
 let L=Math.max(0,...F.lay)+1;const maxL=Math.max(2,Math.min(8,Math.ceil(Math.sqrt(K)*1.2)));const lay=Array.from(F.lay,v=>L>maxL?Math.floor(v*maxL/L):v);L=Math.min(L,maxL);
 const cols=Array.from({length:L},()=>[]);F.order.forEach(x=>cols[lay[x]].push(x));
 const pos=new Float64Array(K);cols.forEach(c=>c.forEach((x,r)=>pos[x]=r));
 for(let sw=0;sw<6;sw++){const fwd=sw%2===0;(fwd?cols:cols.slice().reverse()).forEach(c=>{c.forEach(x=>{const nb=fwd?F.inA[x]:F.outA[x];let s_=0,t=0;nb.forEach(([y,w])=>{s_+=pos[y]*w;t+=w});if(t)pos[x]=s_/t});c.sort((a,b)=>pos[a]-pos[b]);c.forEach((x,r)=>pos[x]=r)})}
 // columns, wrapped into sub-columns when long
 const maxH=Math.max(260,Math.sqrt(cont.kids.reduce((t,e)=>t+e.w*e.h,0))*1.15);
 let xc=0,H=0;const blocks=[];
 cols.forEach(c=>{let sub=[],h=0;const subs=[];c.forEach(x=>{const e=cont.kids[x];if(sub.length&&h+e.h>maxH){subs.push(sub);sub=[];h=0}sub.push(x);h+=e.h+GAPY});if(sub.length)subs.push(sub);
  subs.forEach(sb=>{const w=Math.max(...sb.map(x=>cont.kids[x].w));let y=0;sb.forEach(x=>{const e=cont.kids[x];e.x=xc+(w-e.w)/2;e.y=y;y+=e.h+GAPY});blocks.push([sb,y-GAPY]);H=Math.max(H,y-GAPY);xc+=w+GAPY*2});xc+=GAPX-GAPY*2});
 blocks.forEach(([sb,h])=>{const off=(H-h)/2;sb.forEach(x=>cont.kids[x].y+=off)});
 cont.F=F;cont.lay=lay;return {w:Math.max(xc-GAPX,120),h:Math.max(H,30)}}
function drawFold(){$('#hint').innerHTML=`<b>All variables, unfold step by step</b>, drawn like a flowchart. Each box is a group of variables that are densely linked inside and loosely linked to the rest. <b>Click a box to unfold it</b>: it becomes a frame with its sub-groups (or variables) inside, laid out the same way; click a frame's title to fold it back. Arrows run mostly <b>left to right</b>, from drivers to results; dashed arrows run back. Arrow width = number of links; only each box's strongest links (and only strong backward links) are drawn unless you tick “all links”. Hover to highlight. “Copy as Mermaid” gives this view as Mermaid code.`;
 const root={kids:elemsOf(TOP)};const sz=layout(root);
 // absolute positions, and which shown box each variable belongs to
 const boxes=[],frames=[],disp=new Map();
 (function place(c,ox,oy,depth){c.kids.forEach(e=>{e.ax=ox+e.x;e.ay=oy+e.y;if(e.k==='c'){frames.push([e,depth]);place(e,e.ax+PAD,e.ay+PAD+TITLE,depth+1)}else{e.id=boxes.length;boxes.push(e);membersOf(e).forEach(i=>disp.set(i,e.id))}})})(root,20,40,0);
 const K=boxes.length,Wm=new Map();for(let v=0;v<N;v++){const dv=disp.get(v);if(dv==null)continue;for(const u of D.par[v]){const du=disp.get(u);if(du==null||du===dv)continue;const k=du*K+dv;Wm.set(k,(Wm.get(k)||0)+1)}}
 const E=[...Wm.entries()].map(([k,w])=>[Math.floor(k/K),k%K,w]);
 const all=$('#fall')&&$('#fall').checked,keep=new Set();
 const wmax=Math.max(1,...E.map(e=>e[2]));const cx=e=>e.ax+e.w/2,cy=e=>e.ay+e.h/2,isBack=k=>cx(boxes[E[k][1]])<cx(boxes[E[k][0]])-4;
 if(all||E.length<=60)E.forEach((e,k)=>keep.add(k));else{const by=Array.from({length:K},()=>[]);E.forEach((e,k)=>{if(isBack(k)&&e[2]<0.5*wmax)return;by[e[0]].push(k);by[e[1]].push(k)});by.forEach(ks=>ks.sort((a,b)=>E[b][2]-E[a][2]).slice(0,3).forEach(k=>keep.add(k)))}
 let ed='';[...keep].sort((a,b)=>E[a][2]-E[b][2]).forEach(k=>{const [a,b,w]=E[k],A=boxes[a],B=boxes[b],back=cx(B)<cx(A)-4;let d;
  if(!back){const x1=A.ax+A.w,y1=cy(A),x2=B.ax,y2=cy(B),m=Math.max(18,(x2-x1)/2);d=`M${x1},${y1} C${x1+m},${y1} ${x2-m},${y2} ${x2},${y2}`}
  else{const x1=A.ax+A.w/2,y1=A.ay+A.h,x2=B.ax+B.w/2,y2=B.ay+B.h,dy=Math.max(y1,y2)+28+Math.abs(x1-x2)*0.05;d=`M${x1},${y1} C${x1},${dy} ${x2},${dy} ${x2},${y2}`}
  ed+=`<path d="${d}" fill="none" stroke="${back?'var(--b)':'var(--ln)'}" stroke-opacity="${back?.55:.7}" stroke-width="${(1.1+3.6*Math.sqrt(w/wmax)).toFixed(1)}"${back?' stroke-dasharray="6 5"':''} marker-end="url(#${back?'mkb':'mk'})" data-e="${a},${b}"><title>${esc(blab(A))} → ${esc(blab(B))}: ${w} link${w>1?'s':''}</title></path>`});
 function blab(e){return e.k==='v'?D.nodes[e.i]:e.g.lab}
 const avgGap=g=>{if(!HAS)return null;let s_=0,c=0;for(const i of g.m){const v=gapOf(i);if(v!=null&&isFinite(v)){s_+=v;c++}}return c?s_/c:null};
 const cut=(t,m)=>t.length>m?t.slice(0,m-1)+'…':t;
 let fr='';frames.sort((a,b)=>a[1]-b[1]).forEach(([e,dp])=>{const g=e.g;fr+=`<g class="ff" data-f="${g.__f}"><rect x="${e.ax}" y="${e.ay}" width="${e.w}" height="${e.h}" rx="8" fill="var(--cf)" fill-opacity="${Math.min(.9,.55+dp*.12)}" stroke="var(--cs)" stroke-width="1.2"/><g class="ft" style="cursor:pointer"><title>${esc(g.lab)}: ${g.m.length.toLocaleString()} variables. Click to fold.</title><rect x="${e.ax}" y="${e.ay}" width="${e.w}" height="${TITLE}" rx="8" fill="transparent"/>${(()=>{const k=g.lab.lastIndexOf(' · '),t=k>0?g.lab.slice(k+3)+' · '+g.lab.slice(0,k):g.lab;return `<text x="${e.ax+10}" y="${e.ay+18}" font-size="12" font-weight="600" fill="var(--fg)">▾ ${esc(cut(t,Math.floor(e.w/7.2)-12))} <tspan font-weight="400" fill="var(--mut)">· ${g.m.length.toLocaleString()} vars</tspan></text>`})()}</g></g>`});
 let nd='';boxes.forEach(e=>{if(e.k==='g'){const g=e.g,ch=changed(g),ag=avgGap(g),can=!!(g.ch||g.m.length>1);
   nd+=`<g class="fg" data-f="${g.__f}" data-x="${e.id}" style="cursor:${can?'pointer':'default'}"><title>${esc(g.lab)}\n${g.m.length.toLocaleString()} variables${HAS?`, ${ch.toLocaleString()} changed, average gap ${pct(ag)}`:''}\n${g.k.length} key variable${g.k.length===1?'':'s'}: ${g.k.slice(0,8).map(i=>D.nodes[i]).join(', ')}${g.k.length>8?', …':''}${can?'\nClick to unfold.':''}</title><rect x="${e.ax}" y="${e.ay}" width="${e.w}" height="${e.h}" rx="6" fill="var(--nf)" stroke="var(--ns)" stroke-width="1.3"/>${HAS&&ag!=null?`<rect x="${e.ax}" y="${e.ay}" width="6" height="${e.h}" rx="3" fill="${ag<0?'var(--neg)':ag>0?'var(--pos)':'var(--mut)'}" opacity="${Math.min(1,.35+Math.abs(ag)/4)}"/>`:''}${(()=>{const k=g.lab.lastIndexOf(' · '),con=k>0?g.lab.slice(0,k):g.lab,geo=k>0?g.lab.slice(k+3):'';const l1=geo||con,l2=geo?con:'';
    return `<text x="${e.ax+e.w/2}" y="${e.ay+18}" text-anchor="middle" font-size="12" font-weight="600" fill="var(--fg)">${esc(cut(l1,28))}</text>${l2?`<text x="${e.ax+e.w/2}" y="${e.ay+33}" text-anchor="middle" font-size="10.5" fill="var(--fg)" opacity=".8">${esc(cut(l2,33))}</text>`:''}<text x="${e.ax+e.w/2}" y="${e.ay+(l2?49:38)}" text-anchor="middle" font-size="10.5" fill="var(--mut)">${g.m.length.toLocaleString()} vars${HAS&&ag!=null?' · '+pct(ag):''}${can?' · ▸ unfold':''}</text>`})()}</g>`}
  else{const i=e.i,nm=D.nodes[i],gv=gapOf(i),hl=HL===i,pg=PATH.get(i),par=pg?pg[pg.length-1]:null,key=par&&par.k.includes(i);
   nd+=`<g class="fv" data-i="${i}" data-x="${e.id}" style="cursor:pointer"><title>${esc(nm)}${descOf(i)?' — '+esc(descOf(i)):''}${HAS?(gv!=null?`: gap ${pct(gv)}`:': unchanged'):''}${key?', key variable of its group':''}. Click for its equation and links.</title><rect x="${e.ax}" y="${e.ay}" width="${e.w}" height="${e.h}" rx="${e.h/2}" fill="var(--bg)" stroke="${hl?'var(--b)':'var(--ns)'}" stroke-width="${hl?3:1.1}"/>${HAS&&gv!=null?`<circle cx="${e.ax+12}" cy="${e.ay+e.h/2}" r="4.5" fill="${gv<0?'var(--neg)':gv>0?'var(--pos)':'var(--mut)'}"/>`:''}<text x="${e.ax+22}" y="${e.ay+19}" font-size="11" fill="var(--fg)">${key?'🔑 ':''}${esc(cut(nm,22))}${HAS&&gv!=null?` <tspan fill="var(--mut)" font-size="10">${pct(gv)}</tspan>`:''}</text></g>`}});
 const Wd=sz.w+60,Hd=sz.h+80;VS.full=[0,0,Wd,Hd];VS.v=VS.full.slice();$('#map').style.overflowX='auto';
 $('#map').innerHTML=`<svg id="sv" viewBox="${VS.v.join(' ')}" style="width:${Math.round(Wd)}px;height:${Math.round(Hd)}px;border:1px solid var(--line);border-radius:10px;cursor:grab;touch-action:none"><defs><marker id="mk" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="10" markerHeight="10" markerUnits="userSpaceOnUse" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--ln)"/></marker><marker id="mkb" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="10" markerHeight="10" markerUnits="userSpaceOnUse" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--b)"/></marker></defs><text x="20" y="24" font-size="11" fill="var(--mut)">drivers →</text><text x="${Wd-20}" y="24" text-anchor="end" font-size="11" fill="var(--mut)">→ results</text>${fr}<g>${ed}</g>${nd}</svg>`;
 const sv=$('#sv');panZoom(sv);
 sv.addEventListener('mouseover',e=>{const g=e.target.closest('.fg,.fv');if(!g)return;const x=g.dataset.x,nb=new Set([x]);sv.querySelectorAll('path[data-e]').forEach(p=>{const [a,b]=p.dataset.e.split(',');const on=a===x||b===x;if(on){nb.add(a);nb.add(b)}p.style.opacity=on?1:.1});sv.querySelectorAll('.fg,.fv').forEach(q=>q.style.opacity=nb.has(q.dataset.x)?1:.3)});
 sv.addEventListener('mouseout',e=>{if(e.relatedTarget&&e.relatedTarget.closest&&e.relatedTarget.closest('.fg,.fv'))return;sv.querySelectorAll('path[data-e]').forEach(p=>p.style.opacity='');sv.querySelectorAll('.fg,.fv').forEach(q=>q.style.opacity='')});
 sv.addEventListener('click',e=>{const t=e.target.closest('.ft');if(t){foldOf(BYF.get(t.closest('.ff').dataset.f));draw();return}const b=e.target.closest('.fg');if(b){const g=BYF.get(b.dataset.f);if(g&&(g.ch||g.m.length>1)){FOLD.add(g);draw()}return}const v=e.target.closest('.fv');if(v)showVar(+v.dataset.i)});
 const nv=boxes.filter(b=>b.k==='v').length,tot=E.reduce((t,e)=>t+e[2],0),back=E.filter(([a,b])=>cx(boxes[b])<cx(boxes[a])-4).reduce((t,e)=>t+e[2],0);
 $('#crumb').innerHTML='';
 $('#note').innerHTML=`${K.toLocaleString()} boxes (${(K-nv).toLocaleString()} groups, ${nv.toLocaleString()} variables) in ${frames.length} unfolded frame${frames.length===1?'':'s'}, covering all ${N.toLocaleString()} variables. ${tot?`${Math.round(100*(tot-back)/tot)}% of the links run left to right.`:''} ${keep.size<E.length?`Showing the ${keep.size} strongest of ${E.length} connections.`:''} ${HAS?'<span class="hint">Bar or dot colour: gap B − A (green up, red down).</span>':''}${K>250?' That is a lot of boxes; fold some frames back to keep it readable.':''}`;
 $('#chain').innerHTML='';FV={root,boxes,E,keep}}
let FV=null;
function toMermaid(){if(!FV)return'';const id=s_=>'n'+s_,q=t=>'"'+String(t).replace(/"/g,"'").replace(/[\[\]{}|<>]/g,' ')+'"';const out=['flowchart LR'];let nid=0;const ids=new Map();
 (function walk(c,ind){c.kids.forEach(e=>{if(e.k==='c'){out.push(`${ind}subgraph s${nid++}[${q(e.g.lab+' · '+e.g.m.length+' vars')}]`);out.push(`${ind}  direction LR`);walk(e,ind+'  ');out.push(`${ind}end`)}else{const k='n'+(nid++);ids.set(e.id,k);out.push(`${ind}${k}[${q(e.k==='v'?D.nodes[e.i]:e.g.lab+' · '+e.g.m.length+' vars')}]`)}})})(FV.root,'  ');
 [...FV.keep].forEach(k=>{const [a,b,w]=FV.E[k];out.push(`  ${ids.get(a)} -->|${w}| ${ids.get(b)}`)});return out.join('\n')}
function foldOf(g){FOLD.delete(g);(g.ch||[]).forEach(foldOf)}
$('#crumb').addEventListener('click',e=>{const c=e.target.closest('[data-ff]');if(c){foldOf(BYF.get(c.dataset.ff));draw()}const r=e.target.closest('[data-rp]');if(r){e.preventDefault();RP=RP.slice(0,+r.dataset.rp);draw()}});
$('#ffold').onclick=()=>{FOLD.clear();HL=null;draw()};
$('#fall').onchange=()=>draw();
$('#fmer').onclick=()=>{const t=toMermaid();const done=()=>{$('#findn').textContent='Mermaid code of this view copied ('+t.split('\n').length+' lines).'};
 const show=()=>{$('#panel').innerHTML=`<span class="x" id="px">×</span><h2 style="margin-top:0">Mermaid code of this view</h2><p class="small">Paste into any Mermaid editor (GitHub, Confluence, Notion, mermaid.live).</p><textarea style="width:100%;height:70vh;font:12px monospace" id="mtx"></textarea>`;$('#mtx').value=t;$('#panel').style.display='block';$('#px').onclick=()=>$('#panel').style.display='none';$('#mtx').select()};
 if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(t).then(()=>{done();show()},show);else show()};
$('#fone').onclick=()=>{const front=[];const w=g=>{if(FOLD.has(g))(g.ch||[]).forEach(w);else if((g.ch||g.m.length>1)&&g.lab!=='(not linked)')front.push(g)};TOP.forEach(w);front.forEach(g=>FOLD.add(g));draw()};
// ---------- main chain, zoomable (search path count)
let RP=[],WTS=null;
function weights(){if(WTS)return WTS;const C=D.cid,nc=Math.max(0,...C)+1,inC=Array.from({length:nc},()=>new Set()),outC=Array.from({length:nc},()=>new Set());
 for(let v=0;v<N;v++)for(const u of D.par[v])if(C[u]!==C[v]){inC[C[v]].add(C[u]);outC[C[u]].add(C[v])}
 const lse=a=>{if(!a.length)return 0;let m=-Infinity;for(const x of a)if(x>m)m=x;let t=0;for(const x of a)t+=Math.exp(x-m);return m+Math.log(t)};
 const lm=new Float64Array(nc),lp=new Float64Array(nc);for(let c=0;c<nc;c++){const a=[];inC[c].forEach(q=>a.push(lm[q]));lm[c]=lse(a)}for(let c=nc-1;c>=0;c--){const a=[];outC[c].forEach(q=>a.push(lp[q]));lp[c]=lse(a)}
 let mx=-Infinity;const raw=[];for(let v=0;v<N;v++)for(const u of D.par[v])if(u!==v){const x=lm[C[u]]+lp[C[v]];raw.push([u,v,x]);if(x>mx)mx=x}
 const out=new Map(),inn=new Map();raw.forEach(([u,v,x])=>{const w=Math.exp(x-mx);if(!out.has(u))out.set(u,[]);out.get(u).push([v,w]);if(!inn.has(v))inn.set(v,[]);inn.get(v).push([u,w])});return WTS={out,inn,n:raw.length}}
function routeIn(cont,prev,next){const W=weights(),kids=cont.ch?cont.ch:cont.m.map(i=>({m:[i],v:i,lab:D.nodes[i],k:[]})),own=new Map();kids.forEach((c,x)=>c.m.forEach(i=>own.set(i,x)));
 const K=kids.length,agg=Array.from({length:K},()=>new Map()),ent=new Float64Array(K),ext=new Float64Array(K),inO=new Float64Array(K),outO=new Float64Array(K),inI=new Float64Array(K),outI=new Float64Array(K);
 own.forEach((x,u)=>{for(const [v,w] of W.out.get(u)||[]){const y=own.get(v);if(y!=null){if(y!==x){agg[x].set(y,(agg[x].get(y)||0)+w);outI[x]+=w;inI[y]+=w}}else{outO[x]+=w;if(next&&next.has(v))ext[x]+=w}}
  for(const [q,w] of W.inn.get(u)||[]){if(own.has(q))continue;inO[x]+=w;if(prev&&prev.has(q))ent[x]+=w}});
 if(K===1)return {kids,chain:[0],agg};const all=[...Array(K).keys()];
 let st=prev?all.filter(x=>ent[x]>0):[];if(!st.length){st=all.filter(x=>inO[x]>0);if(!st.length)st=all.filter(x=>inI[x]===0)}if(!st.length)st=all;
 let en=next?all.filter(x=>ext[x]>0):[];if(!en.length){en=all.filter(x=>outO[x]>0);if(!en.length)en=all.filter(x=>outI[x]===0)}if(!en.length)en=all;
 const grow=(ch,seen,thr,fwd)=>{for(;;){const e=fwd?ch[ch.length-1]:ch[0];let b=null;if(fwd)agg[e].forEach((w,y)=>{if(!seen.has(y)&&w>=thr&&(!b||w>b[1]))b=[y,w]});else agg.forEach((m,x)=>{const w=m.get(e);if(w&&!seen.has(x)&&w>=thr&&(!b||w>b[1]))b=[x,w]});if(!b)return ch;fwd?ch.push(b[0]):ch.unshift(b[0]);seen.add(b[0])}};
 if(!prev&&!next){let best=null;agg.forEach((m,x)=>m.forEach((w,y)=>{if(!best||w>best[2])best=[x,y,w]}));if(!best)return {kids,chain:[0],agg};const ch=[best[0],best[1]],seen=new Set(ch);grow(ch,seen,best[2]*0.05,false);grow(ch,seen,best[2]*0.05,true);return {kids,chain:ch,agg}}
 if(prev&&!next){const s0=st.reduce((a,x)=>ent[x]>ent[a]?x:a,st[0]);const mx=Math.max(ent[s0],...agg[s0].values(),0);return {kids,chain:grow([s0],new Set([s0]),mx*0.05,true),agg}}
 if(next&&!prev){const e0=en.reduce((a,x)=>ext[x]>ext[a]?x:a,en[0]);let mx=ext[e0];agg.forEach(m=>{const w=m.get(e0);if(w>mx)mx=w});return {kids,chain:grow([e0],new Set([e0]),mx*0.05,false),agg}}
 const wid=new Float64Array(K).fill(-1),pred=new Int32Array(K).fill(-1),done=new Uint8Array(K);st.forEach(x=>wid[x]=ent[x]);
 for(let it=0;it<K;it++){let u=-1;for(let x=0;x<K;x++)if(!done[x]&&wid[x]>=0&&(u<0||wid[x]>wid[u]))u=x;if(u<0)break;done[u]=1;agg[u].forEach((w,y)=>{const c=Math.min(wid[u],w);if(c>wid[y]){wid[y]=c;pred[y]=u}})}
 let t=-1,tv=-1;en.forEach(x=>{if(wid[x]<0)return;const v=Math.min(wid[x],ext[x]);if(v>tv){tv=v;t=x}});if(t<0)t=st[0];
 const ch=[];for(let x=t,g=0;x>=0&&g<K;x=pred[x],g++)ch.unshift(x);return {kids,chain:ch,agg}}
function drawRoute(){$('#hint').innerHTML=`<b>Main chain, zoomable</b>, like directions on a map. At the top, the groups and the main chain through them: it starts at the link that the most input-to-end chains run through (search path count) and extends along the heaviest links. <b>Click a step</b> to see the chain inside it: it starts where links from the step before enter and ends where links leave for the next step (the widest path: the route whose weakest link is strongest), down to single variables. The path above goes back up.`;
 const root=(TOP.length===1&&!TOP[0].ch)?Object.assign({},TOP[0],{lab:'Whole model'}):{lab:'Whole model',ch:TOP,m:[],k:[]};
 const path=[{node:root},...RP],cur=path[path.length-1];
 $('#crumb').innerHTML=path.map((p,x)=>x===path.length-1?`<b>${esc(p.lab||p.node.lab)}</b>`:`<a href="#" data-rp="${x}">${esc(p.lab||p.node.lab)}</a>`).join(' › ');
 if(!weights().n){$('#note').textContent='No links.';$('#map').innerHTML='';return}
 const R=routeIn(cur.node,cur.prev,cur.next),kids=R.kids,ch=R.chain,nO=kids.length-ch.length;
 $('#note').innerHTML=`${ch.length} step${ch.length===1?'':'s'}${nO>0?` (${nO} other ${cur.node.ch?'group':'variable'}${nO===1?'':'s'} here are off the main chain)`:''}. ${cur.prev?`Starts where links from <b>${esc(cur.prevLab)}</b> enter`:'Starts at the heaviest link'}${cur.next?` and ends where links leave for <b>${esc(cur.nextLab)}</b>`:''}. ${kids.some(c=>c.v==null)?'Click a step (▸) to see the chain inside it.':'Click a variable for its equation and links.'}`;
 const items=[];if(cur.prev)items.push({k:'prev',lab:cur.prevLab});ch.forEach((x,j)=>items.push({k:'step',x,j}));if(cur.next)items.push({k:'next',lab:cur.nextLab});
 const BW=196,BH=46,GX=56,GY=34,PER=Math.max(3,Math.floor(((($('main').clientWidth||1100)-20)+GX)/(BW+GX))),rows=Math.ceil(items.length/PER);
 const P=items.map((it,i)=>{const r=Math.floor(i/PER),c=r%2?PER-1-(i%PER):i%PER;return {x:10+c*(BW+GX),y:12+r*(BH+GY)}});
 const W=10+Math.min(items.length,PER)*(BW+GX)-GX+10,H=12+rows*(BH+GY)-GY+12;
 let ed='';for(let i=1;i<items.length;i++){const a=P[i-1],b=P[i];let d;if(a.y===b.y){const x1=a.x<b.x?a.x+BW:a.x,x2=a.x<b.x?b.x:b.x+BW,y=a.y+BH/2;d=`M${x1},${y} L${x2},${y}`}else{const x=a.x+BW/2;d=`M${x},${a.y+BH} L${x},${b.y}`}
  ed+=`<path d="${d}" fill="none" stroke="var(--a)" stroke-width="2.4" marker-end="url(#mr)"${items[i-1].k!=='step'||items[i].k!=='step'?' stroke-dasharray="5 4"':''}/>`}
 let nd='';items.forEach((it,i)=>{const p=P[i];
  if(it.k!=='step'){nd+=`<g><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="9" fill="var(--bg)" stroke="var(--mut)" stroke-dasharray="5 4"/><text x="${p.x+8}" y="${p.y+17}" font-size="10.5" fill="var(--mut)">${it.k==='prev'?'from the step before':'on to the next step'}</text><text x="${p.x+8}" y="${p.y+34}" font-size="11.5" fill="var(--fg)">${esc(it.lab.length>24?it.lab.slice(0,23)+'…':it.lab)}</text></g>`;return}
  const c=kids[it.x];
  if(c.v!=null){const n=D.nodes[c.v],g=gapOf(c.v);nd+=`<g class="rv" data-i="${c.v}" style="cursor:pointer"><title>${esc(n)}${descOf(c.v)?' — '+esc(descOf(c.v)):''}</title><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="7" fill="var(--bg)"/><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="7" fill="${fillOf(g)}" stroke="var(--fg)" stroke-opacity=".5"/><text x="${p.x+8}" y="${p.y+18}" font-size="10.5" fill="var(--mut)">step ${it.j+1}</text><text x="${p.x+8}" y="${p.y+35}" font-size="12" font-weight="600" fill="var(--fg)">${esc(n.length>23?n.slice(0,22)+'…':n)}${g!=null?` <tspan font-weight="400" font-size="11">${pct(g)}</tspan>`:''}</text></g>`}
  else{const more=!!(c.ch||c.m.length>1);nd+=`<g class="rg" data-j="${it.j}" style="cursor:pointer"><title>${esc(c.lab)}: ${c.m.length} variables. Click to see the chain inside it.</title><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="9" fill="var(--card)" stroke="var(--a)" stroke-width="1.6"/><text x="${p.x+8}" y="${p.y+17}" font-size="10.5" fill="var(--mut)">step ${it.j+1} · ${c.m.length.toLocaleString()} vars${more?' ▸':''}</text><text x="${p.x+8}" y="${p.y+34}" font-size="11.5" font-weight="600" fill="var(--fg)">${esc(c.lab.length>25?c.lab.slice(0,24)+'…':c.lab)}</text></g>`}});
 VS.full=[0,0,W,H];VS.v=VS.full.slice();
 $('#map').innerHTML=`<svg id="sv" viewBox="${VS.v.join(' ')}" style="width:${Math.min(($('main').clientWidth||1100),W)}px;max-width:100%;height:${Math.round(Math.min(700,H*Math.min(1,($('main').clientWidth||1100)/W)))}px;border:1px solid var(--line);border-radius:10px;cursor:grab;touch-action:none"><defs><marker id="mr" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="9" markerHeight="9" markerUnits="userSpaceOnUse" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--a)"/></marker></defs>${ed}${nd}</svg>`;
 const sv=$('#sv');panZoom(sv);
 $('#chain').innerHTML=ch.map(x=>{const c=kids[x];return c.v!=null?vlink(c.v):`<b>${esc(c.lab)}</b>`}).join(' → ');
 sv.addEventListener('click',e=>{const v=e.target.closest('.rv');if(v){showVar(+v.dataset.i);return}const g=e.target.closest('.rg');if(!g)return;const j=+g.dataset.j,c=kids[ch[j]];if(!(c.ch||c.m.length>1))return;
  const bf=j>0?kids[ch[j-1]]:null,af=j<ch.length-1?kids[ch[j+1]]:null,lab=x=>x.v!=null?D.nodes[x.v]:x.lab;
  RP.push({node:c,lab:`step ${j+1}: ${c.lab}`,prev:bf?new Set(bf.m):cur.prev,prevLab:bf?lab(bf):cur.prevLab,next:af?new Set(af.m):cur.next,nextLab:af?lab(af):cur.nextLab});draw()})}
// ---------- around a variable: what feeds it and what it feeds, a few steps each way
let TR=null;
function drawTrace(){$('#hint').innerHTML=`<b>Around a variable.</b> The variable in the middle, the variables that feed it on the left (1, 2, … equations away) and the variables it feeds on the right. Click a box for its equation and links; double-click to put it in the middle. Choose a variable with the search box or from its panel.`;$('#crumb').innerHTML='';$('#chain').innerHTML='';
 if(TR==null){let b=0;for(let i=1;i<N;i++)if(CH[i].length+D.par[i].length>CH[b].length+D.par[b].length)b=i;TR=b}
 const du=+$('#tu').value,dd=+$('#td').value,dist=new Map([[TR,0]]),cap=260;
 const bfs=(nb,lim,sign)=>{let fr=[TR];for(let k=1;k<=lim;k++){const nx=[];for(const u of fr)for(const v of nb[u])if(!dist.has(v)){dist.set(v,sign*k);nx.push(v)}fr=nx;if(dist.size>cap)break}};
 bfs(D.par,du,-1);bfs(CH,dd,1);
 let ids=[...dist.keys()];const cut=ids.length>cap;if(cut)ids=ids.sort((a,b)=>Math.abs(dist.get(a))-Math.abs(dist.get(b))||Math.abs(gapOf(b)||0)-Math.abs(gapOf(a)||0)).slice(0,cap);
 const S=new Set(ids),cols=new Map();ids.forEach(i=>{const c=dist.get(i);if(!cols.has(c))cols.set(c,[]);cols.get(c).push(i)});
 const cks=[...cols.keys()].sort((a,b)=>a-b),BW=176,BH=30,GX=70,SX=12,GY=8,R=Math.max(10,Math.min(30,Math.ceil(Math.sqrt(ids.length)*1.4)));
 cks.forEach(c=>cols.get(c).sort((a,b)=>Math.abs(gapOf(b)||0)-Math.abs(gapOf(a)||0)||(D.nodes[a]<D.nodes[b]?-1:1)));
 const nsub=new Map(),x0=new Map();let xc=10;cks.forEach(c=>{nsub.set(c,Math.max(1,Math.ceil(cols.get(c).length/R)));x0.set(c,xc);xc+=nsub.get(c)*(BW+SX)-SX+GX});
 const P=new Map();let H=0;cks.forEach(c=>{const a=cols.get(c),rk=Math.ceil(a.length/nsub.get(c));a.forEach((i,r)=>{const s=Math.floor(r/rk);P.set(i,{x:x0.get(c)+s*(BW+SX),y:24+(r%rk)*(BH+GY)});H=Math.max(H,24+((r%rk)+1)*(BH+GY))})});
 const W=xc-GX+20;let ed='';const E=[];ids.forEach(v=>D.par[v].forEach(u=>{if(S.has(u))E.push([u,v])}));const EO=Math.max(.12,Math.min(.45,30/Math.max(1,E.length))).toFixed(2);
 E.forEach(([u,v])=>{const A=P.get(u),B=P.get(v),x1=A.x+BW,y1=A.y+BH/2,y2=B.y+BH/2;let d;if(x1<B.x){const m=(x1+B.x)/2;d=`M${x1},${y1} C${m},${y1} ${m},${y2} ${B.x},${y2}`}else{const bx=Math.max(x1,B.x+BW)+24;d=`M${x1},${y1} C${bx},${y1} ${bx},${y2} ${B.x+BW},${y2}`}
  ed+=`<path d="${d}" fill="none" stroke="var(--mut)" stroke-opacity="${EO}" marker-end="url(#mt)" data-e="${u},${v}"/>`});
 const hd=cks.map(c=>`<text x="${x0.get(c)}" y="14" font-size="11" fill="var(--mut)">${c===0?'the variable':c<0?`feeds it, ${-c} step${c===-1?'':'s'} away`:`fed by it, ${c} step${c===1?'':'s'} away`} (${cols.get(c).length})</text>`).join('');
 let nd='';ids.forEach(i=>{const p=P.get(i),n=D.nodes[i],g=gapOf(i),f=i===TR;nd+=`<g class="tv" data-i="${i}" style="cursor:pointer"><title>${esc(n)}${descOf(i)?' — '+esc(descOf(i)):''}${HAS?(g!=null?`: gap ${pct(g)}`:': unchanged'):''}. Click for details, double-click to put it in the middle.</title><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="6" fill="var(--bg)"/><rect x="${p.x}" y="${p.y}" width="${BW}" height="${BH}" rx="6" fill="${fillOf(g)}" stroke="${f?'var(--b)':'var(--line)'}" stroke-width="${f?3:1}"/><text x="${p.x+8}" y="${p.y+19}" font-size="11.5" font-weight="${f?700:400}" fill="var(--fg)">${esc(n.length>21?n.slice(0,20)+'…':n)}${HAS&&g!=null?` <tspan font-size="10.5" fill="var(--mut)">${pct(g)}</tspan>`:''}</text></g>`});
 VS.full=[0,0,W,H+10];VS.v=VS.full.slice();const cw=$('main').clientWidth||1100,w=Math.min(cw,Math.round(W*1.05));
 $('#map').innerHTML=`<svg id="sv" viewBox="${VS.v.join(' ')}" style="width:${w}px;max-width:100%;height:${Math.round(Math.min(760,Math.max(160,w*(H+10)/W)))}px;border:1px solid var(--line);border-radius:10px;cursor:grab;touch-action:none"><defs><marker id="mt" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="8" markerHeight="8" markerUnits="userSpaceOnUse" orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="var(--mut)"/></marker></defs>${hd}${ed}${nd}</svg>`;
 const sv=$('#sv');panZoom(sv);
 sv.addEventListener('mouseover',e=>{const g=e.target.closest('.tv');if(!g)return;const i=g.dataset.i;sv.querySelectorAll('path[data-e]').forEach(p=>{const [u,v]=p.dataset.e.split(',');const on=u===i||v===i;p.setAttribute('stroke',on?'var(--a)':'var(--mut)');p.setAttribute('stroke-opacity',on?'.9':EO)})});
 sv.addEventListener('click',e=>{const g=e.target.closest('.tv');if(g)showVar(+g.dataset.i)});
 sv.addEventListener('dblclick',e=>{const g=e.target.closest('.tv');if(g){TR=+g.dataset.i;draw()}});
 const nu=ids.filter(i=>dist.get(i)<0).length,ndn=ids.filter(i=>dist.get(i)>0).length;
 $('#note').innerHTML=`<b>${esc(D.nodes[TR])}</b>${descOf(TR)?' — '+esc(descOf(TR)):''}: ${nu} variable${nu===1?'':'s'} feeding it within ${du} step${du===1?'':'s'}, ${ndn} fed by it within ${dd}.${cut?` Showing the ${cap} closest.`:''}`}
function draw(){const m=$('#mby').value;$('#mcf').style.display=m==='fold'?'':'none';$('#mct').style.display=m==='trace'?'':'none';m==='fold'?drawFold():m==='trace'?drawTrace():drawRoute()}
$('#tu').onchange=$('#td').onchange=()=>draw();
$('#mby').onchange=()=>{RP=[];draw()};
// ---------- one variable: equation, links, place in the groups
function showVar(i){const n=D.nodes[i],pa=PATH.get(i)||[],g=gapOf(i),ps=D.par[i],cs=CH[i];
 $('#panel').innerHTML=`<span class="x" id="px">×</span><h2 style="margin-top:0">${esc(n)}</h2>
 ${descOf(i)?`<div style="margin:-4px 0 6px">${esc(descOf(i))}</div>`:''}<div class="small">${pa.map(x=>esc(x.lab)).join(' › ')||'(not linked)'}</div>
 ${HAS?`<p>Gap B − A: <b class="${cls(g)}">${g==null?'unchanged':pct(g)}</b> at the end${D.gap[i]?`, largest ${pct(D.gap[i][1])}`:''}</p>`:''}
 <p class="small">Layers: <b>${D.lay[i]}</b> (longest chain of equations downstream)${D.lp[i]>1?` · in a loop of ${D.lp[i]} variables`:''}${D.ex[i]?' · treated as exogenous (--exog)':''}</p>
 ${D.eq[i]?`<pre>${esc(D.eq[i])}</pre>`:`<p class="hint">No equation: an input to the model.</p>`}${D.af&&D.af[i]?`<p class="small">Add-factor series: <b>${D.af[i].map(esc).join(', ')}</b> (left out of the map)</p>`:''}
 <p><b>Uses</b> (${ps.length}): ${ps.slice(0,60).map(vlink).join(', ')||'–'}${ps.length>60?' …':''}</p>
 <p><b>Used by</b> (${cs.length}): ${cs.slice(0,60).map(vlink).join(', ')||'–'}${cs.length>60?' …':''}</p>
 <button class="btn" id="pfind">show it in the map</button> <button class="btn" id="ptrace">around it: what feeds it, what it feeds</button>`;
 $('#panel').style.display='block';$('#px').onclick=()=>$('#panel').style.display='none';$('#pfind').onclick=()=>findVar(i);$('#ptrace').onclick=()=>{TR=i;$('#mby').value='trace';$('#panel').style.display='none';draw()}}
document.addEventListener('click',e=>{const a=e.target.closest('[data-vi]');if(a){e.preventDefault();showVar(+a.dataset.vi)}});
function findVar(i){if($('#mby').value==='trace'){TR=i;draw();return}const pa=PATH.get(i);if(!pa){$('#findn').textContent='not in any group';return}$('#mby').value='fold';RP=[];pa.forEach(g=>{if(g.ch||g.m.length>1)FOLD.add(g)});HL=i;draw();
 $('#findn').textContent=pa.map(g=>g.lab).join(' › ');const b=document.querySelector(`#sv .fv[data-i="${i}"] rect`);if(b){const r=b.getBBox();VS.v=[r.x-450,r.y-250,900+r.width,500];$('#sv').setAttribute('viewBox',VS.v.join(' '));$('#map').scrollIntoView({block:'start'})}}
const NOTFOUND=q=>`No variable matches “${esc(q.trim())}”. It is not in the equations or the data that were read. To see why, run: <code>python model_map.py … --find ${esc(q.trim().toUpperCase())}</code> with the same files.`;
function matches(q){q=q.trim();if(!q)return[];const Q=q.toUpperCase(),ql=q.toLowerCase(),out=[];const ex=NI[Q];if(ex!=null)out.push(ex);
 for(let i=0;i<N&&out.length<40;i++){if(i===ex)continue;if(D.nodes[i].includes(Q))out.push(i)}
 if(D.desc)for(let i=0;i<N&&out.length<40;i++){if(out.includes(i))continue;if((D.desc[i]||'').toLowerCase().includes(ql))out.push(i)}return out}
function pick(i){$('#fl').style.display='none';findVar(i);showVar(i)}
$('#find').addEventListener('input',()=>{const ms=matches($('#find').value);const fl=$('#fl');if(!ms.length){fl.style.display=$('#find').value.trim()?'block':'none';fl.innerHTML='<div class="hint" style="padding:8px">'+NOTFOUND($('#find').value)+'</div>';return}
 fl.innerHTML=ms.map(i=>`<div data-pick="${i}" style="padding:5px 9px;cursor:pointer;border-bottom:1px solid var(--line)"><b>${esc(D.nodes[i])}</b> ${descOf(i)?`<span class="hint">${esc(descOf(i).slice(0,80))}</span>`:''}${HAS&&gapOf(i)!=null?` <span class="${cls(gapOf(i))} small">${pct(gapOf(i))}</span>`:''}</div>`).join('')+(ms.length>=40?'<div class="hint" style="padding:6px 9px">first 40 matches; type more to narrow</div>':'');fl.style.display='block'});
$('#find').addEventListener('keydown',e=>{if(e.key==='Enter'){const ms=matches($('#find').value);if(ms.length)pick(ms[0]);else $('#findn').innerHTML=NOTFOUND($('#find').value)}if(e.key==='Escape')$('#fl').style.display='none'});
$('#fl').addEventListener('click',e=>{const d=e.target.closest('[data-pick]');if(d)pick(+d.dataset.pick)});
document.addEventListener('click',e=>{if(!e.target.closest('#fl')&&e.target.id!=='find')$('#fl').style.display='none'});
// ---------- table of all variables, folded by group
(function(){const tb=$('#ttb'),ST=new Map(),stat=g=>{if(ST.has(g))return ST.get(g);let ch=0,lay=0;for(const i of g.m){if(HAS&&D.gap[i])ch++;if(D.lay[i]>lay)lay=D.lay[i]}const o={ch,lay};ST.set(g,o);return o};
 const rowG=(g,d,pid)=>{const s=stat(g);return `<tr data-id="${g.__f}" data-p="${pid}" data-d="${d}" class="tg"><td style="padding-left:${6+d*18}px;white-space:nowrap"><span class="tg_" style="cursor:pointer;display:inline-block;width:14px">▸</span><b>${esc(g.lab)}</b></td><td class="n">${g.m.length.toLocaleString()}</td>${HAS?`<td class="n">${s.ch.toLocaleString()}</td><td></td>`:''}<td class="n">${g.k.length||'–'}</td><td class="n">${s.lay}</td></tr>`};
 const rowV=(i,d,pid,ks)=>{const g=gapOf(i);return `<tr data-p="${pid}" data-d="${d}"><td style="padding-left:${20+d*18}px">${ks.has(i)?'🔑 ':''}${vlink(i)}${descOf(i)?` <span class="hint">${esc(descOf(i).slice(0,70))}</span>`:''}</td><td></td>${HAS?`<td class="n">${D.gap[i]?'yes':''}</td><td class="n ${cls(g)}">${g==null?'unchanged':pct(g)}</td>`:''}<td></td><td class="n">${D.lay[i]}</td></tr>`};
 const open=id=>{const g=BYF.get(id),tr=tb.querySelector(`tr[data-id="${id}"]`);if(!g||!tr||tr.dataset.open)return;tr.dataset.open='1';tr.querySelector('.tg_').textContent='▾';const d=+tr.dataset.d+1;
  tr.insertAdjacentHTML('afterend',g.ch?g.ch.map(c=>rowG(c,d,id)).join(''):g.m.slice().sort((a,b)=>Math.abs(gapOf(b)||0)-Math.abs(gapOf(a)||0)||(D.nodes[a]<D.nodes[b]?-1:1)).map(i=>rowV(i,d,id,new Set(g.k))).join(''))};
 const close=id=>{const tr=tb.querySelector(`tr[data-id="${id}"]`);if(!tr)return;delete tr.dataset.open;tr.querySelector('.tg_').textContent='▸';const kill=[id];for(let x=0;x<kill.length;x++)tb.querySelectorAll(`tr[data-p="${kill[x]}"]`).forEach(r=>{if(r.dataset.id)kill.push(r.dataset.id);r.remove()})};
 const reset=()=>{tb.innerHTML=TOP.map(g=>rowG(g,0,'root')).join('')};reset();
 tb.addEventListener('click',e=>{const t=e.target.closest('.tg_');if(!t)return;const tr=t.closest('tr');tr.dataset.open?close(tr.dataset.id):open(tr.dataset.id)});
 $('#tx0').onclick=reset;$('#tx1').onclick=()=>{const rows=[...tb.querySelectorAll('tr.tg')].filter(r=>!r.dataset.open);if(!rows.length)return;const md=Math.min(...rows.map(r=>+r.dataset.d));rows.filter(r=>+r.dataset.d===md).forEach(r=>open(r.dataset.id))}})();
draw();
</script></body></html>
'''


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--eq', action='append', help='equation file(s) or folder (repeatable)')
    p.add_argument('--links', action='append', help='instead of equations: csv/xlsx of links with from/to columns')
    p.add_argument('--desc', help='csv/xlsx of variable descriptions (name, description)')
    p.add_argument('--json', help='also write groups, layers, links and gaps as JSON')
    p.add_argument('--csv', help='also write every variable with its group path as CSV')
    p.add_argument('--eq-extra', action='append', help='more equations that override --eq (e.g. a second model version)')
    p.add_argument('--a', action='append', help='optional: scenario A data (csv/xlsx), to colour boxes by the gap')
    p.add_argument('--b', action='append', help='optional: scenario B data (csv/xlsx)')
    p.add_argument('--exog', action='append', help='variables to treat as exogenous (names, or a file of names)')
    p.add_argument('--map-boxes', type=int, default=40, help='most boxes in one view; groups get more levels to fit')
    p.add_argument('--af-suffix', default='_A',
                   help="suffix of add-factor series (FGDP.IUSA_A or FGDP_A.IUSA): left out of the map and shown with "
                        "their variable when no equation uses them; '' to keep them")
    p.add_argument('--map-geo', action='store_true', help='split by geography first (default: by equation links)')
    p.add_argument('--geo-sep', default='auto', help="separator before geography codes in names: '.', '_' or auto")
    p.add_argument('--no-geo-names', action='store_true', help='do not use geographies to name groups')
    p.add_argument('--groups', action='store_true', help='also print the group tree in the terminal')
    p.add_argument('--find', action='append', metavar='NAME',
                   help='explain where a variable is (equations, data file/sheet/row) and why it is or is not in the map')
    p.add_argument('--unread', action='store_true', help='list equation lines that could not be read')
    p.add_argument('--depth', type=int, default=2, help='levels of the group tree to print')
    p.add_argument('--title', help='title of the page')
    p.add_argument('-o', '--out', help='HTML page to write (default model_map.html)')
    p.add_argument('-q', '--quiet', action='store_true')
    args = p.parse_args(argv)
    if bool(args.a) != bool(args.b):
        p.error('give both --a and --b, or neither')
    if not args.eq and not args.links:
        p.error('give --eq (equation files) or --links (a table of links)')
    if args.find:
        for n in args.find:
            find_variable(n, args.eq or (), args.a or (), args.b or (), args.links or ())
            print()
        return 0
    if args.unread and args.eq:
        _, _, _, fails = load_equations(args.eq)
        print('%d equation lines could not be read%s' % (len(fails), ':' if fails else '.'))
        for f in fails:
            print('  ' + f)
        return 0
    build(args)
    return 0




if __name__ == '__main__':
    sys.exit(main())
