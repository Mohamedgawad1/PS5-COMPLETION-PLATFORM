"""
split_data.py - move the datasets out of index.html into d_<name>.js files.

Reads the multi-declaration const statements that used to sit at the top of the
app script (SUBS/PUNCH/ITR/RFC/ITRT/PUNT/MILES on one 6.3 MB line, then
NOTES/PCOL/OVR, then RFCK/DPRSUM, then FG/FZ/IO/TRM/CS/CBL/WIRE), writes one
file per group, and verifies every dataset survives pack -> unpack unchanged.

Run this before patch_index.py strips the data lines out of index.html.
Re-runnable: it always reads the untouched data out of index.html.
"""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ps5data

P = os.path.dirname(os.path.abspath(__file__))
IDX = os.path.join(P, 'index.html')

# ------------------------------------------------------------- JS literal parser
WS = ' \t\r\n'


class JsLit(object):
    """Recursive-descent reader for the JS literal subset used in the data.

    json.loads cannot be used: the data has unquoted object keys, single-quoted
    strings and bare true/false/null.
    """

    def __init__(self, s, i=0):
        self.s = s
        self.i = i

    def ws(self):
        s, n = self.s, len(self.s)
        while self.i < n:
            c = s[self.i]
            if c in WS:
                self.i += 1
            elif c == '/' and s.startswith('//', self.i):
                j = s.find('\n', self.i)
                self.i = n if j < 0 else j + 1
            elif c == '/' and s.startswith('/*', self.i):
                j = s.find('*/', self.i)
                self.i = n if j < 0 else j + 2
            else:
                return

    def value(self):
        self.ws()
        c = self.s[self.i]
        if c == '[':
            return self.arr()
        if c == '{':
            return self.obj()
        if c in '"\'':
            return self.string()
        if self.s.startswith('true', self.i):
            self.i += 4
            return True
        if self.s.startswith('false', self.i):
            self.i += 5
            return False
        if self.s.startswith('null', self.i):
            self.i += 4
            return None
        if self.s.startswith('undefined', self.i):
            self.i += 9
            return None
        return self.number()

    def peek(self):
        if self.i >= len(self.s):
            raise SyntaxError('unexpected end of input at %d' % self.i)
        return self.s[self.i]

    def arr(self):
        self.i += 1
        out = []
        while True:
            self.ws()
            c = self.peek()
            if c == ']':
                self.i += 1
                return out
            if c == ',':
                self.i += 1
                continue
            out.append(self.value())
            self.ws()
            if self.peek() == ',':
                self.i += 1

    def obj(self):
        self.i += 1
        out = {}
        while True:
            self.ws()
            c = self.peek()
            if c == '}':
                self.i += 1
                return out
            if c == ',':
                self.i += 1
                continue
            k = self.string() if c in '"\'' else self.barekey()
            self.ws()
            if self.peek() != ':':
                raise SyntaxError('expected : at %d near %r' % (self.i, self.s[self.i:self.i + 40]))
            self.i += 1
            out[k] = self.value()
            self.ws()
            if self.peek() == ',':
                self.i += 1

    def barekey(self):
        m = re.compile(r'[A-Za-z_$][A-Za-z0-9_$]*').match(self.s, self.i)
        if not m:
            raise SyntaxError('bad key at %d near %r' % (self.i, self.s[self.i:self.i + 40]))
        self.i = m.end()
        return m.group(0)

    def string(self):
        q = self.s[self.i]
        self.i += 1
        buf = []
        while True:
            c = self.s[self.i]
            if c == '\\':
                nxt = self.s[self.i + 1]
                buf.append({'n': '\n', 't': '\t', 'r': '\r', 'b': '\b', 'f': '\f',
                            '0': '\0', 'v': '\v'}.get(nxt, nxt))
                self.i += 2
            elif c == q:
                self.i += 1
                return ''.join(buf)
            else:
                buf.append(c)
                self.i += 1

    def number(self):
        m = re.compile(r'-?(?:\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+)').match(self.s, self.i)
        if not m:
            raise SyntaxError('bad value at %d near %r' % (self.i, self.s[self.i:self.i + 40]))
        self.i = m.end()
        t = m.group(0)
        return int(t) if re.fullmatch(r'-?\d+', t) else float(t)


def grab(src, start):
    """Parse one JS value at src[start:], return (value, end_index)."""
    p = JsLit(src, start)
    return p.value(), p.i


def extract(src):
    """Pull every dataset out of index.html. Returns {NAME: value}."""
    data = {}
    # Each entry is a `const`/`let` statement; names are comma separated
    # (line 338 declares seven datasets in one 6.3 MB statement).
    for stmt in ('SUBS', 'NOTES', 'RFCK', 'FG', 'FZ', 'IO', 'TRM', 'CS', 'CBL', 'WIRE'):
        m = re.search(r'(?:const|let|var)\s+' + stmt + r'\s*=\s*', src)
        if not m:
            raise SystemExit('cannot find declaration of %s' % stmt)
        i = m.end()
        cur = stmt
        while True:
            v, i = grab(src, i)
            data[cur] = v
            j = i
            while j < len(src) and src[j] in WS:
                j += 1
            # `,NAME=` continues the same statement; `;const NAME=` starts the
            # next one (RFCK is followed by `;const DPRSUM=...`).
            if src[j] == ',':
                nm = JsLit(src, j + 1).barekey()
                k = j + 1 + len(nm)
                while src[k] in WS:
                    k += 1
                if src[k] != '=':
                    break
                cur = nm
                i = k + 1
                continue
            km = re.compile(r'(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*').match(src, j + 1 if src[j] == ';' else j)
            if not km:
                break
            cur = km.group(1)
            i = km.end()
    return data


# ------------------------------------------------------------------- done counts
def done_i(s):
    s = str(s or '').strip()
    if re.search(r'to\s*be', s, re.I):
        return False
    return bool(re.match(r'comp', s, re.I) or re.fullmatch(r'closed', s, re.I))


def main():
    src = io.open(IDX, encoding='utf-8').read()
    data = extract(src)
    print('extracted: ' + ', '.join('%s=%d' % (k, len(v) if isinstance(v, (list, dict)) else 1)
                                    for k, v in data.items()))

    data['KCNT'] = {
        'itrT': len(data['ITR']),
        'itrCl': sum(1 for r in data['ITR'] if done_i(r[5])),
        'puT': len(data['PUNCH']),
        # must stay identical to renderDash() in index.html
        'puCl': sum(1 for r in data['PUNCH'] if str(r[5]).strip() == 'Closed'),
    }
    print('KCNT: %s' % data['KCNT'])

    import time
    stamp = time.strftime('%Y-%m-%d %H:%M')
    total = 0
    for key, names in ps5data.FILES.items():
        fname = ps5data.PREFIX + key
        groups = []
        for n in names:
            if n not in data:
                raise SystemExit('missing dataset %s' % n)
            groups.append((n, data[n]))
        banner = ('PS5 platform data - %s - rebuilt %s. '
                  'Generated by split_data.py, do not edit by hand.' % (key, stamp))
        text = ps5data.build_js(groups, banner)
        out = os.path.join(P, fname + '.js')
        with io.open(out, 'w', encoding='utf-8', newline='\n') as fh:
            fh.write(text)
        total += len(text)
        # verify: parse the emitted payload straight back
        head, _, rest = text.partition('\n(function()')
        body = rest.rsplit('})();', 1)[0]
        for line in body.split('\n'):
            line = line.strip()
            if not line.startswith('window.'):
                continue
            nm, _, expr = line[len('window.'):].partition('=')
            nm = nm.rstrip(';')
            payload = expr.rstrip(';')
            packed = payload.startswith('PS5UNPACK(')
            if packed:
                payload = payload[len('PS5UNPACK('):-1]
            got = ps5data.roundtrip(payload)
            if got != data[nm]:
                raise SystemExit('ROUND-TRIP MISMATCH in %s.js for %s' % (fname, nm))
        print('%-12s %9d bytes  %s' % (fname + '.js', len(text),
                                       ' '.join('%s:%d' % (n, len(data[n])) for n in names)))

    print('-' * 46)
    print('%-12s %9d bytes  (was 6,3xx,xxx inlined)' % ('TOTAL', total))
    print('all datasets verified through pack -> unpack')


if __name__ == '__main__':
    main()
