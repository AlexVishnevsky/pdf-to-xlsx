# -*- coding: utf-8 -*-
"""Рушій PDF -> XLSX з точним відтворенням розмітки (сітка, рамки, шрифти, печатки)."""
import os, math
from collections import Counter
import pymupdf as fitz
import xlsxwriter


PT2PX = 4.0 / 3.0
TOLX = 1.2
TOLY = 1.2


def cluster(vals, tol):
    vals = sorted(vals)
    out = []
    for v in vals:
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(g) / len(g) for g in out]


def snap(v, reps, tol):
    best, bd = None, tol
    for r in reps:
        d = abs(v - r)
        if d <= bd:
            bd, best = d, r
    return best


def merge_intervals(iv):
    iv = sorted(iv)
    out = []
    for a, b in iv:
        if out and a <= out[-1][1] + 0.6:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return out


def covered(intervals, a, b, frac=0.92):
    if b - a <= 0.01:
        return False
    tot = 0.0
    for x0, x1 in intervals:
        lo, hi = max(a, x0), min(b, x1)
        if hi > lo:
            tot += hi - lo
    return tot >= (b - a) * frac


def idx_of(cuts, v, tol=1.5):
    for i, c in enumerate(cuts):
        if abs(c - v) <= tol:
            return i
    return min(range(len(cuts)), key=lambda i: abs(cuts[i] - v))


def read_page(page):
    W, H = page.rect.width, page.rect.height
    hs, vs = [], []
    for dr in page.get_drawings():
        stroked = dr.get('color') is not None
        for it in dr['items']:
            if it[0] == 'l' and stroked:
                a, b = it[1], it[2]
                if abs(a.y - b.y) < 0.7 and abs(a.x - b.x) >= 0.7:
                    hs.append([(a.y + b.y) / 2, min(a.x, b.x), max(a.x, b.x)])
                elif abs(a.x - b.x) < 0.7 and abs(a.y - b.y) >= 0.7:
                    vs.append([(a.x + b.x) / 2, min(a.y, b.y), max(a.y, b.y)])
            elif it[0] == 're' and stroked:
                r = it[1]
                hs += [[r.y0, r.x0, r.x1], [r.y1, r.x0, r.x1]]
                vs += [[r.x0, r.y0, r.y1], [r.x1, r.y0, r.y1]]

    lines = []
    for b in page.get_text('rawdict')['blocks']:
        if b['type'] != 0:
            continue
        for l in b['lines']:
            runs = []
            for s in l['spans']:
                chars = s.get('chars', [])
                t = ''.join(c['c'] for c in chars)
                if not t.strip():
                    continue
                runs.append(dict(text=t, font=s['font'], size=s['size'],
                                 flags=s['flags'], color=s['color'],
                                 bbox=fitz.Rect(s['bbox']), chars=chars))
            if not runs:
                continue
            bb = fitz.Rect(runs[0]['bbox'])
            for r in runs[1:]:
                bb |= r['bbox']
            lines.append(dict(bbox=bb, runs=runs, y=bb.y0))

    imgs = []
    for im in page.get_images(full=True):
        for r in page.get_image_rects(im[0]):
            imgs.append(dict(xref=im[0], rect=fitz.Rect(r)))
    return dict(W=W, H=H, hs=hs, vs=vs, lines=lines, imgs=imgs, page=page)


def build_grid(pd):
    ys = cluster([s[0] for s in pd['hs']], TOLY)
    xs = cluster([s[0] for s in pd['vs']], TOLX)
    hmap, vmap = {}, {}
    for y, a, b in pd['hs']:
        hmap.setdefault(snap(y, ys, TOLY), []).append([a, b])
    for x, a, b in pd['vs']:
        vmap.setdefault(snap(x, xs, TOLX), []).append([a, b])
    for k in hmap:
        hmap[k] = merge_intervals(hmap[k])
    for k in vmap:
        vmap[k] = merge_intervals(vmap[k])
    pd['gx'], pd['gy'], pd['hmap'], pd['vmap'] = xs, ys, hmap, vmap


def unit_borders(pd, i, j):
    gx, gy = pd['gx'], pd['gy']
    y0, y1, x0, x1 = gy[i], gy[i + 1], gx[j], gx[j + 1]
    return (covered(pd['hmap'].get(gy[i], []), x0, x1),
            covered(pd['hmap'].get(gy[i + 1], []), x0, x1),
            covered(pd['vmap'].get(gx[j], []), y0, y1),
            covered(pd['vmap'].get(gx[j + 1], []), y0, y1))


def build_cells(pd):
    gx, gy = pd['gx'], pd['gy']
    ni, nj = len(gy) - 1, len(gx) - 1
    if ni <= 0 or nj <= 0:
        pd['cells'] = []
        return
    B = [[unit_borders(pd, i, j) for j in range(nj)] for i in range(ni)]
    active = [[any(B[i][j]) for j in range(nj)] for i in range(ni)]
    parent = {(i, j): (i, j) for i in range(ni) for j in range(nj)}

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    for i in range(ni):
        for j in range(nj):
            if not active[i][j]:
                continue
            if j + 1 < nj and active[i][j + 1] and not B[i][j][3]:
                union((i, j), (i, j + 1))
            if i + 1 < ni and active[i + 1][j] and not B[i][j][1]:
                union((i, j), (i + 1, j))
    groups = {}
    for i in range(ni):
        for j in range(nj):
            if active[i][j]:
                groups.setdefault(find((i, j)), []).append((i, j))
    cells = []
    for g in groups.values():
        i0 = min(p[0] for p in g); i1 = max(p[0] for p in g)
        j0 = min(p[1] for p in g); j1 = max(p[1] for p in g)
        rect = fitz.Rect(gx[j0], gy[i0], gx[j1 + 1], gy[i1 + 1])
        bd = dict(
            top=covered(pd['hmap'].get(gy[i0], []), gx[j0], gx[j1 + 1]),
            bottom=covered(pd['hmap'].get(gy[i1 + 1], []), gx[j0], gx[j1 + 1]),
            left=covered(pd['vmap'].get(gx[j0], []), gy[i0], gy[i1 + 1]),
            right=covered(pd['vmap'].get(gx[j1 + 1], []), gy[i0], gy[i1 + 1]))
        cells.append(dict(rect=rect, border=bd, lines=[]))
    pd['cells'] = cells


# --------------------------------------------------------------- text bits
FONTMAP = {
    'TimesNewRomanPSMT': 'Times New Roman',
    'TimesNewRomanPS-BoldMT': 'Times New Roman',
    'TimesNewRomanPS-ItalicMT': 'Times New Roman',
    'TimesNewRomanPS-BoldItalicMT': 'Times New Roman',
    'Calibri': 'Calibri', 'Calibri-Bold': 'Calibri',
    'ArialMT': 'Arial', 'Arial-BoldMT': 'Arial',
}


def font_name(f):
    if f in FONTMAP:
        return FONTMAP[f]
    f = f.split('+')[-1]
    for suf in ('PS-BoldItalicMT', 'PS-BoldMT', 'PS-ItalicMT', 'PSMT',
                '-BoldItalicMT', '-BoldMT', '-ItalicMT', 'MT',
                '-Bold', '-Italic', '-Regular'):
        if f.endswith(suf):
            f = f[:-len(suf)]
            break
    out = ''
    for i, ch in enumerate(f):
        if i and ch.isupper() and not f[i - 1].isupper():
            out += ' '
        out += ch
    return out.strip() or 'Times New Roman'


def run_style(r):
    return (font_name(r['font']), round(r['size'], 2),
            bool(r['flags'] & 16), bool(r['flags'] & 2), r['color'])


def norm(s):
    return ' '.join(s.split())


def line_text(l):
    return ''.join(r['text'] for r in l['runs'])


def has_tab_gap(l):
    return '    ' in line_text(l)


def split_segments(l, gap=3.2):
    """split a line into positioned pieces at visual gaps"""
    chars = []
    for r in l['runs']:
        for c in r.get('chars', []):
            chars.append((c, r))
    if not chars:
        t = norm(line_text(l))
        return [dict(bbox=l['bbox'], text=t, runs=l['runs'])] if t else []
    segs, cur = [], []
    prev_x1 = None
    for c, r in chars:
        bb = c['bbox']
        if c['c'] == ' ' and not cur:
            prev_x1 = bb[2]
            continue
        if prev_x1 is not None and bb[0] - prev_x1 > gap and cur:
            segs.append(cur)
            cur = []
        cur.append((c, r))
        prev_x1 = bb[2]
    if cur:
        segs.append(cur)
    out = []
    for s in segs:
        txt = ''.join(c['c'] for c, r in s).rstrip()
        if not txt.strip():
            continue
        x0 = min(c['bbox'][0] for c, r in s)
        x1 = max(c['bbox'][2] for c, r in s)
        y0 = min(c['bbox'][1] for c, r in s)
        y1 = max(c['bbox'][3] for c, r in s)
        runs, last = [], None
        for c, r in s:
            st = run_style(r)
            if last != st:
                runs.append(dict(style=st, text=c['c'], size=r['size']))
                last = st
            else:
                runs[-1]['text'] += c['c']
        out.append(dict(bbox=fitz.Rect(x0, y0, x1, y1), text=txt, runs=runs,
                        size=s[0][1]['size']))
    return out


def space_w(size):
    try:
        return fitz.get_text_length(' ', fontname='tiro', fontsize=size)
    except Exception:
        return 0.25 * size


def cell_paragraphs(cell):
    lines = sorted(cell['lines'], key=lambda l: (round(l['bbox'].y0, 1), l['bbox'].x0))
    if not lines:
        return []
    right = cell['rect'].x1
    paras, cur = [], [lines[0]]
    for prev, cu in zip(lines, lines[1:]):
        full = prev['bbox'].x1 >= right - 4.0
        tabbed = has_tab_gap(prev) or has_tab_gap(cu)
        same_row = abs(cu['bbox'].y0 - prev['bbox'].y0) < 0.5
        if full and not tabbed and not same_row:
            cur.append(cu)
        else:
            paras.append(cur)
            cur = [cu]
    paras.append(cur)
    return paras


def cell_runs(cell):
    paras = cell_paragraphs(cell)
    out = []
    for pi, p in enumerate(paras):
        if pi:
            out.append(dict(style=None, text='\n'))
        for li, l in enumerate(p):
            if li:
                out.append(dict(style=None, text=' '))
            segs = split_segments(l)
            for si, s in enumerate(segs):
                if si:
                    g = s['bbox'].x0 - segs[si - 1]['bbox'].x1
                    n = max(1, int(round(g / space_w(s['size']))))
                    out.append(dict(style=None, text=' ' * n))
                for r in s['runs']:
                    out.append(dict(style=r['style'], text=r['text']))
    merged = []
    for r in out:
        if merged and merged[-1]['style'] == r['style']:
            merged[-1]['text'] += r['text']
        else:
            merged.append(dict(r))
    for r in merged:
        if r['style'] is not None:
            r['text'] = r['text'].replace(' ', ' ')
    return merged


def squeeze_runs(runs):
    """collapse 2-3 space runs from justification; keep long tab-like gaps"""
    out = []
    for r in runs:
        t = r['text']
        res, i = '', 0
        while i < len(t):
            if t[i] == ' ':
                j = i
                while j < len(t) and t[j] == ' ':
                    j += 1
                n = j - i
                res += ' ' if n <= 3 else ' ' * n
                i = j
            else:
                res += t[i]
                i += 1
        out.append(dict(style=r['style'], text=res))
    return out

PAD = 1.45
MARGIN_IN = 0.2                 # page margin, inches (printer safe minimum)
MG = MARGIN_IN * 72.0
PITCH_A, PITCH_B = 1.15578, 1.18061   # Excel line pitch (pt) = A*size + B
MAXROW = 400.0                        # Excel row height limit (409.5)
K_CHOICES = [2.0, 2.5, 4.0, 5.0]      # 100/K must be an integer print scale


def xl_pitch(size):
    return PITCH_A * size + PITCH_B


def pick_scale(size, word_pitch):
    for K in K_CHOICES:
        if xl_pitch(size * K) / K - word_pitch <= -0.05:
            return K
    return K_CHOICES[-1]


def word_pitch_of(pages):
    c = Counter()
    for pd in pages:
        for cl in pd['cells']:
            ls = sorted(cl['lines'], key=lambda l: l['bbox'].y0)
            for a, b in zip(ls, ls[1:]):
                d = round(b['bbox'].y0 - a['bbox'].y0, 3)
                if 3.0 < d < 30.0:
                    c[d] += 1
    return c.most_common(1)[0][0] if c else 0.0


def main_size(pages):
    c = Counter()
    for pd in pages:
        for cl in pd['cells']:
            for l in cl['lines']:
                for r in l['runs']:
                    c[round(r['size'], 2)] += len(r['text'])
    return c.most_common(1)[0][0] if c else 10.0


def px_to_width(px):
    if px <= 0:
        return 0
    if px < 12:
        return px / 12.0
    return (px - 5.0) / 7.0


def analyse(doc):
    pages = []
    for page in doc:
        pd = read_page(page)
        build_grid(pd)
        build_cells(pd)
        free = []
        for l in pd['lines']:
            cx = (l['bbox'].x0 + l['bbox'].x1) / 2.0
            cy = (l['bbox'].y0 + l['bbox'].y1) / 2.0
            hit = None
            for cl in pd['cells']:
                r = cl['rect']
                if r.x0 - .3 <= cx <= r.x1 + .3 and r.y0 - .3 <= cy <= r.y1 + .3:
                    hit = cl
                    break
            (hit['lines'] if hit else free).append(l)
        pd['free'] = [s for l in free for s in split_segments(l)]
        pages.append(pd)
    return pages


def columns(pages, W):
    x0, x1 = MG, W - MG
    tabx = cluster([x for p in pages for x in p['gx']], TOLX)
    extra = []
    for p in pages:
        for s in p['free']:
            if snap(s['bbox'].x0, tabx, 0.9) is None:
                extra.append(s['bbox'].x0)
    extra = [x for x in cluster(extra, 1.0) if snap(x, tabx, 0.9) is None]
    vals = sorted(v for v in ([x0] + list(tabx) + list(extra) + [x1])
                  if x0 <= v <= x1)
    fin = [x0]
    for v in vals[1:]:
        if v - fin[-1] >= 0.8:
            fin.append(v)
    if x1 - fin[-1] >= 0.8:
        fin.append(x1)
    else:
        fin[-1] = x1
    return fin


def rows_of_page(pd):
    gy = list(pd['gy'])
    Y0, H = MG, pd['H'] - MG
    bands = []
    for s in pd['free']:
        cy = (s['bbox'].y0 + s['bbox'].y1) / 2.0
        s['in_grid'] = bool(gy) and gy[0] - .3 <= cy <= gy[-1] + .3
        if not s['in_grid']:
            h = max(s['bbox'].height, s['size'] * 1.34)
            bands.append([cy - h / 2.0, cy + h / 2.0])
    bands.sort()
    merged = []
    for b in bands:
        if merged and b[0] < merged[-1][1] + .2:
            if b[1] <= merged[-1][1] + .2:
                continue
            b[0] = merged[-1][1]
        merged.append(b)
    if gy:
        for b in merged:
            if b[1] > gy[0] and b[0] < gy[0]:
                b[1] = min(b[1], gy[0])
            if b[0] < gy[-1] and b[1] > gy[-1]:
                b[0] = max(b[0], gy[-1])
    cuts = [Y0, H] + gy + [v for b in merged for v in b]
    cuts = sorted(x for x in cuts if Y0 <= x <= H)
    fin = [cuts[0]]
    for v in cuts[1:]:
        if v - fin[-1] >= 0.35:
            fin.append(v)
    if H - fin[-1] >= 0.35:
        fin.append(H)
    else:
        fin[-1] = H
    return [(fin[i], fin[i + 1]) for i in range(len(fin) - 1)]


def cell_align(cell):
    lines = cell['lines']
    r = cell['rect']
    if not lines:
        return 'left'
    L, R = r.x0 + PAD, r.x1 - PAD
    cc = (r.x0 + r.x1) / 2.0
    if all(abs((l['bbox'].x0 + l['bbox'].x1) / 2.0 - cc) <= 1.8 for l in lines):
        return 'center'
    lef = all(l['bbox'].x0 - L <= 2.5 for l in lines)
    rig = all(R - l['bbox'].x1 <= 2.5 for l in lines)
    if rig and not lef:
        return 'right'
    if lef and len(lines) > 1:
        full = sum(1 for l in lines if l['bbox'].x1 >= R - 2.0)
        if full >= len(lines) - 1 and full >= 1:
            return 'justify'
    return 'left'


def cell_valign(cell):
    if not cell['lines']:
        return 'vcenter'
    r = cell['rect']
    t = min(l['bbox'].y0 for l in cell['lines'])
    b = max(l['bbox'].y1 for l in cell['lines'])
    d = (t + b) / 2.0 - (r.y0 + r.y1) / 2.0
    if abs(d) <= 2.2:
        return 'vcenter'
    return 'top' if d < 0 else 'bottom'


def convert(pdf, out, tmp):
    doc = fitz.open(pdf)
    pages = analyse(doc)
    W = max(p['W'] for p in pages)
    H = max(p['H'] for p in pages)
    xcuts = columns(pages, W)
    ncol = len(xcuts) - 1
    K = pick_scale(main_size(pages), word_pitch_of(pages))

    wb = xlsxwriter.Workbook(out)
    ws = wb.add_worksheet('Sheet1')
    cache = {}

    def F(**kw):
        key = tuple(sorted(kw.items()))
        if key not in cache:
            cache[key] = wb.add_format(dict(kw))
        return cache[key]

    def style_props(st):
        d = dict(font_name=st[0], font_size=round(st[1] * K, 2))
        if st[2]:
            d['bold'] = True
        if st[3]:
            d['italic'] = True
        if st[4]:
            d['font_color'] = '#%06X' % st[4]
        return d

    # column widths: cumulative pixel rounding keeps absolute x exact
    cum = [int(round((x - xcuts[0]) * K * PT2PX)) for x in xcuts]
    cw = [max(1, cum[j + 1] - cum[j]) for j in range(ncol)]
    cw[-1] = max(1, cw[-1] - 2)
    for j in range(ncol):
        ws.set_column(j, j, px_to_width(cw[j]))

    row0, breaks = 0, []
    for pi, pd in enumerate(pages):
        rows = []
        for a, b in rows_of_page(pd):
            n = max(1, int(math.ceil((b - a) * K / MAXROW)))
            step = (b - a) / n
            for t in range(n):
                rows.append((a + t * step, a + (t + 1) * step))
        nrow = len(rows)
        ccum = [round((y - rows[0][0]) * K / 0.75) * 0.75 for y in
                [a for a, b in rows] + [rows[-1][1]]]
        rh = [max(0.75, round(ccum[i + 1] - ccum[i], 2)) for i in range(nrow)]
        rh[-1] = max(0.75, round(rh[-1] - 1.5 * K, 2))
        for i in range(nrow):
            ws.set_row(row0 + i, rh[i])

        def ridx(y):
            for i, (a, b) in enumerate(rows):
                if y <= b + .35:
                    return i
            return nrow - 1

        occupied = set()
        starts = [a for a, b in rows]
        ends = [b for a, b in rows]

        def put(r0, c0, r1, c1, runs, base):
            fmt = F(**base)
            txt = ''.join(x['text'] for x in runs)
            styles = set(x['style'] for x in runs if x['style'])
            rich = len(styles) > 1 and len(runs) > 1
            if r0 != r1 or c0 != c1:
                ws.merge_range(row0 + r0, c0, row0 + r1, c1,
                               '' if rich else txt, fmt)
            if rich:
                st0 = next(x['style'] for x in runs if x['style'])
                parts = []
                for x in runs:
                    parts.append(F(**style_props(x['style'] or st0)))
                    parts.append(x['text'])
                ws.write_rich_string(row0 + r0, c0, *parts, fmt)
            elif r0 == r1 and c0 == c1:
                ws.write_string(row0 + r0, c0, txt, fmt)

        for cl in sorted(pd['cells'], key=lambda c: (c['rect'].y0, c['rect'].x0)):
            r = cl['rect']
            r0 = idx_of(starts, r.y0, 1.0)
            r1 = max(r0, idx_of(ends, r.y1, 1.0))
            c0 = idx_of(xcuts, r.x0, 1.0)
            c1 = max(c0, idx_of(xcuts, r.x1, 1.0) - 1)
            runs = [x for x in squeeze_runs(cell_runs(cl)) if x['text'] != '']
            base = dict(text_wrap=True, align=cell_align(cl),
                        valign=cell_valign(cl))
            for side in ('top', 'bottom', 'left', 'right'):
                if cl['border'][side]:
                    base[side] = 1
            st = next((x['style'] for x in runs if x['style']), None)
            if st:
                base.update(style_props(st))
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    occupied.add((rr, cc))
            put(r0, c0, r1, c1, runs, base)

        byrow = {}
        for s in pd['free']:
            byrow.setdefault(ridx((s['bbox'].y0 + s['bbox'].y1) / 2.0), []).append(s)
        for ri, segs in byrow.items():
            segs.sort(key=lambda s: s['bbox'].x0)
            st_cols = [idx_of(xcuts, s['bbox'].x0, 1.2) for s in segs]
            for n, s in enumerate(segs):
                c0 = st_cols[n]
                c1 = min(st_cols[n + 1] - 1 if n + 1 < len(segs) else ncol - 1, ncol - 1)
                if (ri, c0) in occupied:
                    continue
                while c1 > c0 and (ri, c1) in occupied:
                    c1 -= 1
                c1 = max(c0, c1)
                runs = squeeze_runs(s['runs'])
                base = dict(align='left', valign='vcenter')
                base.update(style_props(runs[0]['style']))
                for cc in range(c0, c1 + 1):
                    occupied.add((ri, cc))
                put(ri, c0, ri, c1, runs, base)

        for n, im in enumerate(pd['imgs']):
            r = im['rect']
            png = os.path.join(tmp, 'im_%d_%d_%d.png' % (pi, n, im['xref']))
            pix = fitz.Pixmap(doc, im['xref'])
            sm = doc.extract_image(im['xref']).get('smask')
            if sm:
                try:
                    pix = fitz.Pixmap(pix, fitz.Pixmap(doc, sm))
                except Exception:
                    pass
            if pix.colorspace and pix.colorspace.n == 4:
                pix = fitz.Pixmap(fitz.csRGB, pix)
            pix.save(png)
            ri = ridx(r.y0 + .1)
            ci = idx_of(xcuts, r.x0, 1e9)
            if xcuts[ci] > r.x0 and ci > 0:
                ci -= 1
            ws.insert_image(row0 + ri, ci, png, {
                'x_offset': max(0, int(round((r.x0 - xcuts[ci]) * K * PT2PX))),
                'y_offset': max(0, int(round((r.y0 - rows[ri][0]) * K * PT2PX))),
                'x_scale': (r.width * K * PT2PX) / pix.width,
                'y_scale': (r.height * K * PT2PX) / pix.height,
                'object_position': 3})

        row0 += nrow
        if pi < len(pages) - 1:
            breaks.append(row0)

    ws.print_area(0, 0, row0 - 1, ncol - 1)
    ws.set_paper(9)
    ws.set_portrait()
    ws.set_margins(MARGIN_IN, MARGIN_IN, MARGIN_IN, MARGIN_IN)
    ws.set_header('', {'margin': 0})
    ws.set_footer('', {'margin': 0})
    ws.fit_to_pages(1, 0)
    ws.set_zoom(max(10, min(400, int(round(100.0 / K)))))
    ws.hide_gridlines(2)
    if breaks:
        ws.set_h_pagebreaks(breaks)
    ws.ignore_errors({'number_stored_as_text': 'A1:%s%d' %
                      (xlsxwriter.utility.xl_col_to_name(ncol - 1), row0)})
    wb.close()
    return dict(cols=ncol, rows=row0, pages=len(pages), K=K)




# ------------------------------------------------------------ public API
class NotConvertible(Exception):
    """PDF, який цим способом не перетворити (скан, порожній, захищений)."""


def output_path(pdf):
    """Шлях .xlsx поруч із PDF; наявний файл не перезаписуємо — додаємо (2), (3)…"""
    base = os.path.splitext(pdf)[0]
    out, n = base + '.xlsx', 2
    while os.path.exists(out):
        out = '%s (%d).xlsx' % (base, n)
        n += 1
    return out


def convert_file(pdf):
    """Перетворює один PDF; повертає (шлях xlsx, кількість сторінок)."""
    import tempfile
    try:
        doc = fitz.open(pdf)
    except Exception as e:
        raise NotConvertible('не вдалося відкрити PDF (%s)' % e)
    if doc.needs_pass:
        raise NotConvertible('PDF захищений паролем')
    if doc.page_count == 0:
        raise NotConvertible('у PDF немає сторінок')
    if not any(p.get_text().strip() for p in doc):
        raise NotConvertible('це скан без текстового шару — тут нічого розпізнати')
    doc.close()
    out = output_path(pdf)
    with tempfile.TemporaryDirectory(prefix='pdf2xlsx_') as tmp:
        try:
            info = convert(pdf, out, tmp)
        except Exception:
            if os.path.exists(out):
                try:
                    os.remove(out)
                except OSError:
                    pass
            raise
    return out, info['pages']
