#!/usr/bin/env python3
"""Three-way merge for Word files: replay BASE -> MODIFIED onto TARGET, comments included.

    docx_merge.py check BASE MODIFIED TARGET [--report report.json]
    docx_merge.py merge BASE MODIFIED TARGET -o OUT [--resolutions res.json]
                        [--keep-target-on-conflict] [--report report.json] [--force]

Exit codes: 0 = clean, 2 = unresolved conflicts or manual items, 1 = usage or internal error.
"""
import argparse
import collections
import copy
import difflib
import json
import os
import re
import sys
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'
W16CEX = 'http://schemas.microsoft.com/office/word/2018/wordml/cex'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
OBJ = '￼'  # one image, equation or other object that has no plain text


def q(t):
    return '{%s}%s' % (W, t)


def lname(e):
    return etree.QName(e).localname if isinstance(e.tag, str) else ''


PARSER = dict(resolve_entities=False, no_network=True, huge_tree=True)


def parse(data):
    return etree.fromstring(data, etree.XMLParser(**PARSER))


def ser(root):
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------------------------------------------------------------- paragraph model

class Atom:
    __slots__ = ('ch', 'kind', 'el', 'off')

    def __init__(self, ch, kind, el, off=0):
        self.ch, self.kind, self.el, self.off = ch, kind, el, off


def atoms(p):
    """Visible characters of one paragraph, in order. Nested paragraphs and deleted text are skipped."""
    out = []

    def walk(e):
        for c in e:
            n = lname(c)
            if c.tag == q('pPr') or c.tag == q('rPr') or c.tag == q('sdtPr') or n in ('del', 'moveFrom', 'Fallback'):
                continue
            if c.tag == q('t'):
                for i, ch in enumerate(c.text or ''):
                    out.append(Atom(ch, 't', c, i))
            elif c.tag == q('tab') and e.tag == q('r'):
                out.append(Atom('\t', 'tab', c))
            elif c.tag == q('br') and e.tag == q('r'):
                out.append(Atom('\n', 'br', c))
            elif n in ('drawing', 'pict', 'object') or c.tag == '{%s}oMath' % M or c.tag == '{%s}oMathPara' % M:
                out.append(Atom(OBJ, 'obj', c))
            elif c.tag == q('p'):
                continue
            else:
                walk(c)
    walk(p)
    return out


def ptext(p):
    return ''.join(a.ch for a in atoms(p))


def body_paragraphs(root):
    return [p for p in root.iter(q('p')) if not any(a.tag == q('p') for a in p.iterancestors())]


def similar(a, b):
    if a == b:
        return 1.0
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a[:800], b[:800], autojunk=False).ratio()


def align(xs, ys, threshold=0.5):
    """Order-preserving pairs of paragraph indexes: exact matches first, then best match inside changed blocks."""
    pairs = {}
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, xs, ys, autojunk=False).get_opcodes():
        if op == 'equal':
            for a, b in zip(range(i1, i2), range(j1, j2)):
                pairs[a] = b
        elif op == 'replace':
            j = j1
            for i in range(i1, i2):
                best, score = None, threshold
                for c in range(j, j2):
                    if xs[i] and ys[c]:
                        r = similar(xs[i], ys[c])
                        if r > score:
                            best, score = c, r
                if best is not None:
                    pairs[i] = best
                    j = best + 1
            if i2 - i1 == j2 - j1:  # same-sized block with no text in common: blank lines still line up
                used = set(pairs.values())
                for a, b in zip(range(i1, i2), range(j1, j2)):
                    if a not in pairs and b not in used and not xs[a] and not ys[b]:
                        pairs[a] = b
    return pairs


# ---------------------------------------------------------------- three-way text merge

TOKEN = re.compile(r'[A-Za-z0-9_]+(?:[.,:%-][A-Za-z0-9_]+)*%?|\s+|.', re.S)


def opcodes(a, b):
    """Diff in character offsets, compared by whole Latin words and single CJK characters."""
    ta, tb = TOKEN.findall(a), TOKEN.findall(b)
    oa, ob = [0], [0]
    for t in ta:
        oa.append(oa[-1] + len(t))
    for t in tb:
        ob.append(ob[-1] + len(t))
    return [(op, oa[i1], oa[i2], ob[j1], ob[j2]) for op, i1, i2, j1, j2 in
            difflib.SequenceMatcher(None, ta, tb, autojunk=False).get_opcodes()]


def hunks(a, b):
    return [(i1, i2, b[j1:j2]) for op, i1, i2, j1, j2 in opcodes(a, b) if op != 'equal']


def project(a, b, region):
    """Offsets in `b` of the text that corresponds to the base range `region`."""
    s, e = region
    bs = be = None
    for op, i1, i2, j1, j2 in opcodes(a, b):
        if bs is None and i1 <= s <= i2:
            bs = j1 + (s - i1) if op == 'equal' else (j1 if s == i1 else j2)
        if i1 <= e <= i2:
            be = j1 + (e - i1) if op == 'equal' else (j2 if e == i2 else j1)
    bs = 0 if bs is None else bs
    be = bs if be is None or be < bs else be
    return bs, be


def three_way(a, b, c):
    """Yield regions over base text `a`: (kind, a_start, a_end, modified_text, target_text)."""
    ab, ac = hunks(a, b), hunks(a, c)
    tagged = sorted([(s, e, 'm') for s, e, _ in ab] + [(s, e, 't') for s, e, _ in ac])
    clusters = []
    for s, e, side in tagged:
        if clusters and s <= clusters[-1][1]:  # touching counts as overlapping: order is ambiguous there
            clusters[-1][1] = max(clusters[-1][1], e)
            clusters[-1][2].add(side)
        else:
            clusters.append([s, e, {side}])
    regions = []
    for s, e, sides in clusters:
        bs, be = project(a, b, (s, e))
        cs, ce = project(a, c, (s, e))
        mod, tgt = b[bs:be], c[cs:ce]
        if sides == {'t'}:
            continue
        if sides == {'m'}:
            kind = 'apply'
        elif mod == tgt:
            kind = 'same'
        else:
            kind = 'conflict'
        regions.append(dict(kind=kind, a=(s, e), base=a[s:e], modified=mod, target=tgt, c=(cs, ce)))
    return regions


# ---------------------------------------------------------------- run-level editing

def prune(el):
    """Remove an emptied run, then any content control left with nothing in it."""
    parent = el.getparent()
    if parent is None:
        return
    if el.tag == q('r') and not [c for c in el if c.tag != q('rPr')]:
        parent.remove(el)
    elif el.tag == q('sdtContent') and len(el) == 0:
        sdt = parent
        if sdt.getparent() is not None:
            prune_up = sdt.getparent()
            prune_up.remove(sdt)
            prune(prune_up) if prune_up.tag == q('sdtContent') else None
        return
    else:
        return
    if parent.tag == q('sdtContent'):
        prune(parent)


def run_of(atom):
    e = atom.el
    while e is not None and e.tag != q('r'):
        p = e.getparent()
        if p is not None and p.tag in (q('p'), q('hyperlink'), q('sdtContent'), q('ins'), q('smartTag')):
            return e  # object that is not inside a run, such as an equation
        e = p
    return e


def split_run_before(atom):
    """Return the element that starts exactly at `atom`, splitting its run when the atom is mid-run."""
    if atom.kind == 'obj' and atom.el.getparent() is not None and atom.el.getparent().tag != q('r'):
        return atom.el
    r = run_of(atom)
    if atom.kind == 't' and atom.off > 0:
        t = atom.el
        tail = copy.deepcopy(r)
        kids = list(r)
        idx = kids.index(t)
        for c in kids[idx + 1:]:
            r.remove(c)
        tkids = list(tail)
        tt = tkids[idx]
        for c in tkids[:idx]:
            if c.tag != q('rPr'):
                tail.remove(c)
        tt.text = t.text[atom.off:]
        tt.set(XML_SPACE, 'preserve')
        t.text = t.text[:atom.off]
        t.set(XML_SPACE, 'preserve')
        tail.tail = None
        r.addnext(tail)
        return tail
    # first char of its t, or a tab/br: split off anything before it in the same run
    kids = list(r)
    target = atom.el
    while target.getparent() is not r:
        target = target.getparent()
    idx = kids.index(target)
    if any(c.tag != q('rPr') for c in kids[:idx]):
        tail = copy.deepcopy(r)
        tkids = list(tail)
        for c in tkids[:idx]:
            if c.tag != q('rPr'):
                tail.remove(c)
        for c in kids[idx:]:
            r.remove(c)
        tail.tail = None
        r.addnext(tail)
        return tail
    return r


def split_run_after(atom):
    """Return the element after which new content can go, so that `atom` stays before it."""
    if atom.kind == 'obj' and atom.el.getparent() is not None and atom.el.getparent().tag != q('r'):
        return atom.el
    r = run_of(atom)
    kids = list(r)
    target = atom.el
    while target.getparent() is not r:
        target = target.getparent()
    idx = kids.index(target)
    if atom.kind == 't' and atom.off + 1 < len(atom.el.text):
        tail = copy.deepcopy(r)
        tkids = list(tail)
        for c in tkids[:idx]:
            if c.tag != q('rPr'):
                tail.remove(c)
        tt = tkids[idx]
        tt.text = atom.el.text[atom.off + 1:]
        tt.set(XML_SPACE, 'preserve')
        atom.el.text = atom.el.text[:atom.off + 1]
        atom.el.set(XML_SPACE, 'preserve')
        for c in kids[idx + 1:]:
            r.remove(c)
        tail.tail = None
        r.addnext(tail)
        return r
    if kids[idx + 1:]:
        tail = copy.deepcopy(r)
        tkids = list(tail)
        for c in tkids[:idx + 1]:
            if c.tag != q('rPr'):
                tail.remove(c)
        for c in kids[idx + 1:]:
            r.remove(c)
        tail.tail = None
        r.addnext(tail)
    return r


def new_text_run(fmt_from, text):
    r = etree.Element(q('r'))
    if fmt_from is not None and fmt_from.find(q('rPr')) is not None:
        r.append(copy.deepcopy(fmt_from.find(q('rPr'))))
    buf = ''

    def flush():
        nonlocal buf
        if buf:
            t = etree.SubElement(r, q('t'))
            t.text = buf
            t.set(XML_SPACE, 'preserve')
            buf = ''
    for ch in text:
        if ch == '\t':
            flush()
            etree.SubElement(r, q('tab'))
        elif ch == '\n':
            flush()
            etree.SubElement(r, q('br'))
        else:
            buf += ch
    flush()
    return r


def replace_range(p, start, end, text):
    """Replace visible characters [start, end) of paragraph p with text, keeping neighbouring formatting."""
    ats = atoms(p)
    victims = ats[start:end]
    plain = bool(text) and '\t' not in text and '\n' not in text
    host = next((a for a in ats[max(start - 1, 0):end + 1] if a.kind == 't'), None)
    fmt = copy.deepcopy(run_of(host)) if host is not None else None
    placed = False
    if plain and victims and victims[0].kind == 't':
        # stay in the replaced run, so an enclosing comment range keeps the new words
        t = victims[0].el
        at = victims[0].off
        t.text = t.text[:at] + text + t.text[at:]
        t.set(XML_SPACE, 'preserve')
        placed = True
        for a in victims:
            if a.el is t:
                a.off += len(text)
    by_t = collections.defaultdict(list)
    for a in victims:
        if a.kind == 't':
            by_t[id(a.el)].append(a)
    for group in by_t.values():
        t = group[0].el
        drop = {a.off for a in group}
        t.text = ''.join(ch for i, ch in enumerate(t.text) if i not in drop)
        t.set(XML_SPACE, 'preserve')
        if not t.text:
            r = t.getparent()
            r.remove(t)
            prune(r)
    for a in victims:
        if a.kind != 't':
            parent = a.el.getparent()
            parent.remove(a.el)
            prune(parent)
    if not text or placed:
        return
    now = atoms(p)
    before = now[start - 1] if start > 0 else None
    after = now[start] if start < len(now) else None
    if plain and before is not None and before.kind == 't':
        t = before.el
        t.text = t.text[:before.off + 1] + text + t.text[before.off + 1:]
        t.set(XML_SPACE, 'preserve')
    elif plain and after is not None and after.kind == 't':
        t = after.el
        t.text = t.text[:after.off] + text + t.text[after.off:]
        t.set(XML_SPACE, 'preserve')
    elif before is not None:
        split_run_after(before).addnext(new_text_run(fmt, text))
    elif after is not None:
        split_run_before(after).addprevious(new_text_run(fmt, text))
    else:
        ppr = p.find(q('pPr'))
        if ppr is not None:
            ppr.addnext(new_text_run(fmt, text))
        else:
            p.insert(0, new_text_run(fmt, text))


# ---------------------------------------------------------------- comments

class Comments:
    def __init__(self, z):
        self.root = parse(z.read('word/comments.xml')) if 'word/comments.xml' in z.namelist() else None

    def items(self):
        return [] if self.root is None else self.root.findall(q('comment'))


def comment_text(c):
    return ''.join(t.text or '' for t in c.iter(q('t')))


def comment_keys(cs):
    """(author, date, text, n): n separates comments that are otherwise identical."""
    seen = collections.Counter()
    out = {}
    for c in cs.items():
        k = (c.get(q('author')) or '', c.get(q('date')) or '', comment_text(c))
        out[c.get(q('id'))] = k + (seen[k],)
        seen[k] += 1
    return out


def anchor_info(root):
    """comment id -> (start paragraph, start offset, end paragraph, end offset), for a paragraph list."""
    paras = body_paragraphs(root)
    info = {}
    for pi, p in enumerate(paras):
        off = 0
        for e in walk_marks(p):
            if isinstance(e, Atom):
                off += 1
            elif e.tag == q('commentRangeStart'):
                info.setdefault(e.get(q('id')), [None, None, None, None])[:2] = [pi, off]
            elif e.tag == q('commentRangeEnd'):
                info.setdefault(e.get(q('id')), [None, None, None, None])[2:] = [pi, off]
    return info


def walk_marks(p):
    """Atoms and comment markers of one paragraph in document order."""
    out = []
    ats = atoms(p)  # held for the whole walk: lxml element ids are only stable while referenced
    at = collections.defaultdict(list)
    for a in ats:
        at[id(a.el)].append(a)

    def walk(e):
        for c in e:
            if id(c) in at:
                out.extend(at[id(c)])
                continue
            if c.tag in (q('commentRangeStart'), q('commentRangeEnd')):
                out.append(c)
                continue
            if c.tag == q('p') or c.tag == q('pPr'):
                continue
            walk(c)
    walk(p)
    return out


def anchor_text(paras, a):
    if a[0] is None or a[2] is None:
        return ''
    if a[0] == a[2]:
        return ptext(paras[a[0]])[a[1]:a[3]]
    parts = [ptext(paras[a[0]])[a[1]:]]
    parts += [ptext(paras[i]) for i in range(a[0] + 1, a[2])]
    parts.append(ptext(paras[a[2]])[:a[3]])
    return ''.join(parts)


def remove_comment_marks(root, cid):
    for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
        for e in [x for x in root.iter(q(tag)) if x.get(q('id')) == cid]:
            parent = e.getparent()
            parent.remove(e)
            if tag == 'commentReference' and parent.tag == q('r'):
                prune(parent)
            elif parent.tag == q('sdtContent'):
                prune(parent)


def map_offset(src, dst, off, prefer_next):
    """Offset in dst that corresponds to off in src."""
    if src == dst:
        return off
    best = None
    for op, i1, i2, j1, j2 in opcodes(src, dst):
        if op == 'equal' and i1 <= off <= i2:
            best = j1 + off - i1
            if prefer_next and off < i2:
                return best
            if not prefer_next and off > i1:
                return best
        elif i1 <= off <= i2 and best is None:
            best = j1 if prefer_next else j2
    return best if best is not None else min(off, len(dst))


def place_start(p, off, el):
    ats = atoms(p)
    if off < len(ats):
        split_run_before(ats[off]).addprevious(el)
    elif ats:
        split_run_after(ats[-1]).addnext(el)
    else:
        ppr = p.find(q('pPr'))
        if ppr is not None:
            ppr.addnext(el)
        else:
            p.insert(0, el)


def place_end(p, off, end_el, ref_run):
    ats = atoms(p)
    if off > 0 and ats:
        host = split_run_after(ats[min(off, len(ats)) - 1])
    else:
        host = None
    if host is not None:
        host.addnext(end_el)
    elif ats:
        split_run_before(ats[0]).addprevious(end_el)
    else:
        ppr = p.find(q('pPr'))
        if ppr is not None:
            ppr.addnext(end_el)
        else:
            p.insert(0, end_el)
    end_el.addnext(ref_run)


# ---------------------------------------------------------------- planning

def signature(p):
    pp = p.find(q('pPr'))
    return (etree.tostring(pp, method='c14n') if pp is not None else b'')


def plan_merge(za, zb, zc):
    ra, rb, rc = (parse(z.read('word/document.xml')) for z in (za, zb, zc))
    pa, pb, pc = body_paragraphs(ra), body_paragraphs(rb), body_paragraphs(rc)
    ta, tb, tc = [ptext(p) for p in pa], [ptext(p) for p in pb], [ptext(p) for p in pc]
    ab, ac = align(ta, tb), align(ta, tc, threshold=0.4)
    ba = {v: k for k, v in ab.items()}
    plan = dict(edits=[], conflicts=[], manual=[], same=0, already=0, inserts=[], deletes=[], warnings=[])

    for i, j in sorted(ab.items()):
        if ta[i] == tb[j]:
            continue
        if i not in ac:
            plan['conflicts'].append(dict(
                id='P%d-missing' % i, type='paragraph-missing-in-target',
                base=ta[i][:300], modified=tb[j][:300], target=None, context='',
                options=['keep_target'], para_a=i))
            continue
        c = ac[i]
        if tc[c] == tb[j]:
            plan['already'] += 1
            continue
        for n, reg in enumerate(three_way(ta[i], tb[j], tc[c])):
            cid = 'P%d-%d' % (i, n + 1)
            if reg['kind'] == 'same':
                plan['same'] += 1
            elif reg['kind'] == 'apply':
                if OBJ in reg['modified'] and reg['modified'].count(OBJ) > reg['base'].count(OBJ):
                    plan['manual'].append(dict(id=cid, why='modified adds an image or equation', context=ta[i][max(0, reg['a'][0] - 30):reg['a'][0]], text=reg['modified']))
                    continue
                plan['edits'].append(dict(c=c, range=reg['c'], text=reg['modified'], base=reg['base'], id=cid))
            else:
                plan['conflicts'].append(dict(
                    id=cid, type='both-changed', base=reg['base'], modified=reg['modified'], target=reg['target'],
                    context=ta[i][max(0, reg['a'][0] - 30):reg['a'][0]], options=['keep_target', 'use_modified', {'text': '...'}],
                    para_c=c, range=reg['c']))

    # paragraphs only in modified
    for j, t in enumerate(tb):
        if j in ba:
            continue
        prev = j - 1
        while prev >= 0 and prev not in ba:
            prev -= 1
        anchor_c = ac.get(ba[prev]) if prev >= 0 else None
        if OBJ in t:
            plan['manual'].append(dict(id='NEW-%d' % j, why='new paragraph holds an image or equation', context=tb[prev][:60] if prev >= 0 else '', text=t[:200]))
        elif prev >= 0 and anchor_c is None:
            plan['conflicts'].append(dict(id='NEW-%d' % j, type='insert-anchor-missing', base='', modified=t[:300], target=None, context=tb[prev][:60], options=['keep_target']))
        elif prev >= 0 and pb[j].getparent() is not pb[prev].getparent():
            plan['manual'].append(dict(id='NEW-%d' % j, why='new paragraph is in another table cell or block', context=tb[prev][:60], text=t[:200]))
        else:
            plan['inserts'].append(dict(j=j, after_c=anchor_c, prev_j=prev))

    # paragraphs only in base
    for i, t in enumerate(ta):
        if i in ab or not t:
            continue
        if i not in ac:
            plan['already'] += 1
            continue
        c = ac[i]
        if tc[c] != t:
            plan['conflicts'].append(dict(
                id='DEL-%d' % i, type='deleted-vs-edited', base=t[:300], modified='', target=tc[c][:300], context='',
                options=['keep_target', 'use_modified'], para_c=c, delete=True))
        elif pc[c].find(q('pPr')) is not None and pc[c].find(q('pPr')).find(q('sectPr')) is not None:
            plan['manual'].append(dict(id='DEL-%d' % i, why='paragraph holds a section break', context=t[:60], text=''))
        else:
            plan['deletes'].append(c)

    # format-only and structure differences are reported, not synced
    fmt = sum(1 for i, j in ab.items() if ta[i] == tb[j] and i in ac and signature(pa[i]) != signature(pb[j]))
    if fmt:
        plan['warnings'].append('%d paragraph(s) changed only in paragraph formatting; formatting is not synced.' % fmt)
    for tag, name in (('tbl', 'tables'), ('drawing', 'images')):
        x, y = sum(1 for _ in ra.iter(q(tag))), sum(1 for _ in rb.iter(q(tag)))
        if x != y:
            plan['warnings'].append('Number of %s changed from %d to %d; check them by hand.' % (name, x, y))
    for n in za.namelist():
        if re.fullmatch(r'word/(header|footer)\d*\.xml|word/(footnotes|endnotes)\.xml', n) and n in zb.namelist():
            tx = lambda z: ''.join(t.text or '' for t in parse(z.read(n)).iter(q('t')))
            if tx(za) != tx(zb):
                plan['warnings'].append('%s text changed; headers, footers and notes are not synced.' % n)
    if any(True for _ in rc.iter(q('ins'))) or any(True for _ in rc.iter(q('del'))):
        plan['warnings'].append('Target has tracked changes; deleted text is ignored during matching.')

    # comments
    ca, cb, cc = Comments(za), Comments(zb), Comments(zc)
    ka, kb, kc = comment_keys(ca), comment_keys(cb), comment_keys(cc)
    va, vb = set(ka.values()), set(kb.values())
    by_key_c = {v: k for k, v in kc.items()}
    by_key_b = {v: k for k, v in kb.items()}
    rm, add = sorted(va - vb), sorted(vb - va)
    plan['comment_remove'], plan['comment_add'] = [], []
    plan['comment_info'] = dict(removed=[], added=[], missing=[])
    for k in rm:
        if k in by_key_c:
            plan['comment_remove'].append(by_key_c[k])
            plan['comment_info']['removed'].append(k[2][:80])
        else:
            plan['comment_info']['missing'].append(k[2][:80])
    for k in add:
        if k in by_key_c:
            plan['already'] += 1
            continue
        plan['comment_add'].append(by_key_b[k])
        plan['comment_info']['added'].append(k[2][:80])
    plan['ctx'] = dict(ra=ra, rb=rb, rc=rc, pa=pa, pb=pb, pc=pc, ta=ta, tb=tb, tc=tc, ab=ab, ac=ac, ba=ba, cb=cb, cc=cc, kb=kb, kc=kc)
    return plan


def resolve(plan, resolutions):
    unresolved = []
    for c in plan['conflicts']:
        r = resolutions.get(c['id'])
        if r is None:
            unresolved.append(c)
            continue
        if r == 'keep_target':
            continue
        if c.get('delete'):
            if r == 'use_modified':
                plan['deletes'].append(c['para_c'])
            continue
        if c['type'] != 'both-changed':
            unresolved.append(c)
            continue
        text = c['modified'] if r == 'use_modified' else (r.get('text') if isinstance(r, dict) else None)
        if text is None:
            raise SystemExit('Resolution for %s must be keep_target, use_modified or {"text": ...}' % c['id'])
        plan['edits'].append(dict(c=c['para_c'], range=c['range'], text=text, base=c['target'], id=c['id']))
    return unresolved


# ---------------------------------------------------------------- applying

def next_para_id(used):
    n = 0x7A000000
    while True:
        n += 1
        s = '%08X' % n
        if s not in used:
            used.add(s)
            return s


def apply_plan(plan, zc):
    x = plan['ctx']
    rc, pc, pb, tb, ba, ac = x['rc'], x['pc'], x['pb'], x['tb'], x['ba'], x['ac']
    expected = {}
    by_para = collections.defaultdict(list)
    for e in plan['edits']:
        by_para[e['c']].append(e)
    for c, edits in by_para.items():
        for e in sorted(edits, key=lambda d: -d['range'][0]):
            replace_range(pc[c], e['range'][0], e['range'][1], e['text'])

    # insertions, kept in order after the same anchor
    last_after = {}
    new_for_b = {}
    for ins in plan['inserts']:
        j = ins['j']
        host = last_after.get(ins['after_c'])
        if host is None and ins['after_c'] is not None:
            host = pc[ins['after_c']]
        fmt = host if host is not None else (pc[0] if pc else None)
        np = etree.Element(q('p'))
        if fmt is not None and fmt.find(q('pPr')) is not None:
            pp = copy.deepcopy(fmt.find(q('pPr')))
            for s in pp.findall(q('sectPr')):
                pp.remove(s)
            np.append(pp)
        fmt_run = next((r for r in fmt.iter(q('r')) if r.find(q('t')) is not None), None) if fmt is not None else None
        if tb[j]:
            np.append(new_text_run(fmt_run, tb[j]))
        if host is not None:
            host.addnext(np)
        else:
            pc[0].addprevious(np)
        last_after[ins['after_c']] = np
        new_for_b[j] = np

    # comments: remove
    croot = x['cc'].root
    removed_para_ids = set()
    if croot is not None:
        for c in list(croot.findall(q('comment'))):
            if c.get(q('id')) in plan['comment_remove']:
                removed_para_ids |= {p.get('{%s}paraId' % W14) for p in c.iter(q('p'))}
                croot.remove(c)
    for cid in plan['comment_remove']:
        remove_comment_marks(rc, cid)

    # paragraph deletions happen after comment marks are gone
    for c in sorted(set(plan['deletes']), reverse=True):
        p = pc[c]
        parent = p.getparent()
        if parent.tag == q('tc') and len(parent.findall(q('p'))) == 1:
            n = len(atoms(p))
            replace_range(p, 0, n, '')
        else:
            for rng in [e for e in p.iter(q('commentRangeStart'), q('commentRangeEnd'))]:
                plan['warnings'].append('A kept comment was anchored on a deleted paragraph (id %s).' % rng.get(q('id')))
            parent.remove(p)

    # comments: add, anchored by mapping modified offsets onto the merged paragraph
    added_ids = {}
    if plan['comment_add']:
        if croot is None:
            croot = etree.Element(q('comments'), nsmap={'w': W, 'w14': W14})
            x['cc'].root = croot
        ids = [int(c.get(q('id'))) for c in croot.findall(q('comment'))] + [int(e.get(q('id'))) for e in rc.iter(q('commentRangeStart'))]
        nid = max(ids + [-1]) + 1
        used = set()
        for root_ in (rc, croot):
            used |= {e.get('{%s}paraId' % W14) for e in root_.iter(q('p')) if e.get('{%s}paraId' % W14)}
        anchors = anchor_info(x['rb'])
        for b_id in plan['comment_add']:
            src = next(c for c in x['cb'].items() if c.get(q('id')) == b_id)
            a = anchors.get(b_id)
            if a is None or a[0] is None or a[2] is None:
                plan['warnings'].append('Comment "%s" has no anchor in modified; it was skipped.' % comment_text(src)[:40])
                continue
            pp_s = target_para(a[0], ba, ac, pc, new_for_b)
            pp_e = target_para(a[2], ba, ac, pc, new_for_b)
            if pp_s is None or pp_e is None or pp_s.getparent() is None or pp_e.getparent() is None:
                plan['manual'].append(dict(id='COMMENT-%s' % b_id, why='anchor paragraph is not in the target', context=comment_text(src)[:60], text=''))
                continue
            sid = str(nid)
            nid += 1
            ts, te = ptext(pp_s), ptext(pp_e)
            s_off = map_offset(tb[a[0]], ts, a[1], True)
            e_off = map_offset(tb[a[2]], te, a[3], False)
            if pp_s is pp_e and e_off < s_off:
                e_off = s_off
            start = etree.Element(q('commentRangeStart'))
            start.set(q('id'), sid)
            end = etree.Element(q('commentRangeEnd'))
            end.set(q('id'), sid)
            ref = etree.Element(q('r'))
            cr = etree.SubElement(ref, q('commentReference'))
            cr.set(q('id'), sid)
            place_end(pp_e, e_off, end, ref)   # end first: it never shifts the start offset
            place_start(pp_s, s_off, start)
            want = anchor_text(x['pb'], a)
            got = (ptext(pp_s)[s_off:e_off] if pp_s is pp_e else None)
            if got is not None and similar(want, got) < 0.6:
                plan['warnings'].append('Comment "%s" was anchored approximately: wanted "%s", got "%s".' % (comment_text(src)[:30], want[:40], got[:40]))
            nc = copy.deepcopy(src)
            nc.tail = None
            nc.set(q('id'), sid)
            for e in list(nc.iter(q('pStyle'), q('rStyle'))):
                e.getparent().remove(e)
            for p_ in nc.iter(q('p')):
                p_.set('{%s}paraId' % W14, next_para_id(used))
                for at in ('textId',):
                    k = '{%s}%s' % (W14, at)
                    if k in p_.attrib:
                        del p_.attrib[k]
            croot.append(nc)
            added_ids[b_id] = nc
    return removed_para_ids, added_ids


def target_para(j, ba, ac, pc, new_for_b):
    if j in new_for_b:
        return new_for_b[j]
    i = ba.get(j)
    if i is None or i not in ac:
        return None
    return pc[ac[i]]


def comment_side_parts(zc, zb, plan, removed, added_ids):
    """Keep the companion comment parts consistent with comments.xml."""
    out = {}
    names = set(zc.namelist())
    bext = parse(zb.read('word/commentsExtended.xml')) if 'word/commentsExtended.xml' in zb.namelist() else None
    for part, attr in (('word/commentsExtended.xml', '{%s}paraId' % W15), ('word/commentsIds.xml', '{%s}paraId' % W16CID)):
        if part in names:
            root = parse(zc.read(part))
            for e in list(root):
                if e.get(attr) in removed:
                    root.remove(e)
            if part.endswith('Extended.xml'):
                for b_id, nc in added_ids.items():
                    src_para = next(c for c in plan['ctx']['cb'].items() if c.get(q('id')) == b_id).findall(q('p'))[-1].get('{%s}paraId' % W14)
                    done = '0'
                    if bext is not None:
                        for e in bext:
                            if e.get(attr) == src_para:
                                done = e.get('{%s}done' % W15) or '0'
                    n = etree.SubElement(root, '{%s}commentEx' % W15)
                    n.set(attr, nc.findall(q('p'))[-1].get('{%s}paraId' % W14))
                    n.set('{%s}done' % W15, done)
            out[part] = ser(root)
    return out


def ensure_comments_part(z, repl):
    names = set(z.namelist())
    if 'word/comments.xml' in names:
        return
    rels = parse(z.read('word/_rels/document.xml.rels'))
    ids = {r.get('Id') for r in rels}
    n = 1
    while 'rIdCmt%d' % n in ids:
        n += 1
    r = etree.SubElement(rels, '{%s}Relationship' % REL_NS)
    r.set('Id', 'rIdCmt%d' % n)
    r.set('Type', COMMENTS_REL)
    r.set('Target', 'comments.xml')
    repl['word/_rels/document.xml.rels'] = ser(rels)
    ct = parse(z.read('[Content_Types].xml'))
    o = etree.SubElement(ct, '{%s}Override' % CT_NS)
    o.set('PartName', '/word/comments.xml')
    o.set('ContentType', COMMENTS_CT)
    repl['[Content_Types].xml'] = ser(ct)


def verify_output(out_path, plan):
    """Re-read the file: untouched paragraphs must match the target and comment markers must balance."""
    z = zipfile.ZipFile(out_path)
    r = parse(z.read('word/document.xml'))
    ids = {t: [e.get(q('id')) for e in r.iter(q(t))] for t in ('commentRangeStart', 'commentRangeEnd', 'commentReference')}
    problems = []
    if not (sorted(ids['commentRangeStart']) == sorted(ids['commentRangeEnd']) == sorted(ids['commentReference'])):
        problems.append('comment markers are unbalanced')
    croot = parse(z.read('word/comments.xml')) if 'word/comments.xml' in z.namelist() else None
    cids = {c.get(q('id')) for c in croot.findall(q('comment'))} if croot is not None else set()
    if set(ids['commentRangeStart']) != cids:
        problems.append('comments.xml and document markers disagree')
    touched = {e['c'] for e in plan['edits']} | set(plan['deletes'])
    x = plan['ctx']
    want = [t for i, t in enumerate(x['tc']) if i not in touched]
    got = [ptext(p) for p in body_paragraphs(r)]
    sm = difflib.SequenceMatcher(None, want, got, autojunk=False)
    matched = sum(b.size for b in sm.get_matching_blocks())
    if matched != len(want):
        problems.append('%d untouched paragraph(s) changed' % (len(want) - matched))
    return problems


# ---------------------------------------------------------------- reporting and CLI

def summary(plan, unresolved=None):
    conflicts = plan['conflicts'] if unresolved is None else unresolved
    clean = lambda c: {k: v for k, v in c.items() if k not in ('para_a', 'para_c', 'range', 'delete')}
    return dict(
        status='conflict' if conflicts or plan['manual'] else 'clean',
        edits=len(plan['edits']),
        paragraphs_edited=len({e['c'] for e in plan['edits']}),
        paragraphs_inserted=len(plan['inserts']),
        paragraphs_deleted=len(set(plan['deletes'])),
        already_in_target=plan['already'],
        both_made_same_change=plan['same'],
        comments_removed=plan['comment_info']['removed'],
        comments_added=plan['comment_info']['added'],
        comments_already_gone=plan['comment_info']['missing'],
        conflicts=[clean(c) for c in conflicts],
        manual=plan['manual'],
        warnings=plan['warnings'],
        edit_list=[dict(id=e['id'], base=e['base'], new=e['text']) for e in plan['edits']],
    )


def emit(s, report):
    text = json.dumps(s, ensure_ascii=False, indent=2)
    if report:
        with open(report, 'w', encoding='utf-8') as f:
            f.write(text)
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    print(text)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    for name in ('check', 'merge'):
        s = sub.add_parser(name)
        s.add_argument('base')
        s.add_argument('modified')
        s.add_argument('target')
        s.add_argument('--report')
        if name == 'merge':
            s.add_argument('-o', '--output', required=True)
            s.add_argument('--resolutions')
            s.add_argument('--keep-target-on-conflict', action='store_true')
            s.add_argument('--force', action='store_true')
    a = ap.parse_args(argv)
    za, zb, zc = zipfile.ZipFile(a.base), zipfile.ZipFile(a.modified), zipfile.ZipFile(a.target)
    plan = plan_merge(za, zb, zc)
    if a.cmd == 'check':
        s = summary(plan)
        emit(s, a.report)
        return 2 if s['status'] == 'conflict' else 0

    inputs = {os.path.abspath(x) for x in (a.base, a.modified, a.target)}
    if os.path.abspath(a.output) in inputs:
        raise SystemExit('Output must not overwrite an input file.')
    if os.path.exists(a.output) and not a.force:
        raise SystemExit('Output already exists; pass --force to replace it.')
    res = {}
    if a.resolutions:
        with open(a.resolutions, encoding='utf-8') as f:
            res = json.load(f)
    if a.keep_target_on_conflict:
        for c in plan['conflicts']:
            res.setdefault(c['id'], 'keep_target')
    unresolved = resolve(plan, res)
    if unresolved or plan['manual']:
        s = summary(plan, unresolved)
        emit(s, a.report)
        print('Nothing written: resolve the conflicts and manual items first.', file=sys.stderr)
        return 2
    removed, added = apply_plan(plan, zc)
    repl = {'word/document.xml': ser(plan['ctx']['rc'])}
    if plan['ctx']['cc'].root is not None and (plan['comment_remove'] or plan['comment_add']):
        repl['word/comments.xml'] = ser(plan['ctx']['cc'].root)
        repl.update(comment_side_parts(zc, zb, plan, removed, added))
        ensure_comments_part(zc, repl)
    tmp = a.output + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as zo:
        for it in zc.infolist():
            zo.writestr(it, repl.get(it.filename, zc.read(it.filename)), compress_type=it.compress_type)
        have = set(zc.namelist())
        for n, data in repl.items():
            if n not in have:
                zo.writestr(n, data, compress_type=zipfile.ZIP_DEFLATED)
    problems = verify_output(tmp, plan)
    s = summary(plan, [])
    if plan['manual']:
        s['status'] = 'conflict'
    if problems:
        os.remove(tmp)
        s['status'] = 'failed'
        s['problems'] = problems
        emit(s, a.report)
        print('Nothing written: self-check failed.', file=sys.stderr)
        return 1
    os.replace(tmp, a.output)
    s['output'] = a.output
    emit(s, a.report)
    return 0


if __name__ == '__main__':
    sys.exit(main())
