"""End-to-end tests on small generated Word files. Run: python -m unittest discover tests"""
import json
import os
import sys
import tempfile
import unittest
import zipfile
from xml.sax.saxutils import escape

from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'skills', 'docx-three-way-merge', 'scripts'))
import docx_merge as dm  # noqa: E402

W = dm.W
NS = 'xmlns:w="%s" xmlns:w14="%s" xmlns:w15="%s"' % (W, dm.W14, dm.W15)
BOLD = '<w:rPr><w:b/></w:rPr>'
IMG = '<w:r><w:pict/></w:r>'
SDT = '<w:sdt><w:sdtPr><w:tag w:val="x"/></w:sdtPr><w:sdtContent>%s</w:sdtContent></w:sdt>'


def r(text, rpr=''):
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(text))


def p(*parts):
    return '<w:p>%s</w:p>' % ''.join(x if x.startswith('<') else r(x) for x in parts)


def anchored(cid, text):
    return ('<w:commentRangeStart w:id="%s"/>%s<w:commentRangeEnd w:id="%s"/><w:r><w:commentReference w:id="%s"/></w:r>'
            % (cid, r(text), cid, cid))


def comment(cid, text, date='2026-01-01T00:00:00Z', para='0000000%s'):
    return ('<w:comment w:id="%s" w:author="Tester" w:date="%s"><w:p w14:paraId="%s"><w:r><w:t>%s</w:t></w:r></w:p></w:comment>'
            % (cid, date, para % cid, escape(text)))


def build(path, paras, comments=None):
    body = '<w:document %s><w:body>%s<w:sectPr/></w:body></w:document>' % (NS, ''.join(paras))
    ct = ('<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' % dm.CT_NS)
    rel = '<Relationships xmlns="%s">' % dm.REL_NS
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('word/document.xml', body)
        if comments is not None:
            z.writestr('word/comments.xml', '<w:comments %s>%s</w:comments>' % (NS, ''.join(comments)))
            ct += '<Override PartName="/word/comments.xml" ContentType="%s"/>' % dm.COMMENTS_CT
            rel += '<Relationship Id="rId9" Type="%s" Target="comments.xml"/>' % dm.COMMENTS_REL
        z.writestr('[Content_Types].xml', ct + '</Types>')
        z.writestr('word/_rels/document.xml.rels', rel + '</Relationships>')


def read(path):
    z = zipfile.ZipFile(path)
    root = dm.parse(z.read('word/document.xml'))
    texts = [dm.ptext(x) for x in dm.body_paragraphs(root)]
    cs = []
    if 'word/comments.xml' in z.namelist():
        cs = [dm.comment_text(c) for c in dm.parse(z.read('word/comments.xml')).findall(dm.q('comment'))]
    return root, texts, cs


class Case(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.d = self.tmp.name
        self.base, self.mod, self.tgt = [os.path.join(self.d, n + '.docx') for n in 'abc']
        self.out = os.path.join(self.d, 'out.docx')

    def run_merge(self, *extra):
        return dm.main(['merge', self.base, self.mod, self.tgt, '-o', self.out, '--report', os.path.join(self.d, 'r.json')] + list(extra))

    def report(self):
        with open(os.path.join(self.d, 'r.json'), encoding='utf-8') as f:
            return json.load(f)

    def test_clean_edit_lands_on_differently_worded_target(self):
        build(self.base, [p('The rain fell in the north today.'), p('Second line stays.')])
        build(self.mod, [p('The heavy rain fell in the north today.'), p('Second line stays.')])
        build(self.tgt, [p('The rain fell in the north today, as reported.'), p('Second line stays.')])
        self.assertEqual(self.run_merge(), 0)
        self.assertEqual(read(self.out)[1][0], 'The heavy rain fell in the north today, as reported.')

    def test_edit_keeps_target_run_formatting_and_content_controls(self):
        para = '<w:r>%s<w:t xml:space="preserve">Alpha beta gamma</w:t></w:r>' % BOLD
        build(self.base, [p('Alpha beta gamma')])
        build(self.mod, [p('Alpha BETA gamma')])
        build(self.tgt, ['<w:p>%s</w:p>' % (SDT % para)])
        self.assertEqual(self.run_merge(), 0)
        root, texts, _ = read(self.out)
        self.assertEqual(texts[0], 'Alpha BETA gamma')
        self.assertTrue(all(x.find(dm.q('rPr')).find(dm.q('b')) is not None for x in root.iter(dm.q('r'))))
        self.assertEqual(sum(1 for _ in root.iter(dm.q('sdt'))), 1)

    def test_conflict_blocks_output_until_resolved(self):
        build(self.base, [p('Value is red today.')])
        build(self.mod, [p('Value is blue today.')])
        build(self.tgt, [p('Value is green today.')])
        self.assertEqual(dm.main(['check', self.base, self.mod, self.tgt, '--report', os.path.join(self.d, 'r.json')]), 2)
        conflict = self.report()['conflicts'][0]
        self.assertEqual((conflict['modified'], conflict['target']), ('blue', 'green'))
        self.assertEqual(self.run_merge(), 2)
        self.assertFalse(os.path.exists(self.out))
        res = os.path.join(self.d, 'res.json')
        for choice, want in (('use_modified', 'blue'), ('keep_target', 'green'), ({'text': 'teal'}, 'teal')):
            with open(res, 'w', encoding='utf-8') as f:
                json.dump({conflict['id']: choice}, f)
            self.assertEqual(self.run_merge('--resolutions', res, '--force'), 0)
            self.assertEqual(read(self.out)[1][0], 'Value is %s today.' % want)

    def test_same_change_on_both_sides_is_not_a_conflict(self):
        build(self.base, [p('Colour of the sea.')])
        build(self.mod, [p('Color of the sea.')])
        build(self.tgt, [p('Color of the sea.')])
        self.assertEqual(dm.main(['check', self.base, self.mod, self.tgt]), 0)

    def test_deleted_image_goes_away_with_its_sentence(self):
        build(self.base, [p('Keep this. Drop ', IMG, ' this part. End.')])
        build(self.mod, [p('Keep this. End.')])
        build(self.tgt, [p('Keep this. Drop ', IMG, ' this part. End.')])
        self.assertEqual(self.run_merge(), 0)
        root, texts, _ = read(self.out)
        self.assertEqual(texts[0], 'Keep this. End.')
        self.assertEqual(sum(1 for _ in root.iter(dm.q('pict'))), 0)

    def test_new_image_in_modified_needs_manual_handling(self):
        build(self.base, [p('Figure goes here.')])
        build(self.mod, [p('Figure ', IMG, ' goes here.')])
        build(self.tgt, [p('Figure goes here.')])
        self.assertEqual(self.run_merge(), 2)
        self.assertTrue(self.report()['manual'])
        self.assertFalse(os.path.exists(self.out))

    def test_paragraph_insert_and_delete(self):
        build(self.base, [p('One.'), p('Two to remove entirely.'), p('Three.')])
        build(self.mod, [p('One.'), p('A brand new paragraph.'), p('Three.')])
        build(self.tgt, [p('One.'), p('Two to remove entirely.'), p('Three.')])
        self.assertEqual(self.run_merge(), 0)
        self.assertEqual(read(self.out)[1], ['One.', 'A brand new paragraph.', 'Three.'])

    def test_delete_versus_target_edit_is_a_conflict(self):
        build(self.base, [p('One.'), p('The paragraph that goes away.'), p('Three.')])
        build(self.mod, [p('One.'), p('Three.')])
        build(self.tgt, [p('One.'), p('The paragraph that goes away, now edited.'), p('Three.')])
        self.assertEqual(self.run_merge(), 2)
        self.assertEqual(self.report()['conflicts'][0]['type'], 'deleted-vs-edited')

    def test_comment_add_remove_and_keep(self):
        cs = [comment('0', 'keep me'), comment('1', 'remove me')]
        build(self.base, [p('Start ', anchored('0', 'kept words'), ' mid ', anchored('1', 'gone words'), ' tail words here.')], cs)
        build(self.mod, [p('Start ', anchored('0', 'kept words'), ' mid gone words ', anchored('2', 'tail words'), ' here.')],
              [comment('0', 'keep me'), comment('2', 'brand new', '2026-02-02T00:00:00Z')])
        build(self.tgt, [p('Intro. Start ', anchored('5', 'kept words'), ' mid ', anchored('6', 'gone words'), ' tail words here.')],
              [comment('5', 'keep me'), comment('6', 'remove me')])
        self.assertEqual(self.run_merge(), 0)
        root, texts, cs_out = read(self.out)
        self.assertEqual(texts[0], 'Intro. Start kept words mid gone words tail words here.')
        self.assertEqual(sorted(cs_out), ['brand new', 'keep me'])
        anchors = dm.anchor_info(root)
        paras = dm.body_paragraphs(root)
        ids = {dm.comment_text(c): c.get(dm.q('id')) for c in dm.parse(zipfile.ZipFile(self.out).read('word/comments.xml')).findall(dm.q('comment'))}
        self.assertEqual(dm.anchor_text(paras, anchors[ids['brand new']]), 'tail words')
        self.assertEqual(dm.anchor_text(paras, anchors[ids['keep me']]), 'kept words')

    def test_comment_added_to_target_without_comments_part(self):
        build(self.base, [p('Plain words live here.')], [])
        build(self.mod, [p('Plain ', anchored('0', 'words'), ' live here.')], [comment('0', 'first note')])
        build(self.tgt, [p('Plain words live here.')])
        self.assertEqual(self.run_merge(), 0)
        z = zipfile.ZipFile(self.out)
        self.assertIn('word/comments.xml', z.namelist())
        self.assertIn(b'comments.xml', z.read('[Content_Types].xml'))
        self.assertEqual(read(self.out)[2], ['first note'])

    def test_refuses_to_overwrite_an_input(self):
        build(self.base, [p('A')])
        build(self.mod, [p('B')])
        build(self.tgt, [p('A')])
        with self.assertRaises(SystemExit):
            dm.main(['merge', self.base, self.mod, self.tgt, '-o', self.tgt])

    def test_other_package_parts_are_left_alone(self):
        build(self.base, [p('Text one.')])
        build(self.mod, [p('Text two.')])
        build(self.tgt, [p('Text one.')])
        self.assertEqual(self.run_merge(), 0)
        a, b = zipfile.ZipFile(self.tgt), zipfile.ZipFile(self.out)
        self.assertEqual(a.namelist(), b.namelist())
        for n in a.namelist():
            if n != 'word/document.xml':
                self.assertEqual(a.read(n), b.read(n))


if __name__ == '__main__':
    unittest.main()
