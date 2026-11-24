import io, unittest
from docx import Document
from docx.shared import Pt
from engine import analyze
from advanced import advanced
from inspection import extract_sample, preview, reference_audit

def data(d):
    b=io.BytesIO(); d.save(b); return b.getvalue()

class InspectionTests(unittest.TestCase):
    def test_sample(self):
        d=Document(); p=d.add_paragraph(); r=p.add_run('Body text'); r.font.name='Arial'; r.font.size=Pt(12)
        c=extract_sample(data(d))['rules']
        self.assertEqual(c['font'],'Arial'); self.assertEqual(c['size'],12)
    def test_preview(self):
        d=Document(); d.add_heading('Introduction',1); d.add_paragraph('Body text.')
        self.assertEqual(preview(data(d))['blocks'][0]['level'],1)
    def test_list_not_heading(self):
        d=Document(); d.add_paragraph('1. Short item',style='List Number')
        self.assertEqual(analyze(data(d))['rows'][0]['level'],0)
    def test_prose_not_caption(self):
        d=Document(); d.add_heading('Introduction',1); d.add_paragraph('Figure 1 shows the result.')
        out,r=advanced(data(d),{'toc':False}, {0:1,1:0})
        self.assertIn('Figure 1 shows the result.', [p.text for p in Document(io.BytesIO(out)).paragraphs])
    def test_reference_duplicates(self):
        d=Document(); d.add_heading('References',1); d.add_paragraph('Author. 10.1234/example'); d.add_paragraph('Author. 10.1234/example')
        r=reference_audit(data(d)); self.assertEqual(len(r['duplicate_entries']),1); self.assertEqual(r['doi_candidates'],['10.1234/example'])
