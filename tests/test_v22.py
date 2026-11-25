import io, unittest, os
from unittest.mock import patch
from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from docx.oxml.ns import qn
from advanced import advanced
from engine import analyze

def blob(d):
    b=io.BytesIO();d.save(b);return b.getvalue()

class V22Tests(unittest.TestCase):
    def test_cover_no_section_is_noop(self):
        d=Document();d.add_paragraph('Cover')
        raw=blob(d);out,r=advanced(raw,{'protect_cover':True},{})
        self.assertEqual(raw,out)
    def test_cover_preserved(self):
        d=Document();r=d.add_paragraph().add_run('Cover');r.font.size=Pt(24)
        d.add_section(WD_SECTION.NEW_PAGE);d.add_heading('Introduction',1)
        out,r=advanced(blob(d),{'protect_cover':True,'toc':False},{})
        self.assertTrue(r['cover_protected'])
        self.assertEqual(Document(io.BytesIO(out)).paragraphs[0].runs[0].font.size.pt,24)
    def test_selective_body(self):
        d=Document();p=d.add_heading('Heading',1);p.runs[0].font.size=Pt(24)
        d.add_paragraph('Body  text')
        out,r=advanced(blob(d),{'format_headings':False,'toc':False,'cleanup_text':True},{})
        doc=Document(io.BytesIO(out))
        self.assertEqual(doc.paragraphs[0].runs[0].font.size.pt,24)
        self.assertEqual(doc.paragraphs[1].text,'Body text')
    def test_smart_numbering(self):
        d=Document();d.add_paragraph('Cover');d.add_section(WD_SECTION.NEW_PAGE);d.add_paragraph('Front matter');d.add_section(WD_SECTION.NEW_PAGE);d.add_heading('Introduction',1)
        out,r=advanced(blob(d),{'smart_pages':True,'protect_cover':True,'toc':False},{})
        self.assertTrue(r['smart_pages_applied'])
        sections=Document(io.BytesIO(out)).sections
        self.assertEqual(sections[1]._sectPr.find(qn('w:pgNumType')).get(qn('w:fmt')),'lowerRoman')
        self.assertEqual(sections[2]._sectPr.find(qn('w:pgNumType')).get(qn('w:start')),'1')
    def test_title_not_heading(self):
        d=Document();d.add_paragraph('Research proposal',style='Title')
        self.assertEqual(analyze(blob(d))['rows'][0]['level'],0)
    def test_access(self):
        import base64
        os.environ.setdefault('SECRET_KEY','test-secret-'+'x'*40)
        from production import app
        with patch.dict(os.environ,{'PERSONAL_ACCESS_CODE':'test-code'}):
            client=app.test_client()
            self.assertEqual(client.get('/').status_code,401)
            token=base64.b64encode(b'studio:test-code').decode()
            self.assertEqual(client.get('/',headers={'Authorization':'Basic '+token}).status_code,200)
            self.assertEqual(client.get('/health').status_code,200)

    def test_missing_heading_styles(self):
        from engine import heading_level
        d=Document()
        for n in range(2,10):
            d.styles[f'Heading {n}'].delete()
        d.add_paragraph('Introduction')
        d.add_paragraph('Deep subsection')
        out,r=advanced(blob(d),{'toc':False},{0:1,1:7})
        doc=Document(io.BytesIO(out))
        self.assertEqual(heading_level(doc.paragraphs[1]),7)
        self.assertEqual(doc.styles['Heading 7']._element.pPr.find(qn('w:outlineLvl')).get(qn('w:val')),'6')
        self.assertTrue(r['integrity']['passed'])
    def test_missing_heading_styles_with_cover(self):
        from engine import heading_level
        d=Document();d.add_paragraph('Cover')
        d.add_section(WD_SECTION.NEW_PAGE);d.add_paragraph('Introduction')
        d.styles['Heading 7'].delete()
        out,r=advanced(blob(d),{'protect_cover':True,'toc':False},{2:7})
        doc=Document(io.BytesIO(out))
        self.assertEqual(doc.paragraphs[0].text,'Cover')
        self.assertEqual(heading_level(doc.paragraphs[2]),7)
