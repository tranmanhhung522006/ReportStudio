import io, unittest
from docx import Document
from docx.oxml.ns import qn
from advanced import advanced, audit, integrity, field
from requirement_parser import parse_rules, extract_text, ai_rules

def sample():
    d=Document()
    d.add_paragraph('1. Introduction')
    d.add_paragraph('Research body text.')
    d.add_paragraph('Figure 9: Example image')
    d.add_paragraph('Table 8: Example table')
    p=d.add_paragraph('Citation: ')
    field(p,'ADDIN EN.CITE test','(Author, 2020)')
    d.add_table(rows=1,cols=1).cell(0,0).text='Example'
    b=io.BytesIO(); d.save(b); return b.getvalue()

class AdvancedTests(unittest.TestCase):
    def test_requirements(self):
        rules=parse_rules('Font: Times New Roman\nFont size: 13 pt\nLeft margin: 3 cm\nLine spacing: 1.5\nPaper: A4')['rules']
        self.assertEqual(rules['size'],13)
        self.assertEqual(rules['left'],3)
        self.assertEqual(rules['line'],1.5)
        self.assertEqual(rules['paper'],'A4')
    def test_extract(self):
        self.assertIn('Introduction',extract_text(sample(),'requirements.docx'))
    def test_consent(self):
        with self.assertRaises(ValueError): ai_rules('Font: Arial')
    def test_advanced(self):
        source=sample()
        out,r=advanced(source,{}, {0:1,1:0,2:0,3:0,4:0})
        self.assertTrue(r['integrity']['passed'])
        self.assertEqual(r['integrity']['citation_fields'],1)
        self.assertEqual(r['captions']['Figure'],1)
        d=Document(io.BytesIO(out))
        instructions=' '.join(x.text or '' for x in d._element.iter(qn('w:instrText')))
        self.assertIn('SEQ Figure',instructions)
        self.assertIn('SEQ Table',instructions)
        self.assertIn('PAGE',' '.join(x.text or '' for x in d.sections[0].footer._element.iter(qn('w:instrText'))))
        self.assertTrue(integrity(source,out)['passed'])
    def test_audit(self):
        self.assertGreater(audit(sample(),{})['failures'],0)
    def test_bad_paper(self):
        with self.assertRaises(ValueError): advanced(sample(),{'paper':'Invalid'}, {})
if __name__=='__main__': unittest.main()
