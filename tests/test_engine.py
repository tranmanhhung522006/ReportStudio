import unittest
from io import BytesIO
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from engine import analyze, format_doc

def fixture():
    d=Document()
    d.add_paragraph('1. Introduction')
    d.add_paragraph('1.1. Research Background')
    d.add_paragraph('This is a normal paragraph.')
    d.add_paragraph('References')
    p=d.add_paragraph('Citation')
    p.add_run()._r.append(OxmlElement('w:fldChar'))
    b=BytesIO(); d.save(b); return b.getvalue()

class EngineTests(unittest.TestCase):
    def test_detection(self):
        r=analyze(fixture())
        self.assertEqual([x['level'] for x in r['rows']],[1,2,0,1,0])
        self.assertTrue(r['rows'][-1]['protected'])
    def test_output(self):
        source=fixture()
        result,report=format_doc(source,{}, {0:1,1:2,2:0,3:1,4:0})
        d=Document(BytesIO(result))
        p=next(p for p in d.paragraphs if p.text=='Introduction')
        self.assertEqual(p.style.name,'Heading 1')
        self.assertIsNotNone(p._p.pPr.numPr)
        self.assertEqual(report['heading'],3)
        self.assertTrue(report['toc_added'])
        self.assertEqual(report['protected'],1)
        self.assertEqual(Document(BytesIO(source)).paragraphs[0].text,'1. Introduction')
    def test_invalid_config(self):
        with self.assertRaises(ValueError): format_doc(fixture(),{'size':100},{})
    def test_outline(self):
        d=Document(); d.add_paragraph('Custom section'); b=BytesIO();d.save(b)
        self.assertEqual(analyze(b.getvalue(),'1.2 Custom section')['rows'][0]['level'],2)

if __name__=='__main__': unittest.main()
