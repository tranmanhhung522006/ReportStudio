import io, unittest
from docx import Document
from docx.shared import Inches
from PIL import Image
from advanced import advanced
from document_preview import full_preview, compare_images

def save(doc):
    b=io.BytesIO();doc.save(b);return b.getvalue()

def picture():
    b=io.BytesIO();Image.new('RGB',(32,20),'blue').save(b,format='PNG');b.seek(0);return b

class PreviewTests(unittest.TestCase):
    def test_entire_body_with_images_and_tables(self):
        d=Document()
        for i in range(350):d.add_paragraph('Paragraph '+str(i))
        d.add_picture(picture(),width=Inches(1))
        cell=d.add_table(rows=1,cols=1).cell(0,0)
        cell.paragraphs[0].add_run().add_picture(picture())
        d.add_paragraph('Last paragraph')
        p=full_preview(save(d))
        self.assertEqual(len(p['blocks']),353)
        self.assertEqual(p['blocks'][-1]['segments'][0]['text'],'Last paragraph')
        self.assertEqual(p['image_count'],2)
        self.assertEqual(len(p['assets']),1)
    def test_image_preservation(self):
        d=Document();d.add_heading('Introduction',1)
        d.add_picture(picture(),width=Inches(1))
        d.add_table(rows=1,cols=1).cell(0,0).paragraphs[0].add_run().add_picture(picture())
        d.sections[0].header.paragraphs[0].add_run().add_picture(picture())
        raw=save(d);out,report=advanced(raw,{'toc':False},{0:1})
        self.assertTrue(report['integrity']['image_integrity']['passed'])
        self.assertEqual(compare_images(raw,out)['before'],3)
        self.assertFalse(compare_images(raw,save(Document()))['passed'])
