import io, tempfile, unittest
from unittest.mock import patch
from docx import Document
from advanced import advanced, field
from citation_guard import compare_citations
from upload_cache import UploadCache, UploadExpired
import doi_check

def make_doc(citation='(Author, 2020)'):
    d=Document(); d.add_heading('Introduction',1)
    p=d.add_paragraph('Research: '); field(p,'ADDIN EN.CITE <EndNote><Cite><Author>Author</Author></Cite></EndNote>',citation)
    field(d.add_paragraph(),'ADDIN EN.REFLIST','Author. (2020). Example.')
    b=io.BytesIO(); d.save(b); return b.getvalue()

class V21Tests(unittest.TestCase):
    def test_upload_expiry_and_isolation(self):
        now=[0]
        with tempfile.TemporaryDirectory() as root:
            cache=UploadCache(root,clock=lambda:now[0])
            key=cache.put('a',b'abc')
            self.assertEqual(cache.get('a',key),b'abc')
            with self.assertRaises(UploadExpired): cache.get('b',key)
            now[0]=1801
            with self.assertRaises(UploadExpired): cache.get('a',key)
    def test_citation_preserved(self):
        original=make_doc()
        result,report=advanced(original,{}, {0:1,1:0,2:0})
        self.assertTrue(report['integrity']['citation_integrity']['passed'])
        self.assertEqual(report['integrity']['citation_integrity']['before'],2)
        self.assertFalse(compare_citations(original,make_doc('(Different, 2021)'))['passed'])
    def test_crossref(self):
        doi='10.1234/example-test'
        with patch('doi_check.fetch',return_value=('ok',{'message':{'DOI':doi,'title':['Example'],'published':{'date-parts':[[2024]]}}})):
            r=doi_check.check_many([doi])[0]
        self.assertEqual(r['status'],'verified'); self.assertEqual(r['year'],2024)
    def test_datacite_fallback(self):
        doi='10.1234/example-fallback'
        with patch('doi_check.fetch',side_effect=[('missing',None),('ok',{'data':{'attributes':{'doi':doi,'titles':[{'title':'Example'}],'publicationYear':2025}}})]):
            self.assertEqual(doi_check.check_many([doi])[0]['source'],'DataCite')
    def test_invalid_doi(self):
        with self.assertRaises(ValueError): doi_check.check_many(['https://example.com'])
