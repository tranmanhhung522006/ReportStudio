import os, unittest
from unittest.mock import patch
os.environ.setdefault('SECRET_KEY','test-only-key-'+'x'*40)
from production import app
from storage import safe_filename

class ProductionTests(unittest.TestCase):
    def setUp(self):
        app.config['TESTING']=True
        self.client=app.test_client()
    def test_health(self):
        self.assertEqual(self.client.get('/health').status_code,200)
    def test_csrf(self):
        self.assertEqual(self.client.post('/api/requirements',json={'text':'Font: Arial'}).status_code,403)
    def test_parse(self):
        r=self.client.post('/api/requirements',json={'text':'Font: Arial'},headers={'X-ReportStudio':'1'})
        self.assertEqual(r.status_code,200)
        self.assertEqual(r.json['rules']['font'],'Arial')
    def test_no_database(self):
        with patch('storage.configured',return_value=False):
            self.assertEqual(self.client.get('/api/history').json,{'files':[]})
    def test_safe_name(self):
        self.assertEqual(safe_filename('../../sample.docx'),'sample_formatted.docx')
    def test_session_isolation(self):
        with patch('production.database',return_value=True),patch('storage.list_files',return_value=[]) as query:
            self.client.get('/api/history')
            app.test_client().get('/api/history')
            self.assertNotEqual(query.call_args_list[0].args[0],query.call_args_list[1].args[0])

    def test_multipart(self):
        import io, json
        from docx import Document
        b=io.BytesIO(); d=Document(); d.add_heading('Introduction',1); d.save(b); b.seek(0)
        r=self.client.post('/api/analyze',data={'document':(b,'sample.docx'),'payload':json.dumps({})},headers={'X-ReportStudio':'1'})
        self.assertEqual(r.status_code,200)
        self.assertEqual(r.json['preview']['blocks'][0]['level'],1)
    def test_missing_file(self):
        with patch('production.database',return_value=True),patch('storage.get_file',return_value=None):
            self.assertEqual(self.client.get('/api/history/00000000-0000-0000-0000-000000000001').status_code,404)
