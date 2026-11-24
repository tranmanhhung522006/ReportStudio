"""Local HTTP service; documents are processed in memory."""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import base64, json, webbrowser, threading
from engine import analyze, format_doc, MAX_BYTES
ROOT=Path(__file__).resolve().parent/'web'
class Handler(BaseHTTPRequestHandler):
    def send(self, code, data, content_type='application/json; charset=utf-8'):
        if isinstance(data,dict): data=json.dumps(data,ensure_ascii=False).encode()
        self.send_response(code); self.send_header('Content-Type',content_type); self.send_header('Content-Length',str(len(data))); self.send_header('Cache-Control','no-store'); self.send_header('X-Content-Type-Options','nosniff'); self.send_header('Content-Security-Policy',"default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self' data:; frame-ancestors 'none'; base-uri 'none'"); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        mapping={'/':('index.html','text/html; charset=utf-8'),'/style.css':('style.css','text/css; charset=utf-8'),'/app.js':('app.js','text/javascript; charset=utf-8')}
        if self.path not in mapping: return self.send(404,{'error':'Not found'})
        name,mime=mapping[self.path]; self.send(200,(ROOT/name).read_bytes(),mime)
    def do_POST(self):
        if self.path not in ('/api/analyze','/api/format'): return self.send(404,{'error':'Not found'})
        if self.headers.get('Origin') not in (None,'http://127.0.0.1:8765','http://localhost:8765'): return self.send(403,{'error':'Invalid origin'})
        if self.headers.get('Host') not in ('127.0.0.1:8765','localhost:8765'): return self.send(403,{'error':'Invalid host'})
        if self.headers.get_content_type()!='application/json': return self.send(415,{'error':'JSON required'})
        try:
            n=int(self.headers.get('Content-Length','0'))
            if not 0<n<MAX_BYTES*2: raise ValueError('Invalid request size.')
            req=json.loads(self.rfile.read(n)); data=base64.b64decode(req['document'],validate=True)
            if self.path=='/api/analyze': self.send(200,analyze(data,req.get('outline','')))
            else:
                out, report=format_doc(data,req['config'],req['decisions']); self.send(200,{'document':base64.b64encode(out).decode(),'report':report})
        except Exception as e: self.send(400,{'error':str(e)[:400] or 'Unable to process document.'})
if __name__=='__main__':
    server=ThreadingHTTPServer(('127.0.0.1',8765),Handler); print('Report Studio: http://127.0.0.1:8765'); threading.Timer(.4,lambda:webbrowser.open('http://127.0.0.1:8765')).start()
    try: server.serve_forever()
    except KeyboardInterrupt: server.server_close()
