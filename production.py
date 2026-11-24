"""Production WSGI application for Render; local server.py remains supported."""
import os, uuid, base64, threading, time, hmac, hashlib, logging
from collections import OrderedDict
from pathlib import Path
from flask import Flask, request, session, jsonify, send_file, send_from_directory, Response
from io import BytesIO
from werkzeug.exceptions import HTTPException
from engine import analyze, MAX_BYTES
from advanced import advanced, audit, pdf_export
from requirement_parser import extract_text, parse_rules, ai_rules
import storage
from upload_cache import UploadCache, UploadExpired
from doi_check import check_many
uploads=UploadCache()
from inspection import extract_sample, preview, reference_audit
from zipfile import BadZipFile
from docx.opc.exceptions import PackageNotFoundError

ROOT=Path(__file__).resolve().parent/'web'
app=Flask(__name__,static_folder=None)
secret=os.environ.get('SECRET_KEY')
if not secret or len(secret)<32: raise RuntimeError('Set SECRET_KEY to a random secret of at least 32 characters.')
app.config.update(SECRET_KEY=secret,MAX_CONTENT_LENGTH=MAX_BYTES*2,SESSION_COOKIE_HTTPONLY=True,
                  SESSION_COOKIE_SECURE=os.environ.get('RENDER')=='true' or os.environ.get('SECURE_COOKIE')=='1',
                  SESSION_COOKIE_SAMESITE='Strict',PERMANENT_SESSION_LIFETIME=86400*7, MAX_FORM_MEMORY_SIZE=512*1024, MAX_FORM_PARTS=8)
log=logging.getLogger(__name__)
PDF_ENABLED=os.getenv('ENABLE_PDF','0')=='1'
AI_ENABLED=os.getenv('ENABLE_AI','0')=='1'
processing=threading.BoundedSemaphore(1)
rate_lock=threading.Lock()
rates=OrderedDict()
schema_lock=threading.Lock()
schema_ready=False

def database():
    global schema_ready
    if not storage.configured(): return False
    with schema_lock:
        if not schema_ready:
            storage.initialize()
            schema_ready=True
    return True

def owner():
    if 'owner' not in session:
        session['owner']=str(uuid.uuid4())
        session.permanent=True
    return uuid.UUID(session['owner'])

def allowed_rate(key,limit,window=60):
    now=time.monotonic()
    with rate_lock:
        bucket=rates.get(key,[])
        bucket=[t for t in bucket if now-t<window]
        if len(bucket)>=limit: return False
        rates[key]=bucket+[now];rates.move_to_end(key)
        while len(rates)>4096: rates.popitem(last=False)
        return True

def error(message,code=400): return jsonify(error=message),code

@app.before_request
def guard():
    access=os.getenv('PERSONAL_ACCESS_CODE','')
    if access and request.path not in ('/health','/api/cleanup'):
        auth=request.authorization
        if not auth or auth.type!='basic' or not hmac.compare_digest(auth.username or '', 'studio') or not hmac.compare_digest((auth.password or '').encode(),access.encode()):
            return Response('Personal workspace / Không gian cá nhân',401,{'WWW-Authenticate':'Basic realm="ReportStudio", charset="UTF-8"'})
    if request.path.startswith('/api/') and request.method in ('POST','DELETE'):
        # Require a same-origin browser request or a non-browser client with the custom header.
        origin=request.headers.get('Origin')
        expected=os.getenv('PUBLIC_ORIGIN') or os.getenv('RENDER_EXTERNAL_URL')
        if origin and origin.rstrip('/')!=(expected or request.host_url).rstrip('/'):
            return error('Invalid origin.',403)
        if request.headers.get('X-ReportStudio')!='1': return error('Missing request header.',403)
    if request.path.startswith('/api/'):
        # Render's edge forwards the client IP; use the rightmost forwarded value.
        forwarded=request.headers.get('X-Forwarded-For','').split(',')[-1].strip() if os.environ.get('RENDER')=='true' else ''
        ip=forwarded or request.remote_addr or 'unknown'
        if not allowed_rate('ip:'+ip,40): return error('Too many requests. Try again shortly.',429)

@app.after_request
def headers(response):
    response.headers['Cache-Control']='no-store'
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['Referrer-Policy']='no-referrer'
    response.headers['X-Frame-Options']='DENY'
    response.headers['Permissions-Policy']='camera=(), microphone=(), geolocation=()'
    response.headers['Content-Security-Policy']="default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self' data:; frame-src blob:; frame-ancestors 'none'; base-uri 'none'; object-src 'none'"
    return response

@app.errorhandler(HTTPException)
def http_error(e): return error(e.name,e.code)

@app.errorhandler(Exception)
def unexpected(e):
    log.error('Request failed: %s',type(e).__name__)
    return error('Unable to complete the request.',500)

@app.get('/health')
def health(): return jsonify(status='ok')

@app.get('/health/database')
def database_health():
    if not storage.configured(): return jsonify(status='disabled')
    try:
        database()
        with storage.connect() as db: db.execute('SELECT 1')
        return jsonify(status='ok')
    except Exception: return jsonify(status='unavailable'),503

@app.get('/')
def index():
    owner()
    return send_from_directory(ROOT,'index.html')

@app.get('/<any(style.css,app.js,i18n.js):filename>')
def assets(filename): return send_from_directory(ROOT,filename)

@app.get('/api/capabilities')
def capabilities():
    owner()
    return jsonify(history=storage.configured(),pdf=PDF_ENABLED,ai=AI_ENABLED and bool(os.getenv('OPENAI_API_KEY')),online=True,ttl_hours=24,max_saved_bytes=storage.MAX_FILE,multipart=True,inspection=True,upload_sessions=True,doi=True)

@app.get('/api/history')
def history():
    if not database(): return jsonify(files=[])
    return jsonify(files=storage.list_files(owner()))

@app.get('/api/history/<uuid:file_id>')
def download(file_id):
    if not database(): return error('History unavailable.',404)
    row=storage.get_file(owner(),file_id)
    if not row: return error('File not found or expired.',404)
    return send_file(BytesIO(row['content']),mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',as_attachment=True,download_name=row['filename'])

@app.delete('/api/history/<uuid:file_id>')
def remove(file_id):
    if not database() or not storage.delete_file(owner(),file_id): return error('File not found.',404)
    return jsonify(deleted=True)

@app.post('/api/cleanup')
def cleanup_route():
    token=os.getenv('CLEANUP_TOKEN','')
    if not token or not hmac.compare_digest(request.headers.get('Authorization',''),'Bearer '+token):
        return error('Unauthorized.',401)
    return jsonify(deleted=storage.cleanup() if database() else 0)

@app.delete('/api/uploads/<uuid:upload_id>')
def delete_upload(upload_id):
    return jsonify(deleted=uploads.delete(owner(),str(upload_id)))

@app.post('/api/doi')
def verify_doi():
    if not request.is_json: return error('JSON required.',415)
    if not allowed_rate('doi:'+str(owner()),12,60): return error('DOI rate limit reached.',429)
    if not processing.acquire(blocking=False): return error('Server busy. Retry shortly.',429)
    try:
        return jsonify(results=check_many(request.get_json().get('dois',[])))
    except (ValueError,TypeError,AttributeError) as exc: return error(str(exc)[:240])
    finally: processing.release()

@app.post('/api/<any(analyze,format,requirements,audit,pdf,sample,preview,references):action>')
def process(action):
    if not request.is_json and request.mimetype!='multipart/form-data': return error('JSON or multipart required.',415)
    if not processing.acquire(blocking=False): return error('Server busy. Retry shortly.',429)
    try:
        if request.is_json:
            req=request.get_json()
        else:
            import json
            req=json.loads(request.form.get('payload','{}'))
        if not isinstance(req,dict): return error('Invalid request.')
        def doc():
            if req.get('upload_id'): return uploads.get(owner(),req['upload_id'])
            upload=request.files.get('document') if not request.is_json else None
            data=upload.read(MAX_BYTES+1) if upload is not None else base64.b64decode(req['document'],validate=True)
            if len(data)>MAX_BYTES: raise ValueError('Maximum file size is 15 MB.')
            return data
        if action=='requirements':
            source=extract_text(doc(),req['filename']) if req.get('document') or (not request.is_json and 'document' in request.files) else req.get('text','')
            if req.get('ai'):
                if not AI_ENABLED: return error('Cloud AI disabled.',403)
                if not allowed_rate('ai:'+str(owner()),3,86400): return error('Daily AI limit reached.',429)
                return jsonify(ai_rules(source,req.get('consent',False)))
            return jsonify(parse_rules(source))
        if action=='pdf' and not PDF_ENABLED: return error('PDF conversion is disabled on this server.',403)
        data=doc()
        if action=='sample': return jsonify(extract_sample(data))
        if action=='preview': return jsonify(preview(data,req.get('outline','')))
        if action=='references': return jsonify(reference_audit(data))
        if action=='analyze':
            result=analyze(data,req.get('outline','')); result['preview']=preview(data,req.get('outline',''))
            if req.get('cache_upload') is True: result['upload_id']=uploads.put(owner(),data)
            return jsonify(result)
        if action=='audit': return jsonify(audit(data,req.get('config',{}),req.get('outline','')))
        if action=='pdf': return jsonify(pdf=base64.b64encode(pdf_export(data)).decode())
        out,report=advanced(data,req['config'],req['decisions'],req.get('outline',''))
        result={'document':base64.b64encode(out).decode(),'report':report,'saved':None}
        if req.get('save_history') is True and storage.configured():
            try:
                database()
                result['saved']=storage.save(owner(),req.get('filename'),out)
            except storage.StorageFull as exc: result['storage_warning']=str(exc)
            except Exception as exc:
                log.warning('History save failed: %s',type(exc).__name__)
                result['storage_warning']='History unavailable. Download your result now.'
        return jsonify(result)
    except UploadExpired as exc:
        return jsonify(error=str(exc),code='upload_expired'),410
    except (ValueError,KeyError,TypeError,BadZipFile,PackageNotFoundError) as exc:
        return error(str(exc)[:240])
    finally: processing.release()

def cleanup_loop():
    while True:
        try:
            if database(): storage.cleanup()
        except Exception as exc: log.warning('Cleanup failed: %s',type(exc).__name__)
        time.sleep(900)

def upload_cleanup_loop():
    while True:
        time.sleep(60)
        uploads.cleanup()

threading.Thread(target=upload_cleanup_loop,daemon=True,name='upload-cleanup').start()
if storage.configured():
    threading.Thread(target=cleanup_loop,daemon=True,name='history-cleanup').start()
