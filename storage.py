"""Neon/PostgreSQL short-lived document storage. No original uploads are stored."""
import os
import uuid
import re
import psycopg
from psycopg.rows import dict_row

TTL_HOURS = 24
MAX_FILE = 5 * 1024 * 1024
MAX_OWNER = 20 * 1024 * 1024
MAX_GLOBAL = 100 * 1024 * 1024
MAX_COUNT = 5

SCHEMA = """
CREATE TABLE IF NOT EXISTS reportstudio_exports (
 id uuid PRIMARY KEY,
 owner_id uuid NOT NULL,
 filename text NOT NULL,
 content bytea NOT NULL,
 size_bytes integer NOT NULL CHECK (size_bytes > 0 AND size_bytes <= 5242880),
 created_at timestamptz NOT NULL DEFAULT now(),
 expires_at timestamptz NOT NULL DEFAULT (now() + interval '24 hours')
);
CREATE INDEX IF NOT EXISTS reportstudio_exports_expiry ON reportstudio_exports(expires_at);
CREATE INDEX IF NOT EXISTS reportstudio_exports_owner ON reportstudio_exports(owner_id, created_at DESC);
"""

class StorageFull(ValueError): pass

def configured():
    return bool(os.environ.get('DATABASE_URL'))

def connect():
    return psycopg.connect(os.environ['DATABASE_URL'], connect_timeout=10,
                          options='-c statement_timeout=15000', row_factory=dict_row,
                          prepare_threshold=None)

def initialize():
    with connect() as db:
        db.execute('SELECT pg_advisory_xact_lock(182736451)')
        for statement in SCHEMA.split(';'):
            if statement.strip(): db.execute(statement)

def cleanup():
    with connect() as db:
        return db.execute('DELETE FROM reportstudio_exports WHERE expires_at <= now()').rowcount

def safe_filename(name):
    name = str(name or 'report').replace('\\','/').rsplit('/',1)[-1]
    name = re.sub(r'[\x00-\x1f\x7f<>:"|?*]', '_', name).strip(' .')[:120]
    name = re.sub(r'\.docx$', '', name, flags=re.I)
    return (name or 'report') + ('_formatted' if not name.endswith('_formatted') else '') + '.docx'

def meta(row):
    return dict(id=str(row['id']), filename=row['filename'], size_bytes=row['size_bytes'],
                created_at=row['created_at'].isoformat(), expires_at=row['expires_at'].isoformat())

def save(owner, filename, data):
    if not 0 < len(data) <= MAX_FILE: raise StorageFull('Result is too large to keep online (5 MB). Download it now.')
    with connect() as db:
        # Serialize quota checks across all app instances, including overlapping deploys.
        db.execute('SELECT pg_advisory_xact_lock(182736451)')
        db.execute('DELETE FROM reportstudio_exports WHERE expires_at <= now()')
        limits=db.execute('SELECT count(*) AS n, coalesce(sum(size_bytes),0) AS size FROM reportstudio_exports WHERE owner_id=%s',(owner,)).fetchone()
        total=db.execute('SELECT coalesce(sum(size_bytes),0) AS size FROM reportstudio_exports').fetchone()['size']
        if limits['n']>=MAX_COUNT or limits['size']+len(data)>MAX_OWNER:
            raise StorageFull('Your history is full. Delete a saved file or download this result now.')
        if total+len(data)>MAX_GLOBAL: raise StorageFull('Temporary storage is full. Download this result now.')
        row=db.execute('INSERT INTO reportstudio_exports(id,owner_id,filename,content,size_bytes) VALUES(%s,%s,%s,%s,%s) RETURNING id,filename,size_bytes,created_at,expires_at',
                       (uuid.uuid4(),owner,safe_filename(filename),data,len(data))).fetchone()
        return meta(row)

def list_files(owner):
    with connect() as db:
        db.execute('DELETE FROM reportstudio_exports WHERE expires_at <= now()')
        rows=db.execute('SELECT id,filename,size_bytes,created_at,expires_at FROM reportstudio_exports WHERE owner_id=%s AND expires_at>now() ORDER BY created_at DESC LIMIT %s',(owner,MAX_COUNT)).fetchall()
        return [meta(row) for row in rows]

def get_file(owner, file_id):
    with connect() as db:
        return db.execute('SELECT filename,content FROM reportstudio_exports WHERE id=%s AND owner_id=%s AND expires_at>now()',(file_id,owner)).fetchone()

def delete_file(owner,file_id):
    with connect() as db:
        return db.execute('DELETE FROM reportstudio_exports WHERE id=%s AND owner_id=%s',(file_id,owner)).rowcount == 1
