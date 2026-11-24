"""Bounded ephemeral uploads. One Gunicorn worker, no original files in Neon."""
import os, time, uuid, tempfile, threading
from pathlib import Path
from collections import OrderedDict

TTL=30*60
LIMIT=45*1024*1024
MAX_ITEMS=12
MAX_OWNER_ITEMS=3
class UploadExpired(ValueError): pass

class UploadCache:
    def __init__(self,root=None,clock=time.monotonic):
        self.temp=tempfile.TemporaryDirectory(prefix='reportstudio-uploads-') if root is None else None
        self.root=Path(root or self.temp.name)
        self.root.mkdir(parents=True,exist_ok=True)
        self.clock=clock
        self.lock=threading.RLock()
        self.items=OrderedDict()
    def _remove(self,key):
        row=self.items.pop(key,None)
        if row:
            try: row[1].unlink(missing_ok=True)
            except OSError: pass
    def cleanup(self):
        with self.lock:
            for key,row in list(self.items.items()):
                if row[3]<=self.clock(): self._remove(key)
    def put(self,owner,data):
        if not 0<len(data)<=15*1024*1024: raise ValueError('Invalid upload size.')
        with self.lock:
            self.cleanup()
            owned=[key for key,row in self.items.items() if row[0]==owner]
            while len(owned)>=MAX_OWNER_ITEMS: self._remove(owned.pop(0))
            while self.items and (len(self.items)>=MAX_ITEMS or sum(row[2] for row in self.items.values())+len(data)>LIMIT):
                self._remove(next(iter(self.items)))
            key=str(uuid.uuid4()); path=self.root/key
            path.write_bytes(data)
            self.items[key]=(owner,path,len(data),self.clock()+TTL)
            return key
    def get(self,owner,key):
        with self.lock:
            self.cleanup()
            row=self.items.get(str(key))
            if not row or row[0]!=owner: raise UploadExpired('Upload expired or unavailable. Analyze the original file again.')
            try: return row[1].read_bytes()
            except OSError:
                self._remove(str(key))
                raise UploadExpired('Upload expired or unavailable. Analyze the original file again.')
    def delete(self,owner,key):
        with self.lock:
            row=self.items.get(str(key))
            if row and row[0]==owner: self._remove(str(key)); return True
            return False
