"""Bounded online DOI checks using Crossref and DataCite metadata APIs."""
import json,re,time,threading
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote
from urllib.request import Request,urlopen
from urllib.error import HTTPError,URLError

DOI=re.compile(r'^10\.\d{4,9}/[^\s<>"]{1,240}$',re.I)
CACHE=OrderedDict()
LOCK=threading.Lock()
MAX_DOIS=10

def fetch(url):
    req=Request(url,headers={'Accept':'application/json','User-Agent':'ReportStudio/2.1 (DOI metadata check)'})
    try:
        with urlopen(req,timeout=4) as r:
            return 'ok',json.loads(r.read(512*1024))
    except HTTPError as e:
        return ('missing' if e.code==404 else 'unavailable'),None
    except (OSError,ValueError,URLError):
        return 'unavailable',None

def check_one(doi):
    key=doi.casefold()
    with LOCK:
        cached=CACHE.get(key)
        if cached and time.monotonic()-cached[0]<3600: return cached[1]
    encoded=quote(doi,safe='')
    c,data=fetch('https://api.crossref.org/works/'+encoded)
    result={'doi':doi,'status':'unavailable','source':None,'title':None,'year':None}
    if c=='ok' and isinstance(data,dict):
        m=data.get('message',{})
        if isinstance(m,dict) and str(m.get('DOI','')).casefold()==key:
            title=m.get('title') or []
            result.update(status='verified',source='Crossref',title=str(title[0])[:400] if title else None)
            dates=(m.get('published') or {}).get('date-parts') or []
            if dates and dates[0]: result['year']=dates[0][0]
    if result['status']!='verified':
        d,data=fetch('https://api.datacite.org/dois/'+encoded)
        if d=='ok' and isinstance(data,dict):
            m=(data.get('data') or {}).get('attributes') or {}
            if str(m.get('doi','')).casefold()==key:
                titles=m.get('titles') or []
                result.update(status='verified',source='DataCite',title=str(titles[0].get('title',''))[:400] if titles else None,year=m.get('publicationYear'))
        elif c=='missing' and d=='missing':
            result['status']='not_found_in_registries'
    with LOCK:
        CACHE[key]=(time.monotonic(),result); CACHE.move_to_end(key)
        while len(CACHE)>512: CACHE.popitem(last=False)
    return result

def check_many(values):
    if not isinstance(values,list) or len(values)>MAX_DOIS: raise ValueError('Check up to 10 DOIs per request.')
    clean=list(dict.fromkeys(x.strip() for x in values if isinstance(x,str)))
    if any(not DOI.fullmatch(x) for x in clean): raise ValueError('Invalid DOI syntax.')
    with ThreadPoolExecutor(max_workers=4) as pool:
        return list(pool.map(check_one,clean))
