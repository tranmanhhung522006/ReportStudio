"""Compare citation field runs and cached results without executing Word fields."""
from collections import Counter
from io import BytesIO
from zipfile import ZipFile
from lxml import etree
import re, hashlib
from docx.oxml.ns import qn

PATTERN=re.compile(r'ADDIN\s+(?:EN\.|EN\b|ZOTERO|CSL|MENDELEY)',re.I)

def signature(nodes):
    # Preserve field instructions, displayed text, field markers and embedded field data.
    values=[]
    for n in nodes:
        tag=n.tag.rsplit('}',1)[-1]
        if tag in ('instrText','t','delText','fldData','fldChar','fldSimple'):
            values.append((tag,n.text or '',tuple(sorted(n.attrib.items()))))
    return hashlib.sha256(repr(values).encode()).hexdigest()

def snapshot(data):
    found=Counter()
    with ZipFile(BytesIO(data)) as z:
        for name in z.namelist():
            if not name.startswith('word/') or not name.endswith('.xml'): continue
            root=etree.fromstring(z.read(name),parser=etree.XMLParser(resolve_entities=False,no_network=True))
            stack=[]
            for n in root.iter():
                tag=n.tag.rsplit('}',1)[-1]
                if tag=='fldSimple':
                    if PATTERN.search(n.get(qn('w:instr'),'')): found[(name,signature(n.iter()))]+=1
                if tag=='fldChar' and n.get(qn('w:fldCharType'))=='begin':
                    stack.append([])
                for group in stack: group.append(n)
                if tag=='fldChar' and n.get(qn('w:fldCharType'))=='end' and stack:
                    group=stack.pop()
                    instruction=''.join(x.text or '' for x in group if x.tag==qn('w:instrText'))
                    if PATTERN.search(instruction): found[(name,signature(group))]+=1
            for group in stack:
                instruction=''.join(x.text or '' for x in group if x.tag==qn('w:instrText'))
                if PATTERN.search(instruction): found[(name,signature(group))]+=1
    return found

def compare_citations(before,after):
    a,b=snapshot(before),snapshot(after)
    return {'passed':a==b,'before':sum(a.values()),'after':sum(b.values()),'changed_or_missing':sum((a-b).values()),'added_or_changed':sum((b-a).values())}
