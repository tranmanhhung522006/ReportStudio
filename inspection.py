"""Read-only document inspection, approximate preview and sample style extraction."""
from collections import Counter
import re
from docx.oxml.ns import qn
from engine import load, heading_level, protected, recognize, parse_outline

def most(values):
    values=[v for v in values if v is not None]
    return Counter(values).most_common(1)[0][0] if values else None

def inherited(p, attr):
    value=getattr(p.paragraph_format,attr)
    style=p.style
    while value is None and style is not None:
        value=getattr(style.paragraph_format,attr)
        style=style.base_style
    return value

def font_value(run,p,attr):
    value=getattr(run.font,attr)
    style=p.style
    while value is None and style is not None:
        value=getattr(style.font,attr)
        style=style.base_style
    return value

def extract_sample(data):
    from io import BytesIO
    from zipfile import ZipFile, ZIP_DEFLATED
    from engine import MAX_BYTES, MAX_EXPANDED
    if len(data)>MAX_BYTES: raise ValueError('Template too large.')
    with ZipFile(BytesIO(data)) as z:
        if len(z.infolist())>3000 or sum(x.file_size for x in z.infolist())>MAX_EXPANDED or any('vbaproject' in x.filename.lower() for x in z.infolist()): raise ValueError('Unsafe or macro-enabled template.')
        manifest=z.read('[Content_Types].xml')
        old=b'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'
        if old in manifest:
            out=BytesIO()
            with ZipFile(out,'w',ZIP_DEFLATED) as dest:
                for item in z.infolist(): dest.writestr(item.filename,manifest.replace(old,b'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml') if item.filename=='[Content_Types].xml' else z.read(item))
            data=out.getvalue()
    d=load(data); rules={}; evidence={}; warnings=[]
    def put(k,v,why):
        if v is not None: rules[k]=v; evidence[k]=why
    body=[p for p in d.paragraphs if p.text.strip() and not heading_level(p) and p.style.name in ('Normal','Body Text','No Spacing')]
    put('font',most(font_value(r,p,'name') for p in body for r in p.runs if r.text.strip()),'Most common body font')
    for level,key in [(0,'size'),(1,'h1'),(2,'h2'),(3,'h3')]:
        paras=body if not level else [p for p in d.paragraphs if heading_level(p)==level]
        sizes=[font_value(r,p,'size') for p in paras for r in p.runs if r.text.strip()]
        v=most(x.pt for x in sizes if x is not None)
        put(key,v,'Most common visible font size in sample')
    spacing=most(inherited(p,'line_spacing') for p in body)
    if isinstance(spacing,float) and 1<=spacing<=3: put('line',spacing,'Most common body line spacing')
    for attr,key in [('space_before','before'),('space_after','after'),('first_line_indent','first_indent')]:
        v=most(inherited(p,attr) for p in body)
        if v is not None:
            val=round(v.cm if key=='first_indent' else v.pt,2)
            if 0<=val<=(3 if key=='first_indent' else 72): put(key,val,'Most common body paragraph setting')
    align=most(inherited(p,'alignment') for p in body)
    if align in (0,1,3): put('alignment',{0:'left',1:'center',3:'justify'}[int(align)],'Most common body alignment')
    if d.sections:
        sec=d.sections[0]
        for k in ('left','right','top','bottom'):
            v=getattr(sec,k+'_margin')
            if v is not None and .5<=v.cm<=5: put(k,round(v.cm,2),'First section margin')
        if sec.page_width and sec.page_height:
            dims=sorted([sec.page_width.cm,sec.page_height.cm])
            for name,expected in [('A4',(21,29.7)),('A5',(14.8,21)),('Letter',(21.59,27.94))]:
                if all(abs(x-y)<.12 for x,y in zip(dims,expected)): put('paper',name,'First section paper size')
        if len(d.sections)>1: warnings.append('Sample contains multiple sections. Only first-section page layout is suggested.')
    if not body: warnings.append('No standard body paragraphs detected; some settings could not be inferred.')
    warnings.append('Sample extraction suggests common styles only. Review exceptions and heading levels.')
    return dict(rules=rules,evidence=evidence,warnings=warnings,mode='sample')

def preview(data,outline=''):
    d=load(data); mapping=parse_outline(outline); blocks=[]
    # Text-only preview; never return executable HTML or embedded external resources.
    for i,p in enumerate(d.paragraphs[:300]):
        lev,_,_=recognize(p,mapping)
        if p.text.strip():
            blocks.append(dict(index=i,text=p.text[:2500],level=lev,protected=protected(p)))
    return dict(blocks=blocks,truncated=len(d.paragraphs)>300,tables=len(d.tables),images=len(d.inline_shapes),note='Approximate text preview. Tables, images, fields and exact page breaks are not rendered.')

def reference_audit(data):
    d=load(data); refs=[]; seen=set(); duplicates=[]; in_refs=False; dois=set()
    for p in d.paragraphs:
        t=p.text.strip()
        if re.fullmatch(r'(references|bibliography|tài liệu tham khảo)',t,re.I): in_refs=True; continue
        if in_refs and heading_level(p)==1: break
        if in_refs and t:
            k=' '.join(t.casefold().split())
            if k in seen: duplicates.append(t[:180])
            seen.add(k); refs.append(t)
            dois.update(x.rstrip('.,;)') for x in re.findall(r'10\.\d{4,9}/[^\s<>"]+',t,re.I))
    return dict(reference_entries=len(refs),duplicate_entries=duplicates[:30],doi_candidates=sorted(dois)[:100],note='DOI patterns are detected locally, not verified online. Citation completeness and style require manual review.')
