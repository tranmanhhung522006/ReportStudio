"""Full ordered DOCX preview and strict image relationship verification."""
import base64, hashlib, posixpath, re
from collections import Counter
from io import BytesIO
from zipfile import ZipFile
from lxml import etree
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from engine import load, heading_level

REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
V='urn:schemas-microsoft-com:vml'
O='urn:schemas-microsoft-com:office:office'
SAFE_MIME={'image/png','image/jpeg','image/gif','image/bmp','image/webp'}
IMAGE_LIMIT=16*1024*1024

def local(n): return n.tag.rsplit('}',1)[-1] if isinstance(n.tag,str) else ''
def xml(data): return etree.fromstring(data,parser=etree.XMLParser(resolve_entities=False,no_network=True))
def image_ids(n):
    # Includes DrawingML, VML and modern SVG fallback references.
    for key,val in n.attrib.items():
        ns,_,tag=key[1:].partition('}') if key.startswith('{') else ('','',key)
        if (ns==REL and tag in ('embed','link','id') and local(n) in ('blip','svgBlip','imagedata','fill','image')) or (ns==O and tag=='relid'):
            yield val

def image_snapshot(data):
    result=Counter()
    with ZipFile(BytesIO(data)) as z:
        names=set(z.namelist())
        for part in sorted(names):
            if not part.startswith('word/') or not part.endswith('.xml') or '/_rels/' in part: continue
            root=xml(z.read(part))
            refs=[rid for n in root.iter() for rid in image_ids(n)]
            if not refs: continue
            relpath=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
            rels={r.get('Id'):r for r in xml(z.read(relpath))} if relpath in names else {}
            for rid in refs:
                r=rels.get(rid)
                if r is None: result[(part,'missing',rid)]+=1;continue
                target=r.get('Target','')
                if r.get('TargetMode')=='External':
                    result[(part,'external',target)]+=1;continue
                path=posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join(posixpath.dirname(part),target))
                if path not in names: result[(part,'missing',path)]+=1;continue
                result[(part,'embedded',hashlib.sha256(z.read(path)).hexdigest())]+=1
    return result

def compare_images(before,after):
    a,b=image_snapshot(before),image_snapshot(after)
    missing=a-b; added=b-a
    return {'passed':a==b,'before':sum(a.values()),'after':sum(b.values()),
            'missing_or_changed':sum(missing.values()),'added_or_changed':sum(added.values()),
            'external':sum(v for k,v in a.items() if k[1]=='external'),
            'unresolved_original':sum(v for k,v in a.items() if k[1]=='missing')}

def full_preview(data):
    """No 300-paragraph truncation. Return body blocks, not untrusted HTML."""
    d=load(data); assets={}; used=0; warnings=set(); image_count=0
    def image_ref(rid,owner,holder):
        nonlocal used,image_count
        image_count+=1
        try:
            rel=owner.rels[rid]
            if rel.is_external:
                warnings.add('Linked external images are not fetched.')
                return {'kind':'image','missing':'External image (not fetched)'}
            part=rel.target_part; mime=part.content_type; raw=part.blob
            if mime not in SAFE_MIME:
                warnings.add('Some vector or legacy images need Word/PDF for display.')
                return {'kind':'image','missing':'Image format requires Word/PDF'}
            key=hashlib.sha256(raw).hexdigest()
            if key not in assets:
                if used+len(raw)>IMAGE_LIMIT:
                    warnings.add('Image preview byte limit reached; remaining images are preserved in DOCX.')
                    return {'kind':'image','missing':'Preview image limit reached'}
                assets[key]={'mime':mime,'base64':base64.b64encode(raw).decode('ascii')}
                used+=len(raw)
            extent=next(holder.iter(qn('wp:extent')),None)
            width=None
            if extent is not None:
                try: width=max(1,min(2000,int(extent.get('cx','0'))/9525))
                except ValueError: pass
            return {'kind':'image','asset':key,'width':width}
        except (KeyError,ValueError,AttributeError):
            warnings.add('Some original image links could not be resolved.')
            return {'kind':'image','missing':'Unresolved image reference'}
    def value(n,owner):
        kind=local(n)
        if kind in ('drawing','pict'):
            # A drawing can have a primary image and fallback; render the first supported one.
            refs=list(dict.fromkeys(rid for x in n.iter() for rid in image_ids(x)))
            if not refs:
                warnings.add('Some drawings or charts require Word/PDF.')
                return [{'kind':'image','missing':'Drawing/chart requires Word/PDF'}]
            for rid in refs:
                r=image_ref(rid,owner,n)
                if 'asset' in r: return [r]
            return [r]
        if kind=='AlternateContent':
            fallback=next((x for x in n if local(x)=='Fallback'),None)
            choice=fallback if fallback is not None else next(iter(n),None)
            return value(choice,owner) if choice is not None else []
        if kind in ('t','delText') and kind!='delText': return [{'kind':'text','text':n.text or ''}]
        if kind=='tab': return [{'kind':'text','text':'\t'}]
        if kind in ('br','cr'):
            return [{'kind':'break'}] if n.get(qn('w:type'))=='page' else [{'kind':'text','text':'\n'}]
        if kind in ('pPr','rPr','instrText','fldChar','del','sectPr','lastRenderedPageBreak'): return []
        if kind=='r':
            rp=n.find(qn('w:rPr')); items=[]
            for x in n:
                for t in value(x,owner):
                    if t['kind']=='text' and rp is not None:
                        for tag,key in [('b','bold'),('i','italic'),('u','underline')]:
                            setting=rp.find(qn('w:'+tag))
                            if setting is not None: t[key]=setting.get(qn('w:val'),'1') not in ('0','false','none','off')
                        size=rp.find(qn('w:sz'))
                        if size is not None:
                            try: t['size']=float(size.get(qn('w:val')))/2
                            except (ValueError,TypeError): pass
                        font=rp.find(qn('w:rFonts'))
                        if font is not None: t['font']=font.get(qn('w:ascii')) or font.get(qn('w:hAnsi'))
                    items.append(t)
            return items
        return [t for x in n for t in value(x,owner)]
    def blocks(parent,owner):
        out=[]
        for n in parent:
            tag=local(n)
            if tag=='p':
                p=Paragraph(n,d._body); st=p.style
                f=p.paragraph_format
                align=f.alignment if f.alignment is not None else st.paragraph_format.alignment if st is not None else None
                line=f.line_spacing if f.line_spacing is not None else st.paragraph_format.line_spacing if st is not None else None
                b={'kind':'paragraph','level':heading_level(p),'segments':value(n,owner),
                   'alignment':{0:'left',1:'center',2:'right',3:'justify'}.get(align,'left'),
                   'page_break':bool(f.page_break_before)}
                if isinstance(line,float): b['line']=line
                if st is not None:
                    if st.font.name: b['font']=st.font.name
                    if st.font.size: b['size']=st.font.size.pt
                out.append(b)
            elif tag=='tbl':
                rows=[]
                for tr in n:
                    if local(tr)!='tr': continue
                    cells=[]
                    for tc in tr:
                        if local(tc)!='tc': continue
                        span=tc.find('./'+qn('w:tcPr')+'/'+qn('w:gridSpan'))
                        cells.append({'span':int(span.get(qn('w:val'),'1')) if span is not None else 1,'blocks':blocks(tc,owner)})
                    rows.append(cells)
                out.append({'kind':'table','rows':rows})
            elif tag in ('sdt','sdtContent','customXml','ins'):
                out.extend(blocks(n,owner))
        return out
    result=blocks(d.element.body,d.part)
    section=d.sections[0]
    dims={k:round(getattr(section,k).cm,2) for k in ('page_width','page_height','top_margin','bottom_margin','left_margin','right_margin') if getattr(section,k) is not None}
    return {'blocks':result,'assets':assets,'image_count':image_count,'page':dims,'warnings':sorted(warnings),
            'note':'Full document body preview; actual Word pagination, floating positions, headers/footers and complex objects may differ.'}
