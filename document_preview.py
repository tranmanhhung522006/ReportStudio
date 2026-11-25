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


HIGHLIGHTS={'yellow':'#ffff00','green':'#00ff00','cyan':'#00ffff','magenta':'#ff00ff','blue':'#0000ff','red':'#ff0000','darkBlue':'#000080','darkCyan':'#008080','darkGreen':'#008000','darkMagenta':'#800080','darkRed':'#800000','darkYellow':'#808000','darkGray':'#808080','lightGray':'#c0c0c0','black':'#000000','white':'#ffffff'}
def child_val(node,tag,default=None):
    child=node.find(qn('w:'+tag)) if node is not None else None
    return child.get(qn('w:val'),default) if child is not None else default

def numeric(value,default=0):
    try: return float(value)
    except (TypeError,ValueError): return default

def style_chain(style):
    seen=set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style=style.base_style

def numbering_preview(d):
    """Render common native numbering without changing the document."""
    try: root=d.part.numbering_part.element
    except (KeyError,NotImplementedError): return lambda p: ''
    abstract={x.get(qn('w:abstractNumId')):x for x in root.findall(qn('w:abstractNum'))}
    nums={x.get(qn('w:numId')):x for x in root.findall(qn('w:num'))}
    counters={}
    def level_for(num,level):
        for override in num.findall(qn('w:lvlOverride')):
            if override.get(qn('w:ilvl'))==str(level):
                embedded=override.find(qn('w:lvl'))
                if embedded is not None:return embedded,child_val(override,'startOverride')
        ab=abstract.get(child_val(num,'abstractNumId'))
        if ab is None:return None,None
        base=next((x for x in ab.findall(qn('w:lvl')) if x.get(qn('w:ilvl'))==str(level)),None)
        override=next((x for x in num.findall(qn('w:lvlOverride')) if x.get(qn('w:ilvl'))==str(level)),None)
        return base,child_val(override,'startOverride')
    def fmt(n,kind):
        if kind in ('upperRoman','lowerRoman'):
            if not 0<n<4000:return str(n)
            out=''
            for k,v in [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]:
                q,n=divmod(n,k);out+=v*q
            return out if kind=='upperRoman' else out.lower()
        if kind in ('upperLetter','lowerLetter') and 0<n<=702:
            out=''
            while n:n,r=divmod(n-1,26);out=chr(65+r)+out
            return out if kind=='upperLetter' else out.lower()
        return str(n)
    def marker(p):
        props=[p._p.pPr]+[st._element.pPr for st in style_chain(p.style)]
        numid=next((child_val(np,'numId') for pp in props if pp is not None for np in [pp.find(qn('w:numPr'))] if np is not None and child_val(np,'numId') is not None),None)
        if not numid or numid=='0' or numid not in nums:return ''
        level=int(numeric(next((child_val(np,'ilvl') for pp in props if pp is not None for np in [pp.find(qn('w:numPr'))] if np is not None and child_val(np,'ilvl') is not None),'0')))
        if not 0<=level<=8:return ''
        num=nums[numid];lv,start=level_for(num,level)
        if lv is None:return ''
        kind=child_val(lv,'numFmt','decimal')
        if kind=='none':return ''
        if kind=='bullet':return '•'
        state=counters.setdefault(numid,{})
        for i in list(state):
            if i>level:state.pop(i)
        for i in range(level+1):
            other,override=level_for(num,i)
            initial=int(numeric(override or child_val(other,'start','1'),1))
            if i not in state:state[i]=initial
            elif i==level:state[i]+=1
        text=child_val(lv,'lvlText','%'+str(level+1)+'.')
        return re.sub(r'%([1-9])',lambda m:fmt(state.get(int(m.group(1))-1,1),child_val(level_for(num,int(m.group(1))-1)[0],'numFmt','decimal')),text)
    return marker

def full_preview(data):
    """No 300-paragraph truncation. Return body blocks, not untrusted HTML."""
    d=load(data); assets={}; used=0; warnings=set(); image_count=0
    marker=numbering_preview(d)
    # Resolve the OOXML color scheme as Word does, including theme tint/shade.
    theme={}
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        tree=xml(d.part.part_related_by(RT.THEME).blob)
        scheme=next((n for n in tree.iter() if local(n)=='clrScheme'),None)
        if scheme is not None:
            for n in scheme:
                v=next(iter(n),None)
                if v is not None:
                    color=v.get('lastClr') or v.get('val')
                    if color and re.fullmatch(r'[0-9a-fA-F]{6}',color):theme[local(n)]=color
    except (KeyError,AttributeError,StopIteration,ValueError):pass
    aliases={'text1':'dk1','text2':'dk2','background1':'lt1','background2':'lt2','hyperlink':'hlink','followedHyperlink':'folHlink','dark1':'dk1','dark2':'dk2','light1':'lt1','light2':'lt2'}
    def resolve_color(rp):
        n=rp.find(qn('w:color')) if rp is not None else None
        if n is None:return None
        token=n.get(qn('w:themeColor'))
        color=theme.get(aliases.get(token,token)) if token else None
        from_theme=color is not None
        color=color or n.get(qn('w:val'))
        if color=='auto':return '#000000'
        if not color or not re.fullmatch(r'[0-9a-fA-F]{6}',color):return None
        rgb=[int(color[i:i+2],16) for i in (0,2,4)]
        for attr in ('themeShade','themeTint') if from_theme else ():
            factor=n.get(qn('w:'+attr))
            if factor and re.fullmatch(r'[0-9a-fA-F]{2}',factor):
                t=int(factor,16)/255
                rgb=[round(v*t if attr=='themeShade' else 255-(255-v)*t) for v in rgb]
        return '#'+''.join(f'{v:02x}' for v in rgb)
    styles_by_id={st.style_id:st for st in d.styles}
    def run_props(rp,paragraph_style):
        cs=styles_by_id.get(child_val(rp,'rStyle'))
        return [rp]+[st._element.rPr for st in style_chain(cs)]+[st._element.rPr for st in style_chain(paragraph_style)]
    def effective_text(props,t):
        for r in props:
            if r is None:continue
            for tag,key in [('b','bold'),('i','italic'),('u','underline')]:
                v=child_val(r,tag,'1') if r.find(qn('w:'+tag)) is not None else None
                if v is not None and key not in t:t[key]=v not in ('0','false','none','off')
            if 'color' not in t:
                color=resolve_color(r)
                if color:t['color']=color
            if 'highlight' not in t:
                h=child_val(r,'highlight')
                if h in HIGHLIGHTS:t['highlight']=HIGHLIGHTS[h]
                elif h=='none':t['highlight']='#ffffff'
            if 'size' not in t:
                size=child_val(r,'sz')
                if size is not None:
                    try:t['size']=float(size)/2
                    except (ValueError,TypeError):pass
            if not t.get('font'):
                font=r.find(qn('w:rFonts'))
                if font is not None:
                    name=font.get(qn('w:ascii')) or font.get(qn('w:hAnsi'))
                    if name:t['font']=name
        return t
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
    def value(n,owner,paragraph_style=None):
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
            return value(choice,owner,paragraph_style) if choice is not None else []
        if kind in ('t','delText') and kind!='delText': return [{'kind':'text','text':n.text or ''}]
        if kind=='tab': return [{'kind':'text','text':'\t'}]
        if kind in ('br','cr'):
            return [{'kind':'break'}] if n.get(qn('w:type'))=='page' else [{'kind':'text','text':'\n'}]
        if kind in ('pPr','rPr','instrText','fldChar','del','sectPr','lastRenderedPageBreak'): return []
        if kind=='r':
            rp=n.find(qn('w:rPr')); items=[]
            for x in n:
                for t in value(x,owner,paragraph_style):
                    if t['kind']=='text':effective_text(run_props(rp,paragraph_style),t)
                    items.append(t)
            return items
        return [t for x in n for t in value(x,owner,paragraph_style)]
    def blocks(parent,owner):
        out=[]
        for n in parent:
            tag=local(n)
            if tag=='p':
                p=Paragraph(n,d._body); st=p.style
                f=p.paragraph_format
                align=f.alignment if f.alignment is not None else st.paragraph_format.alignment if st is not None else None
                line=f.line_spacing if f.line_spacing is not None else st.paragraph_format.line_spacing if st is not None else None
                b={'kind':'paragraph','level':heading_level(p),'segments':value(n,owner,st),
                   'alignment':{0:'left',1:'center',2:'right',3:'justify'}.get(align,'left'),
                   'page_break':bool(f.page_break_before)}
                b['number']=marker(p)
                props=[p._p.pPr]+[style._element.pPr for style in style_chain(st)]
                for key,tag,attr,scale in [('left_indent','ind','left',20),('right_indent','ind','right',20),('first_indent','ind','firstLine',20),('before','spacing','before',20),('after','spacing','after',20)]:
                    for pp in props:
                        element=pp.find(qn('w:'+tag)) if pp is not None else None
                        if element is None:continue
                        raw=element.get(qn('w:'+attr))
                        if key=='first_indent' and raw is None and element.get(qn('w:hanging')) is not None:raw=-numeric(element.get(qn('w:hanging')))
                        if raw is not None:b[key]=numeric(raw)/scale;break
                if isinstance(line,float): b['line']=line
                for style in style_chain(st):
                    if 'font' not in b and style.font.name:b['font']=style.font.name
                    if 'size' not in b and style.font.size:b['size']=style.font.size.pt
                    if 'bold' not in b and style.font.bold is not None:b['bold']=style.font.bold
                effective_text([style._element.rPr for style in style_chain(st)],b)
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
