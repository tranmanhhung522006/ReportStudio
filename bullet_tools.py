"""Opt-in native Word bullet formatting; never change numbered headings."""
from copy import deepcopy
import re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

SYMBOLS={'disc':'•','circle':'○','solid_circle':'●','small_circle':'◦','square':'■','small_square':'▪',
         'outline_square':'□','diamond':'◆','outline_diamond':'◇','arrow':'➜','triangle':'▸',
         'dash':'–','check':'✓','star':'★','plus':'+','cross':'✕','chevron':'›','dot':'·'}
DEFAULTS={'bullet_level1':'preserve','bullet_level2':'preserve','bullet_level3':'preserve',
          'bullet_convert_manual':False,'bullet_align':False,'bullet_indent':0.75,'bullet_gap':0.4}
MANUAL=re.compile(r'^\s*([•●○◦■▪□◆◇➜▸✓★✕›·])\s+')

def el(tag,**attrs):
    x=OxmlElement('w:'+tag)
    for k,v in attrs.items():x.set(qn('w:'+k),str(v))
    return x

def child(node,tag):
    return node.find(qn('w:'+tag)) if node is not None else None

def val(node,tag,default=None):
    x=child(node,tag)
    return x.get(qn('w:val'),default) if x is not None else default

def set_child(node,tag,value):
    x=child(node,tag)
    if x is None:x=el(tag);node.append(x)
    x.set(qn('w:val'),str(value))
    return x

def list_ref(p):
    """Return effective numbering id and level, including inherited list styles."""
    props=[p._p.pPr];st=p.style;seen=set()
    while st is not None and st.style_id not in seen:
        seen.add(st.style_id);props.append(st._element.pPr);st=st.base_style
    n=lev=None
    for pp in props:
        np=child(pp,'numPr')
        if np is None:continue
        if n is None:n=val(np,'numId')
        if lev is None:lev=val(np,'ilvl')
    try:return (int(n),max(0,min(8,int(lev or 0)))) if n and int(n)>0 else (0,0)
    except (ValueError,TypeError):return (0,0)

def is_list(p):
    return bool(list_ref(p)[0] or MANUAL.match(p.text) or (p.style and p.style.name.lower().startswith(('list bullet','list number'))))

def validate(c):
    for key in ('bullet_level1','bullet_level2','bullet_level3'):
        if c[key] not in ('preserve',*SYMBOLS):raise ValueError('Invalid '+key)
    for key,lo,hi in [('bullet_indent',0.25,2.5),('bullet_gap',0.15,1.5)]:
        c[key]=float(c[key])
        if not lo<=c[key]<=hi:raise ValueError('Invalid '+key)

def format_lists(d,c,skip):
    from engine import heading_level,protected,CAPTION
    report={'native_bullets':0,'manual_bullets':0,'bullet_lists':0}
    if not (c['bullet_convert_manual'] or c['bullet_align'] or any(c[k]!='preserve' for k in ('bullet_level1','bullet_level2','bullet_level3'))):return report
    root=d.part.numbering_part.element
    nums={int(x.get(qn('w:numId'))):x for x in root.findall(qn('w:num'))}
    abstracts={int(x.get(qn('w:abstractNumId'))):x for x in root.findall(qn('w:abstractNum'))}
    next_a=max(abstracts,default=-1)+1;next_n=max(nums,default=0)+1
    clones={}
    def effective_level(num,abstract,lev):
        for o in num.findall(qn('w:lvlOverride')):
            if o.get(qn('w:ilvl'))==str(lev) and child(o,'lvl') is not None:return child(o,'lvl')
        return next((x for x in abstract.findall(qn('w:lvl')) if x.get(qn('w:ilvl'))==str(lev)),None)
    def configure(lvl,lev,default=None):
        choice=c['bullet_level'+str(min(3,lev+1))]
        glyph=SYMBOLS.get(choice,default)
        if glyph:
            set_child(lvl,'lvlText',glyph)
            rp=child(lvl,'rPr')
            if rp is None:rp=el('rPr');lvl.append(rp)
            fonts=child(rp,'rFonts')
            if fonts is None:fonts=el('rFonts');rp.insert(0,fonts)
            fonts.attrib.clear()
            for k in ('ascii','hAnsi','eastAsia','cs'):fonts.set(qn('w:'+k),'Arial')
        if c['bullet_align']:
            pp=child(lvl,'pPr')
            if pp is None:
                pp=el('pPr');rp=child(lvl,'rPr')
                if rp is not None:rp.addprevious(pp)
                else:lvl.append(pp)
            ind=child(pp,'ind')
            if ind is None:ind=el('ind');pp.append(ind)
            left=Cm(c['bullet_indent']*(lev+1)).twips
            gap=min(left,Cm(c['bullet_gap']).twips)
            for k in ('firstLine','firstLineChars','leftChars','hangingChars','start','startChars'):ind.attrib.pop(qn('w:'+k),None)
            ind.set(qn('w:left'),str(left));ind.set(qn('w:hanging'),str(gap))
    def install(abstract,num):
        nonlocal next_a,next_n
        abstract.set(qn('w:abstractNumId'),str(next_a));num.set(qn('w:numId'),str(next_n))
        set_child(num,'abstractNumId',next_a)
        first=root.find(qn('w:num'))
        if first is not None:first.addprevious(abstract)
        else:root.append(abstract)
        root.append(num);n=next_n;next_a+=1;next_n+=1;report['bullet_lists']+=1;return n
    def assign(p,n,lev):
        np=p._p.get_or_add_pPr().get_or_add_numPr();np.get_or_add_numId().val=n;np.get_or_add_ilvl().val=lev
        if c['bullet_align']:
            p.paragraph_format.left_indent=Cm(c['bullet_indent']*(lev+1))
            p.paragraph_format.first_line_indent=-Cm(min(c['bullet_indent']*(lev+1),c['bullet_gap']))
    manual_num=None;last_manual=False
    for p in d.paragraphs:
        if p._p in skip or protected(p) or heading_level(p) or CAPTION.match(p.text):
            last_manual=False;continue
        n,lev=list_ref(p)
        if n:
            last_manual=False
            num=nums.get(n)
            if num is None:continue
            try:abstract=abstracts.get(int(val(num,'abstractNumId')))
            except (TypeError,ValueError):continue
            if abstract is None:continue
            lvl=effective_level(num,abstract,lev)
            if val(lvl,'numFmt')!='bullet':continue
            if c['bullet_level'+str(min(3,lev+1))]=='preserve' and not c['bullet_align']:continue
            if n not in clones:
                a=deepcopy(abstract);b=deepcopy(num)
                for x in a.findall(qn('w:lvl')):
                    if val(x,'numFmt')=='bullet':configure(x,int(x.get(qn('w:ilvl'))))
                for o in b.findall(qn('w:lvlOverride')):
                    x=child(o,'lvl')
                    if val(x,'numFmt')=='bullet':configure(x,int(o.get(qn('w:ilvl'))))
                clones[n]=install(a,b)
            assign(p,clones[n],lev);report['native_bullets']+=1
            continue
        m=MANUAL.match(p.text)
        if not c['bullet_convert_manual'] or not m or not p.text[m.end():].strip():
            last_manual=False;continue
        if not all(all(x.tag in (qn('w:rPr'),qn('w:t')) for x in r._r) for r in p.runs) or ''.join(r.text for r in p.runs)!=p.text:
            last_manual=False;continue
        # A manual marker contains no reliable nesting metadata; convert as level one.
        # Preserve each source glyph when no replacement is requested.
        key=SYMBOLS.get(c['bullet_level1'],m.group(1))
        if not last_manual or manual_num is None or manual_glyph!=key:
            a=el('abstractNum');a.append(el('multiLevelType',val='multilevel'))
            for i in range(9):
                lvl=el('lvl',ilvl=i)
                for tag,v in [('start',1),('numFmt','bullet'),('suff','space'),('lvlText',key),('lvlJc','left')]:lvl.append(el(tag,val=v))
                configure(lvl,i,key)
                a.append(lvl)
            manual_num=install(a,el('num'));manual_glyph=key
        left=m.end()
        for r in p.runs:
            if left<=0:break
            t=r.text;cut=min(left,len(t));r.text=t[cut:];left-=cut
        assign(p,manual_num,0);report['manual_bullets']+=1;last_manual=True
    return report
