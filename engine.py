"""DOCX formatter: local, reviewable, and non-destructive to the input."""
from io import BytesIO
import re
import unicodedata
from zipfile import ZipFile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

MAX_BYTES = 15 * 1024 * 1024
MAX_EXPANDED = 50 * 1024 * 1024
MAX_PARAGRAPHS = 12000
DEFAULT = dict(font='Times New Roman',size=13,line=1.5,left=3,right=2,top=2,bottom=2,h1=16,h2=14,h3=13,numbering=True,toc=True,tables=True,figures=True,chapter_break=False)
PREFIX = re.compile(r'^\s*(\d+(?:\.\d+){0,8})[.)]?\s+(?=\D)')
CHAPTER = re.compile(r'^\s*(?:chapter|chương)\s+(\d+|[IVXLCDM]+)\s*[:.\-–]?\s+',re.I)
UNNUMBERED = set('abstract|acknowledgements|acknowledgments|references|bibliography|appendix|appendices|tóm tắt|lời cảm ơn|tài liệu tham khảo|phụ lục|table of contents|mục lục|list of figures|list of tables|danh mục hình|danh mục bảng'.split('|'))
MAJOR = set('introduction|literature review|research methodology|methodology|results|discussion|results and discussion|conclusion|conclusions|recommendations|system design|implementation|testing and evaluation|requirement analysis|giới thiệu|tổng quan|phương pháp nghiên cứu|kết quả|thảo luận|kết luận'.split('|'))
MINOR = set('research background|research objectives|research questions|theoretical framework|hypothesis development|data collection|sampling method|research design|data analysis|research gap|scope of study|problem statement|significance of study|bối cảnh nghiên cứu|mục tiêu nghiên cứu|câu hỏi nghiên cứu|phương pháp thu thập dữ liệu'.split('|'))
CAPTION = re.compile(r'^\s*(?:figure|table|hình|bảng)\s+\d+(?:[.\-]\d+)*\s*[:.\-–]\s+\S',re.I)
PROTECTED = set('fldChar fldSimple instrText sdt ins del oMath oMathPara hyperlink footnoteReference endnoteReference drawing pict object bookmarkStart bookmarkEnd commentRangeStart commentRangeEnd'.split())

def load(data):
    if len(data)>MAX_BYTES: raise ValueError('Maximum document size: 15 MB.')
    with ZipFile(BytesIO(data)) as z:
        if len(z.infolist())>3000 or sum(x.file_size for x in z.infolist()) > MAX_EXPANDED: raise ValueError('Expanded document exceeds safe processing limits (50 MB / 3000 parts).')
        if 'word/document.xml' not in z.namelist() or any(x.flag_bits & 1 for x in z.infolist()): raise ValueError('A non-encrypted .docx file is required.')
        if any('vbaproject' in x.filename.lower() for x in z.infolist()): raise ValueError('Macros are not supported.')
    d=Document(BytesIO(data))
    if len(d.paragraphs)>MAX_PARAGRAPHS: raise ValueError('Document contains too many paragraphs (maximum 12000).')
    return d

def norm(s): return unicodedata.normalize('NFKC',s).strip().casefold().rstrip(':')
def has(p,tags): return any(x.tag.rsplit('}',1)[-1] in tags for x in p._p.iter())
def protected(p): return has(p,PROTECTED)
def prefix(s): return CHAPTER.match(s) or PREFIX.match(s)
def unnumbered(s):
    n=norm(s)
    return n in UNNUMBERED or bool(re.match(r'^(appendix|phụ lục)\s+[a-z0-9]+(?:\b|:)',n))

def heading_level(p):
    style=p.style
    if style:
        m=re.fullmatch(r'Heading\s*([1-9])',style.name or '',re.I) or re.fullmatch(r'Heading([1-9])',style.style_id or '',re.I)
        if m: return int(m.group(1))
    pp=p._p.pPr
    if pp is not None:
        ol=pp.find(qn('w:outlineLvl'))
        if ol is not None:
            val=int(ol.get(qn('w:val'),'9'))
            if 0<=val<9: return val+1
    return 0

def recognize(p,outline):
    s=p.text.strip(); level=heading_level(p)
    if level: return level,'Existing Word heading','high'
    if not s or len(s)>200 or CAPTION.match(s) or protected(p): return 0,'',''
    key=norm(s)
    m=prefix(s)
    if key in outline or (m and norm(s[m.end():]) in outline): return outline.get(key,outline.get(norm(s[m.end():]) if m else key)),'Provided outline','high'
    if (p._p.pPr is not None and p._p.pPr.numPr is not None) or (p.style and 'list' in p.style.name.lower()): return 0,'',''
    if unnumbered(s): return 1,'Unnumbered section','high'
    if CHAPTER.match(s): return 1,'Chapter prefix','high'
    m=PREFIX.match(s)
    if m and len(s.split())<=12 and not s.endswith(('.', '?', '!')) and not re.search(r'[;:]\s+\S',s[m.end():]):
        return min(9,len(m.group(1).split('.'))),'Numbered heading candidate','medium'
    if key in MAJOR: return 1,'Common report section','medium'
    if key in MINOR: return 2,'Common subsection','medium'
    runs=[r for r in p.runs if r.text.strip()]
    if runs and all(r.bold is True for r in runs) and len(s.split())<=14 and not s.endswith(('.', '?','!')):
        return 2,'Short bold paragraph; verify','low'
    return 0,'',''

def parse_outline(text):
    result={}
    for line in text.splitlines():
        s=line.strip()
        if not s: continue
        m=prefix(s)
        level=(1 if CHAPTER.match(s) else len(m.group(1).split('.'))) if m else min(9,1+(len(line)-len(line.lstrip(' ')))//2)
        result[norm(s)]=level
        if m: result[norm(s[m.end():])]=level
    return result

def analyze(data,outline_text=''):
    d=load(data); outline=parse_outline(outline_text); rows=[]
    for i,p in enumerate(d.paragraphs):
        level,reason,confidence=recognize(p,outline)
        if p.text.strip():
            rows.append(dict(index=i,text=p.text[:350],level=level,reason=reason,confidence=confidence,protected=protected(p),unnumbered=unnumbered(p.text)))
    return dict(rows=rows,paragraphs=len(d.paragraphs),tables=len(d.tables),sections=len(d.sections),headings=sum(r['level']>0 for r in rows),protected=sum(r['protected'] for r in rows))

def config(raw):
    c={**DEFAULT,**{k:v for k,v in raw.items() if k in DEFAULT}}
    if not isinstance(c['font'],str) or not 1<=len(c['font'])<=80: raise ValueError('Invalid font.')
    for k in ('size','h1','h2','h3','line','left','right','top','bottom'):
        c[k]=float(c[k]); lo,hi=(0.5,5) if k in ('left','right','top','bottom') else (1,3) if k=='line' else (8,40)
        if not lo<=c[k]<=hi: raise ValueError('Invalid value: '+k)
    for k in ('numbering','toc','tables','figures','chapter_break'):
        if not isinstance(c[k],bool): raise ValueError('Invalid switch: '+k)
    return c

def el(tag,**attrs):
    x=OxmlElement('w:'+tag)
    for k,v in attrs.items(): x.set(qn('w:'+k),str(v))
    return x

def set_font(obj,name,size):
    obj.font.name=name; obj.font.size=Pt(size)
    fonts=obj._element.get_or_add_rPr().rFonts
    for k in ('ascii','hAnsi','eastAsia','cs'): fonts.set(qn('w:'+k),name)
    for k in ('asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme','csTheme'): fonts.attrib.pop(qn('w:'+k),None)

def native_numbering(d):
    root=d.part.numbering_part.element
    a=max([int(x.get(qn('w:abstractNumId'))) for x in root.findall(qn('w:abstractNum'))]+[-1])+1
    n=max([int(x.get(qn('w:numId'))) for x in root.findall(qn('w:num'))]+[0])+1
    abstract=el('abstractNum',abstractNumId=a); abstract.append(el('multiLevelType',val='multilevel'))
    for i in range(9):
        lvl=el('lvl',ilvl=i)
        for tag,val in [('start',1),('numFmt','decimal'),('suff','space'),('lvlText','.'.join('%'+str(j+1) for j in range(i+1))+'.'),('lvlJc','left')]: lvl.append(el(tag,val=val))
        abstract.append(lvl)
    first=root.find(qn('w:num'))
    if first is not None: first.addprevious(abstract)
    else: root.append(abstract)
    num=el('num',numId=n); num.append(el('abstractNumId',val=a)); root.append(num)
    return n

def apply_num(p,n,level):
    np=p._p.get_or_add_pPr().get_or_add_numPr()
    np.get_or_add_ilvl().val=level-1; np.get_or_add_numId().val=n

def strip_prefix(p):
    m=prefix(p.text)
    if not m: return True
    if protected(p) or any(any(child.tag not in {qn('w:rPr'),qn('w:t')} for child in r._r) for r in p.runs): return False
    if ''.join(r.text for r in p.runs)!=p.text: return False
    count=m.end()
    for r in p.runs:
        if count<=0: break
        part=r.text; removed=min(count,len(part)); r.text=part[removed:]; count-=removed
    return count==0

def add_toc(d,first):
    instructions=[x.text or '' for x in d._element.iter(qn('w:instrText'))]+[x.get(qn('w:instr'),'') for x in d._element.iter(qn('w:fldSimple'))]
    if any(re.search(r'\bTOC\b',x,re.I) and not re.search(r'\\c\s',x,re.I) for x in instructions): return False
    title=first.insert_paragraph_before('Table of Contents'); title.style=d.styles['Normal']; title.runs[0].bold=True
    p=first.insert_paragraph_before(); p.add_run()._r.append(el('fldChar',fldCharType='begin'))
    instr=el('instrText'); instr.set(qn('xml:space'),'preserve'); instr.text=' TOC \\o "1-9" \\h \\z \\u '; p.add_run()._r.append(instr)
    p.add_run()._r.append(el('fldChar',fldCharType='separate')); p.add_run('Update this field in Microsoft Word (F9).'); p.add_run()._r.append(el('fldChar',fldCharType='end'))
    return True

def format_doc(data,raw_config,decisions):
    d=load(data); c=config(raw_config)
    changes=dict(heading=0,body=0,tables=0,figures=0,protected=0,toc_added=False,warnings=[])
    warnings=changes['warnings']; chosen={int(k):int(v) for k,v in decisions.items()}; paras=list(d.paragraphs)
    if any(i<0 or i>=len(paras) or not 0<=v<=9 for i,v in chosen.items()): raise ValueError('Invalid paragraph selection.')
    for sec in d.sections:
        landscape=sec.page_width and sec.page_height and sec.page_width>sec.page_height
        sec.page_width=Cm(29.7 if landscape else 21); sec.page_height=Cm(21 if landscape else 29.7)
        for k in ('left','right','top','bottom'): setattr(sec,k+'_margin',Cm(c[k]))
    normal=d.styles['Normal']; set_font(normal,c['font'],c['size']); normal.paragraph_format.line_spacing=c['line']; normal.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    for level in range(1,10):
        s=d.styles[f'Heading {level}']; set_font(s,c['font'],c['h'+str(min(level,3))]); s.font.bold=True; s.paragraph_format.keep_with_next=True; s.paragraph_format.space_before=Pt(12); s.paragraph_format.space_after=Pt(6); s.paragraph_format.line_spacing=c['line']; s.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
    num=native_numbering(d) if c['numbering'] else None; first=None; previous=0
    for i,p in enumerate(paras):
        level=chosen.get(i,heading_level(p)); old=heading_level(p)
        if not p.text.strip():
            if c['figures'] and has(p,{'drawing','pict'}): p.alignment=WD_ALIGN_PARAGRAPH.CENTER; changes['figures']+=1
            continue
        if protected(p):
            changes['protected']+=1
            if level!=old: warnings.append(f'Paragraph {i+1}: protected content; heading change skipped.')
            continue
        if level:
            p.style=d.styles[f'Heading {level}']; pp=p._p.get_or_add_pPr(); ol=pp.find(qn('w:outlineLvl'))
            if ol is not None: pp.remove(ol)
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.line_spacing=c['line']; p.paragraph_format.keep_with_next=True
            for r in p.runs: set_font(r,c['font'],c['h'+str(min(level,3))]); r.bold=True
            if level>previous+1 and not unnumbered(p.text): warnings.append(f'Paragraph {i+1}: heading jumps from level {previous} to {level}.')
            if not unnumbered(p.text):
                if num is not None:
                    if strip_prefix(p): apply_num(p,num,level)
                    else: warnings.append(f'Paragraph {i+1}: numbering skipped to preserve complex text.')
                previous=level
            else: apply_num(p,0,1)
            if level==1 and c['chapter_break'] and first is not None and not unnumbered(p.text): p.paragraph_format.page_break_before=True
            if first is None and not unnumbered(p.text): first=p
            changes['heading']+=1
        else:
            if old and i in chosen:
                p.style=normal; apply_num(p,0,1)
                pp=p._p.get_or_add_pPr(); ol=pp.find(qn('w:outlineLvl'))
                if ol is not None: pp.remove(ol)
            if CAPTION.match(p.text) or (p.style and p.style.name=='Caption'):
                if c['figures']: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            elif p.style and p.style.name in ('Normal','Body Text','No Spacing'):
                p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.line_spacing=c['line']; changes['body']+=1
            else: continue
            for r in p.runs: set_font(r,c['font'],c['size'])
    if c['tables']:
        for t in d.tables:
            seen=set()
            for row in t.rows:
                for cell in row.cells:
                    if cell._tc in seen: continue
                    seen.add(cell._tc)
                    for p in cell.paragraphs:
                        if protected(p): continue
                        for r in p.runs: set_font(r,c['font'],c['size'])
            changes['tables']+=1
    if c['toc'] and first is not None: changes['toc_added']=add_toc(d,first)
    if changes['toc_added']: warnings.append('TOC field inserted. Open in Microsoft Word and update it with F9; page numbers require Word pagination.')
    if changes['protected']: warnings.append('Protected paragraphs were skipped. Review fields, citations, images and equations in Word.')
    output=BytesIO(); d.save(output); return output.getvalue(),changes
