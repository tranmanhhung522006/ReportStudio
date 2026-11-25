"""DOCX formatter: local, reviewable, and non-destructive to the input."""
from io import BytesIO
import re
import unicodedata
from zipfile import ZipFile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

MAX_BYTES = 15 * 1024 * 1024
MAX_EXPANDED = 50 * 1024 * 1024
MAX_PARAGRAPHS = 12000
DEFAULT = dict(font='Times New Roman',size=13,line=1.5,left=3,right=2,top=2,bottom=2,h1=16,h2=14,h3=13,numbering=True,toc=True,tables=True,figures=True,chapter_break=False,format_body=True,format_headings=True,format_layout=True,protect_cover=False,cleanup_text=False,format_citations=False)
PREFIX = re.compile(r'^\s*(\d+(?:\.\d+){0,8})[.)]?\s+(?=\D)')
CHAPTER = re.compile(r'^\s*(?:chapter|chương)\s+(\d+|[IVXLCDM]+)\s*[:.\-–]?\s+',re.I)
UNNUMBERED = set('abstract|acknowledgements|acknowledgments|references|bibliography|appendix|appendices|tóm tắt|lời cảm ơn|tài liệu tham khảo|phụ lục|table of contents|mục lục|list of figures|list of tables|danh mục hình|danh mục bảng'.split('|'))
MAJOR = set('introduction|literature review|research methodology|methodology|results|discussion|results and discussion|conclusion|conclusions|recommendations|system design|implementation|testing and evaluation|requirement analysis|giới thiệu|tổng quan|phương pháp nghiên cứu|kết quả|thảo luận|kết luận'.split('|'))
MINOR = set('research background|research objectives|research questions|theoretical framework|hypothesis development|data collection|sampling method|research design|data analysis|research gap|scope of study|problem statement|significance of study|bối cảnh nghiên cứu|mục tiêu nghiên cứu|câu hỏi nghiên cứu|phương pháp thu thập dữ liệu'.split('|'))
CAPTION = re.compile(r'^\s*(?:figure|table|hình|bảng)\s+\d+(?:[.\-]\d+)*\s*[:.\-–]\s+\S',re.I)
PROTECTED = set('fldChar fldSimple instrText sdt ins del oMath oMathPara hyperlink footnoteReference endnoteReference drawing pict object AlternateContent imagedata blip svgBlip shape image graphic textbox txbxContent bookmarkStart bookmarkEnd commentRangeStart commentRangeEnd'.split())

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
        # Custom/localized paragraph styles may carry the native outline level.
        sp=style._element.pPr
        if sp is not None:
            ol=sp.find(qn('w:outlineLvl'))
            if ol is not None:
                val=ol.get(qn('w:val'),'9')
                if val.isdigit() and 0<=int(val)<9: return int(val)+1
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
    if not s or len(s)>200 or CAPTION.match(s) or protected(p) or (p.style and p.style.name in ('Title','Subtitle','Caption')): return 0,'',''
    key=norm(s)
    m=prefix(s)
    if key in outline or (m and norm(s[m.end():]) in outline): return outline.get(key,outline.get(norm(s[m.end():]) if m else key)),'Provided outline','high'
    if (p._p.pPr is not None and p._p.pPr.numPr is not None) or (p.style and 'list' in p.style.name.lower()): return 0,'',''
    if unnumbered(s): return 1,'Unnumbered section','high'
    if CHAPTER.match(s): return 1,'Chapter prefix','high'
    m=PREFIX.match(s)
    if m and len(s.split())<=12 and not s.endswith(('.', '?', '!')) and not re.search(r'[;:]\s+\S',s[m.end():]):
        depth=min(9,len(m.group(1).split('.')))
        title=norm(s[m.end():]); confidence='high' if title in MAJOR or title in MINOR else 'medium'
        return depth,'Numbered heading candidate',confidence
    if key in MAJOR: return 1,'Common report section','medium'
    if key in MINOR: return 2,'Common subsection','medium'
    runs=[r for r in p.runs if r.text.strip()]
    if runs and all(r.bold is True or (r.bold is None and p.style and p.style.font.bold is True) for r in runs) and len(s.split())<=14 and not s.endswith(('.', '?','!')):
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

def heading_features(p):
    from collections import Counter
    runs=[r for r in p.runs if r.text.strip()]
    def effective(r,attr):
        value=getattr(r.font,attr)
        if value is not None:return value
        for st in (r.style,p.style):
            seen=set()
            while st is not None and st.style_id not in seen:
                seen.add(st.style_id)
                v=getattr(st.font,attr)
                if v is not None:return v
                st=st.base_style
        return None
    sizes=Counter()
    for r in runs:
        v=effective(r,'size')
        if v is not None:sizes[round(v.pt,1)]+=len(r.text)
    size=sizes.most_common(1)[0][0] if sizes else None
    bold=bool(runs) and all(effective(r,'bold') is True for r in runs)
    f=p.paragraph_format
    before=f.space_before
    if before is None and p.style:before=p.style.paragraph_format.space_before
    keep=f.keep_with_next
    if keep is None and p.style:keep=p.style.paragraph_format.keep_with_next
    return {'size':size,'bold':bold,'space_before':round(before.pt,1) if before is not None else 0,'keep_next':bool(keep),'style':p.style.name if p.style else ''}

STANDALONE_CHAPTER=re.compile(r'^\s*(?:chapter|chương|phần|part)\s+(?:\d+|[IVXLCDM]+)\s*[:.\-–]?\s*$',re.I)
CASE_TITLE=re.compile(r'^(?:case(?:\s+study)?|tình huống)\s*(?:\d+)?\s*:\s*\S',re.I)
QUESTION=re.compile(r'^(?:ans(?:wer)?\s*\d*\s*:|question\s*\d*\s*:|câu\s*hỏi\s*\d*\s*:|trả lời\s*:|what\b|why\b|how\b|when\b|where\b|which\b|who\b|describe\b|explain\b|discuss\b|analy[sz]e\b|compare\b|identify\b|list\b|tại sao\b|vì sao\b|hãy\b)',re.I)
CAPTION_LIKE=re.compile(r'^(?:fig(?:ure)?\.?|table|hình|bảng|biểu đồ|chart)\s+\d+(?:[.\-]\d+)*(?:\s|[:.\-–]|$)',re.I)

def analyze(data,outline_text=''):
    from collections import Counter
    d=load(data); outline=parse_outline(outline_text); rows=[]
    paras=list(d.paragraphs); features=[heading_features(p) if p.text.strip() else {} for p in paras]
    sizes=Counter()
    for p,f in zip(paras,features):
        if len(p.text.split())>=18 and not heading_level(p) and f.get('size'):
            sizes[f['size']]+=len(p.text)
    baseline=sizes.most_common(1)[0][0] if sizes else None
    signatures=Counter((f.get('style'),f.get('size'),f.get('bold')) for p,f in zip(paras,features)
        if f and 2<=len(p.text.split())<=14 and not p.text.strip().endswith(('.', '?','!')) and f.get('bold'))
    previous=0
    for i,p in enumerate(paras):
        t=p.text.strip()
        if not t:continue
        f=features[i];old=heading_level(p);m=prefix(t)
        title=(t[m.end():] if m else t).strip();key=norm(title)
        explicit=outline.get(norm(t),outline.get(key))
        isprotected=protected(p)
        candidate=0;confidence='';reason='Body text'
        native_list=bool((p._p.pPr is not None and p._p.pPr.numPr is not None) or 'list' in f['style'].lower())
        toc_style=bool(re.match(r'^(toc|table of figures|mục lục|danh mục hình|danh mục bảng)\s*\d*$',f['style'],re.I))
        large=bool(baseline and f['size'] and f['size']>=baseline+1.5)
        distinctive=large or f['bold'] or f['keep_next'] or f['space_before']>=10
        repeated=signatures[(f['style'],f['size'],f['bold'])]>=3
        if isprotected:
            candidate=old;confidence='high' if old else '';reason='Protected Word content'
        elif explicit is not None:
            candidate=explicit;confidence='high';reason='Reference outline match'
        elif toc_style or CAPTION_LIKE.match(t) or f['style'] in ('Title','Subtitle','Caption'):
            reason='Caption, title or generated index'
        elif old:
            candidate=old;confidence='high';reason='Existing Word heading'
        elif STANDALONE_CHAPTER.fullmatch(t) or CHAPTER.match(t):
            candidate=1;confidence='high';reason='Explicit chapter label'
        elif unnumbered(t):
            candidate=1;confidence='high';reason='Recognized unnumbered section'
        elif native_list:
            reason='Native list item'
        elif QUESTION.match(title) or t.endswith('?') or len(t)>200 or len(t.split())>18:
            reason='Question, answer or body sentence'
        elif CASE_TITLE.match(title):
            candidate=2 if previous else 1;confidence='high' if previous and distinctive else 'medium';reason='Case study title'
        elif m and (key in MAJOR or key in MINOR):
            candidate=min(9,len(m.group(1).split('.')));confidence='high';reason='Numbered known section'
        elif key in MAJOR or key in MINOR:
            candidate=1 if key in MAJOR else 2;confidence='high' if distinctive else 'medium';reason='Known academic section'
        elif m and not t.endswith(('.', '!', '?')):
            candidate=min(9,len(m.group(1).split('.')))
            confidence='high' if candidate>=2 and distinctive and candidate<=previous+1 else 'medium' if distinctive else 'low'
            reason='Numbered title with formatting evidence' if distinctive else 'Ambiguous numbered paragraph'
        elif distinctive and not t.endswith(('.', '!', '?')) and len(t.split())<=14:
            candidate=1 if large and baseline and f['size']>=baseline+3 else min(3,max(2,previous))
            confidence='medium' if large or (f['bold'] and repeated and (f['keep_next'] or f['space_before']>=10)) else 'low'
            reason='Larger text or repeated heading format' if confidence=='medium' else 'Short emphasized text; review'
        # Only apply high-confidence new headings automatically. Keep existing headings.
        level=candidate if confidence=='high' else 0
        warning=''
        if candidate and not unnumbered(t) and candidate>previous+1:
            warning='Heading level jumps over a parent'
            if not old and explicit is None and confidence=='high' and not isprotected:
                confidence='medium';level=0
        if level and not unnumbered(t):previous=level
        rows.append(dict(index=i,text=p.text[:350],level=level,suggested_level=candidate,original_level=old,
                         reason=reason,confidence=confidence,protected=isprotected,unnumbered=unnumbered(t),
                         warning=warning,features=f,format_group='|'.join(str(f.get(k,'')) for k in ('style','size','bold','space_before','keep_next'))))
    return dict(rows=rows,paragraphs=len(paras),tables=len(d.tables),sections=len(d.sections),
                headings=sum(r['level']>0 for r in rows),protected=sum(r['protected'] for r in rows),
                needs_review=sum(r['suggested_level']>0 and r['confidence']!='high' for r in rows),body_font_size=baseline)

def config(raw):
    c={**DEFAULT,**{k:v for k,v in raw.items() if k in DEFAULT}}
    if not isinstance(c['font'],str) or not 1<=len(c['font'])<=80: raise ValueError('Invalid font.')
    for k in ('size','h1','h2','h3','line','left','right','top','bottom'):
        c[k]=float(c[k]); lo,hi=(0.5,5) if k in ('left','right','top','bottom') else (1,3) if k=='line' else (8,40)
        if not lo<=c[k]<=hi: raise ValueError('Invalid value: '+k)
    for k in ('numbering','toc','tables','figures','chapter_break','format_body','format_headings','format_layout','protect_cover','cleanup_text','format_citations'):
        if not isinstance(c[k],bool): raise ValueError('Invalid switch: '+k)
    return c

def el(tag,**attrs):
    x=OxmlElement('w:'+tag)
    for k,v in attrs.items(): x.set(qn('w:'+k),str(v))
    return x

def set_font(obj,name,size):
    obj.font.name=name; obj.font.size=Pt(size)
    rp=obj._element.get_or_add_rPr()
    # Word can use szCs rather than sz for complex-script text.
    cs=rp.find(qn('w:szCs'))
    if cs is not None: cs.set(qn('w:val'),str(round(size*2)))
    fonts=obj._element.get_or_add_rPr().rFonts
    for k in ('ascii','hAnsi','eastAsia','cs'): fonts.set(qn('w:'+k),name)
    for k in ('asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme','csTheme'): fonts.attrib.pop(qn('w:'+k),None)


def ensure_paragraph_style(d, name):
    """Resolve existing named/built-in paragraph styles; create only if absent."""
    styles=d.styles
    canonical=name.replace(' ','').casefold()
    for style in styles:
        if style.type==WD_STYLE_TYPE.PARAGRAPH and ((style.name or '').casefold()==name.casefold() or (style.style_id or '').casefold()==canonical):
            return style
    # A character/table style can have the same display name. Avoid collisions.
    used_names={(s.name or '').casefold() for s in styles}
    used_ids={(s.style_id or '').casefold() for s in styles}
    candidate=name
    suffix=1
    while candidate.casefold() in used_names or candidate.replace(' ','').casefold() in used_ids:
        candidate=f'ReportStudio {name} {suffix}'
        suffix+=1
    style=styles.add_style(candidate, WD_STYLE_TYPE.PARAGRAPH)
    if name!='Normal':
        style.base_style=ensure_paragraph_style(d,'Normal')
    match=re.fullmatch(r'Heading ([1-9])',name)
    if match:
        level=int(match.group(1))
        pp=style._element.get_or_add_pPr()
        pp.append(el('outlineLvl',val=level-1))
        style.paragraph_format.keep_with_next=True
        style.font.bold=True
        style.next_paragraph_style=ensure_paragraph_style(d,'Normal')
    elif name=='Normal':
        style._element.set(qn('w:default'),'1')
    return style

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
    title=first.insert_paragraph_before('Table of Contents'); title.style=ensure_paragraph_style(d,'Normal'); title.runs[0].bold=True
    p=first.insert_paragraph_before(); p.add_run()._r.append(el('fldChar',fldCharType='begin'))
    instr=el('instrText'); instr.set(qn('xml:space'),'preserve'); instr.text=' TOC \\o "1-9" \\h \\z \\u '; p.add_run()._r.append(instr)
    p.add_run()._r.append(el('fldChar',fldCharType='separate')); p.add_run('Update this field in Microsoft Word (F9).'); p.add_run()._r.append(el('fldChar',fldCharType='end'))
    return True

def format_doc(data,raw_config,decisions):
    d=load(data); c=config(raw_config)
    from layout_tools import cover_elements, safe_runs
    skip=cover_elements(d,c['protect_cover'])
    changes=dict(heading=0,body=0,tables=0,figures=0,protected=0,toc_added=False,warnings=[])
    warnings=changes['warnings']; changes['cover_protected']=bool(skip); changes['cleanup_runs']=0; changes['citation_paragraphs_formatted']=0
    if c['protect_cover'] and not skip: raise ValueError('Cover protection requires a section break after the cover. Turn off cover protection or add a section break in Word.')
    chosen={int(k):int(v) for k,v in decisions.items()}; paras=list(d.paragraphs)
    if any(i<0 or i>=len(paras) or not 0<=v<=9 for i,v in chosen.items()): raise ValueError('Invalid paragraph selection.')
    for i,sec in enumerate(d.sections):
        if not c['format_layout'] or (skip and i==0): continue
        landscape=sec.page_width and sec.page_height and sec.page_width>sec.page_height
        sec.page_width=Cm(29.7 if landscape else 21); sec.page_height=Cm(21 if landscape else 29.7)
        for k in ('left','right','top','bottom'): setattr(sec,k+'_margin',Cm(c[k]))
    normal=ensure_paragraph_style(d,'Normal')
    if c['format_body'] and not skip: set_font(normal,c['font'],c['size']); normal.paragraph_format.line_spacing=c['line']; normal.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    for level in range(1,10):
        if not c['format_headings'] or skip: continue
        s=ensure_paragraph_style(d,f'Heading {level}'); set_font(s,c['font'],c['h'+str(min(level,3))]); s.font.bold=True; s.paragraph_format.keep_with_next=True; s.paragraph_format.space_before=Pt(12); s.paragraph_format.space_after=Pt(6); s.paragraph_format.line_spacing=c['line']; s.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
    num=native_numbering(d) if c['numbering'] and c['format_headings'] else None; first=None; previous=0
    for i,p in enumerate(paras):
        if p._p in skip: continue
        level=chosen.get(i,heading_level(p)); old=heading_level(p)
        if not p.text.strip():
            if c['figures'] and has(p,{'drawing','pict'}): p.alignment=WD_ALIGN_PARAGRAPH.CENTER; changes['figures']+=1
            continue
        if protected(p):
            changes['protected']+=1
            if c['format_body'] and c['format_citations'] and not old and any(re.search(r'ADDIN\s+(?:EN\.|EN\b|ZOTERO|CSL|MENDELEY)',n.text or '',re.I) for n in p._p.iter(qn('w:instrText'))):
                p.paragraph_format.line_spacing=c['line']; p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; changes['citation_paragraphs_formatted']+=1
            # Format only plain-text runs outside fields; never replace XML or image data.
            if c['format_body'] and not old and not level and not has(p,{'sdt','ins','del','oMath','oMathPara','txbxContent'}):
                in_field=0
                for run in p.runs:
                    markers=list(run._r.iter(qn('w:fldChar')))
                    if markers:
                        for marker in markers:
                            kind=marker.get(qn('w:fldCharType'))
                            if kind=='begin': in_field+=1
                            elif kind=='end': in_field=max(0,in_field-1)
                        continue
                    if not in_field and run.text.strip() and all(x.tag in (qn('w:rPr'),qn('w:t'),qn('w:tab'),qn('w:br')) for x in run._r):
                        set_font(run,c['font'],c['size'])
                        changes['protected_text_runs_formatted']=changes.get('protected_text_runs_formatted',0)+1
            if level!=old: warnings.append(f'Paragraph {i+1}: protected content; heading change skipped.')
            continue
        if level and not c['format_headings']:
            if first is None and old and not unnumbered(p.text): first=p
            continue
        if level:
            p.style=ensure_paragraph_style(d,f'Heading {level}'); pp=p._p.get_or_add_pPr(); ol=pp.find(qn('w:outlineLvl'))
            if ol is not None: pp.remove(ol)
            p.paragraph_format.left_indent=Cm(0); p.paragraph_format.right_indent=Cm(0); p.paragraph_format.first_line_indent=Cm(0)
            p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.line_spacing=c['line']; p.paragraph_format.keep_with_next=True; p.paragraph_format.space_before=Pt(12); p.paragraph_format.space_after=Pt(6)
            for r in p.runs: set_font(r,c['font'],c['h'+str(min(level,3))]); r.bold=True
            if level>previous+1 and not unnumbered(p.text): warnings.append(f'Paragraph {i+1}: heading jumps from level {previous} to {level}.')
            if not unnumbered(p.text):
                if num is not None:
                    if strip_prefix(p): apply_num(p,num,level)
                    else: warnings.append(f'Paragraph {i+1}: numbering skipped to preserve complex text.')
                previous=level
            elif c['numbering']: apply_num(p,0,1)
            if level==1 and c['chapter_break'] and first is not None and not unnumbered(p.text): p.paragraph_format.page_break_before=True
            if first is None and not unnumbered(p.text): first=p
            changes['heading']+=1
        else:
            if old and not c['format_headings']: continue
            if old and i in chosen:
                p.style=normal; apply_num(p,0,1)
                pp=p._p.get_or_add_pPr(); ol=pp.find(qn('w:outlineLvl'))
                if ol is not None: pp.remove(ol)
            if CAPTION.match(p.text) or (p.style and p.style.name=='Caption'):
                if c['figures']: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                if not c['format_body']: continue
            elif not old or i in chosen:
                if not c['format_body']: continue
                if p.style and p.style.name in ('Title','Subtitle'): continue
                if p._p.pPr is not None and p._p.pPr.numPr is None:
                    p.paragraph_format.left_indent=Cm(0); p.paragraph_format.right_indent=Cm(0)
                p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.line_spacing=c['line']; changes['body']+=1
            else: continue
            for r in p.runs:
                set_font(r,c['font'],c['size'])
                if c['cleanup_text'] and not old and safe_runs(p):
                    text=re.sub(r' {2,}',' ',r.text)
                    if text!=r.text: r.text=text; changes['cleanup_runs']+=1
    if c['tables']:
        for t in d.tables:
            if t._tbl in skip: continue
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
