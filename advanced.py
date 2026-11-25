"""Advanced formatting, audit, integrity checks and PDF export."""
import io, re, shutil, subprocess, tempfile
from pathlib import Path
from collections import Counter
from zipfile import ZipFile
from lxml import etree
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from engine import load, config, format_doc, heading_level, protected, el, set_font, unnumbered, prefix, parse_outline, ensure_paragraph_style
from citation_guard import compare_citations
from document_preview import compare_images
from layout_tools import cover_elements, fit_tables, resize_images, smart_pages

EXTRA=dict(paper='A4',alignment='justify',before=0,after=6,first_indent=0,table_size=12,caption_size=12,caption_numbers=True,figure_list=True,table_list=True,page_numbers=True,page_start=1,page_position='center',page_format='decimal',update_fields=True,heading_uppercase=False,heading_bold=True,numbering_style='decimal',language='en',reference_style='none',table_fit=False,image_resize=False,smart_pages=False)
def settings(raw):
    c={**config(raw),**EXTRA,**{k:v for k,v in raw.items() if k in EXTRA}}
    for k,v in EXTRA.items():
        if type(v) is bool and type(c[k]) is not bool: raise ValueError('Invalid '+k)
    for k,lo,hi in [('before',0,72),('after',0,72),('first_indent',0,3),('table_size',8,40),('caption_size',8,40),('page_start',1,9999)]:
        c[k]=float(c[k])
        if not lo<=c[k]<=hi: raise ValueError('Invalid '+k)
    for k,allowed in {'paper':['A4','A5','Letter'],'alignment':['justify','left','center'],'page_position':['left','center','right'],'page_format':['decimal','lowerRoman','upperRoman'],'numbering_style':['decimal','chapter','roman'],'language':['en','vi'],'reference_style':['none','APA','Harvard','IEEE']}.items():
        if c[k] not in allowed: raise ValueError('Invalid '+k)
    return c

def fields(data,allow_resize=False):
    counts=Counter()
    with ZipFile(io.BytesIO(data)) as z:
        for name in z.namelist():
            if not name.startswith('word/') or not name.endswith('.xml'): continue
            root=etree.fromstring(z.read(name),parser=etree.XMLParser(resolve_entities=False,no_network=True))
            # Join consecutive instruction fragments within each paragraph; never refresh citations.
            for p in root.iter(qn('w:p')):
                instr=''.join(n.text or '' for n in p.iter(qn('w:instrText')))
                if re.search(r'ADDIN\s+(?:EN\.|EN\b|ZOTERO|CSL|MENDELEY)',instr,re.I): counts[('citation',instr)]+=1
            for n in root.iter(qn('w:fldSimple')):
                instr=n.get(qn('w:instr'),'')
                if re.search(r'ADDIN\s+(?:EN\.|EN\b|ZOTERO|CSL|MENDELEY)',instr,re.I): counts[('citation',instr)]+=1
            for tag in ('oMath','footnoteReference','endnoteReference','drawing','hyperlink','bookmarkStart','commentRangeStart'):
                for n in root.iter():
                    if n.tag.rsplit('}',1)[-1]==tag:
                        if tag=='drawing' and allow_resize:
                            import copy
                            n=copy.deepcopy(n)
                            for x in n.iter():
                                if x.tag in (qn('wp:extent'),qn('a:ext')):
                                    x.attrib.pop('cx',None);x.attrib.pop('cy',None)
                        counts[(tag,etree.tostring(n,method='c14n'))]+=1
        for name in z.namelist():
            if name.startswith(('word/media/','customXml/','word/embeddings/')) and not name.endswith('/'):
                import hashlib
                counts[('part',name,hashlib.sha256(z.read(name)).hexdigest())]+=1
    return counts

def integrity(before,after,allow_resize=False):
    a,b=fields(before,allow_resize),fields(after,allow_resize); missing=a-b
    citations=compare_citations(before,after)
    images=compare_images(before,after)
    return {'passed':not missing and citations['passed'] and images['passed'],'image_integrity':images,'citation_integrity':citations,'protected_objects':sum(a.values()),'citation_fields':sum(v for k,v in a.items() if k[0]=='citation'),'missing':sum(missing.values()),'missing_types':sorted({k[0] for k in missing})}

def audit(data,raw,outline=''):
    d=load(data); c=settings(raw); issues=[]; checks=Counter(); failed=Counter()
    def check(category,ok,location,detail):
        checks[category]+=1
        if not ok:
            failed[category]+=1
            if len(issues)<500: issues.append(dict(category=category,location=location,detail=detail))
    paper={'A4':(21,29.7),'A5':(14.8,21),'Letter':(21.59,27.94)}[c['paper']]
    for i,s in enumerate(d.sections,1):
        for k in ('left','right','top','bottom'):
            v=getattr(s,k+'_margin'); check('Margins',v is not None and abs(v.cm-c[k])<.06,f'Section {i}',f'{k} margin expected {c[k]} cm')
        dims=sorted([s.page_width.cm,s.page_height.cm]) if s.page_width and s.page_height else [0,0]
        check('Paper',all(abs(x-y)<.08 for x,y in zip(dims,sorted(paper))),f'Section {i}',f'Expected {c["paper"]}')
    last=0; actual=set(); caption_counts=Counter(); skip=0
    for i,p in enumerate(d.paragraphs,1):
        if not p.text.strip(): continue
        lev=heading_level(p); loc=f'Paragraph {i}'
        if lev:
            title=p.text.strip(); m=prefix(title); actual.add((title[m.end():] if m else title).casefold().rstrip(':'))
            if not unnumbered(title):
                check('Hierarchy',lev<=last+1,loc,f'Heading level {lev} follows {last}'); last=lev
        if protected(p): skip+=1; continue
        if not lev and p.style.name not in ('Normal','Body Text','No Spacing','Caption'): continue
        expected=c['h'+str(min(lev,3))] if lev else c['caption_size'] if p.style.name=='Caption' else c['size']
        for run in p.runs:
            if not run.text.strip(): continue
            f=run.font.name or p.style.font.name or ensure_paragraph_style(d,'Normal').font.name
            size=run.font.size or p.style.font.size or ensure_paragraph_style(d,'Normal').font.size
            check('Font',f==c['font'],loc,f'Expected {c["font"]}, found {f or "inherited"}')
            check('Font size',size is not None and abs(size.pt-expected)<.1,loc,f'Expected {expected} pt')
        if not lev and p.style.name!='Caption':
            v=p.paragraph_format.line_spacing
            if v is None: v=p.style.paragraph_format.line_spacing
            check('Line spacing',isinstance(v,float) and abs(v-c['line'])<.01,loc,f'Expected {c["line"]}')
    for key in parse_outline(outline):
        if prefix(key): continue
        check('Outline',key in actual,'Outline',f'Missing heading: {key}')
    instr=' '.join(n.text or '' for n in d._element.iter(qn('w:instrText')))
    check('TOC',not c['toc'] or bool(re.search(r'\bTOC\b',instr)), 'Document','Missing TOC field')
    if c['reference_style']!='none':
        check('References',any(unnumbered(p.text) and re.search(r'reference|bibliography|tài liệu tham khảo',p.text,re.I) for p in d.paragraphs),'Document','Missing reference section; citation style needs manual review')
    total=sum(checks.values()); bad=sum(failed.values())
    return {'score':round(100*(total-bad)/total) if total else 100,'checks':total,'failures':bad,'categories':[{'name':k,'checks':v,'failed':failed[k]} for k,v in checks.items()],'issues':issues,'protected_skipped':skip,'note':'Rule checks only; not a guarantee of academic compliance. Citation style and actual pagination require manual review.'}

def field(p,instruction,placeholder='1'):
    p.add_run()._r.append(el('fldChar',fldCharType='begin'))
    n=el('instrText'); n.set(qn('xml:space'),'preserve'); n.text=' '+instruction+' '; p.add_run()._r.append(n)
    p.add_run()._r.append(el('fldChar',fldCharType='separate')); p.add_run(placeholder); p.add_run()._r.append(el('fldChar',fldCharType='end'))

CAP=re.compile(r'^\s*(Figure|Table|Hình|Bảng)\s+(\d+(?:[.\-]\d+)*)\s*[:.\-–]\s+(.+)$',re.I)
def advanced(data,raw,decisions,outline=''):
    c=settings(raw); basic,report=format_doc(data,c,decisions); d=load(basic)
    skip=cover_elements(d,c['protect_cover'])
    if c['protect_cover'] and not skip:
        report.update(integrity=integrity(data,data),audit=audit(data,c,outline),captions={},change_report={'unchanged':True})
        return data,report
    paper={'A4':(21,29.7),'A5':(14.8,21),'Letter':(21.59,27.94)}[c['paper']]
    align={'left':WD_ALIGN_PARAGRAPH.LEFT,'center':WD_ALIGN_PARAGRAPH.CENTER,'justify':WD_ALIGN_PARAGRAPH.JUSTIFY,'right':WD_ALIGN_PARAGRAPH.RIGHT}
    for i,s in enumerate(d.sections):
        if not c['format_layout'] or (skip and i==0): continue
        dims=paper[::-1] if s.page_width>s.page_height else paper
        s.page_width=Cm(dims[0]); s.page_height=Cm(dims[1])
    for n in range(1,10):
        if c['format_headings'] and not skip: ensure_paragraph_style(d,f'Heading {n}').font.bold=c['heading_bold']
    normal=ensure_paragraph_style(d,'Normal')
    if c['format_body'] and not skip: normal.paragraph_format.alignment=align[c['alignment']]
    counts=Counter(); first=None; existing_caption_fields=False
    for p in list(d.paragraphs):
        if p._p in skip: continue
        lev=heading_level(p)
        if first is None and lev and not unnumbered(p.text): first=p
        if protected(p):
            if any('SEQ' in (n.text or '') for n in p._p.iter(qn('w:instrText'))): existing_caption_fields=True
            continue
        m=CAP.match(p.text)
        if m and c['figures']:
            kind='Figure' if m.group(1).casefold() in ('figure','hình') else 'Table'
            counts[kind]+=1
            if c['caption_numbers'] and not lev and all(all(x.tag in (qn('w:rPr'),qn('w:t')) for x in r._r) for r in p.runs):
                # Retain title text and native sequence; existing complex captions are never rewritten.
                p.clear(); p.style=ensure_paragraph_style(d,'Caption'); label=({'Figure':'Hình','Table':'Bảng'}[kind] if c['language']=='vi' else kind)
                p.add_run(label+' '); field(p,'SEQ '+kind+' \\* ARABIC',str(counts[kind])); p.add_run(': '+m.group(3))
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs: set_font(r,c['font'],c['caption_size'])
        elif lev and c['format_headings']:
            for r in p.runs:
                r.bold=c['heading_bold']
                if lev==1 and c['heading_uppercase']: r.text=r.text.upper()
        elif not lev and c['format_body'] and p.style.name in ('Normal','Body Text','No Spacing'):
            p.alignment=align[c['alignment']]; f=p.paragraph_format
            f.space_before=Pt(c['before']); f.space_after=Pt(c['after']); f.first_line_indent=Cm(c['first_indent'])
    if c['tables']:
        for t in d.tables:
            if t._tbl in skip: continue
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if not protected(p):
                            for r in p.runs: set_font(r,c['font'],c['table_size'])
    if c['format_headings'] and c['numbering'] and c['numbering_style']!='decimal' and report['heading']:
        root=d.part.numbering_part.element
        abstracts=root.findall(qn('w:abstractNum'))
        if abstracts:
            for level in abstracts[-1].findall(qn('w:lvl')):
                i=int(level.get(qn('w:ilvl')))
                if i==0:
                    if c['numbering_style']=='roman': level.find(qn('w:numFmt')).set(qn('w:val'),'upperRoman')
                    else: level.find(qn('w:lvlText')).set(qn('w:val'),('Chương ' if c['language']=='vi' else 'Chapter ')+'%1.')
    instructions=' '.join(n.text or '' for n in d._element.iter(qn('w:instrText')))
    if first is not None:
        for k,kind,title in [('figure_list','Figure','Danh mục hình' if c['language']=='vi' else 'List of Figures'),('table_list','Table','Danh mục bảng' if c['language']=='vi' else 'List of Tables')]:
            if c[k] and counts[kind] and c['caption_numbers'] and not re.search(r'TOC\s+\\c\s+"?'+kind,instructions,re.I):
                p=first.insert_paragraph_before(title); p.runs[0].bold=True
                field(first.insert_paragraph_before(),'TOC \\h \\z \\c "'+kind+'"','Update field in Word (F9).')
    if c['page_numbers'] and c['smart_pages']:
        smart_pages(d,c,field,report)
    elif c['page_numbers']:
        for i,s in enumerate(d.sections):
            if skip and i==0: continue
            footer=s.footer
            if skip and i==1 and footer.is_linked_to_previous: footer.is_linked_to_previous=False
            if i and footer.is_linked_to_previous: continue
            if not any(re.search(r'\bPAGE\b',n.text or '') for n in footer._element.iter(qn('w:instrText'))) and not any(re.search(r'\bPAGE\b',n.get(qn('w:instr'),'')) for n in footer._element.iter(qn('w:fldSimple'))):
                p=footer.paragraphs[0] if not footer.paragraphs[0].text and not protected(footer.paragraphs[0]) else footer.add_paragraph()
                p.alignment=align[c['page_position']]; field(p,'PAGE')
            pg=s._sectPr.find(qn('w:pgNumType'))
            if pg is None: pg=el('pgNumType'); s._sectPr.append(pg)
            pg.set(qn('w:fmt'),c['page_format'])
            if i==0: pg.set(qn('w:start'),str(int(c['page_start'])))
    # Avoid global field auto-update when citations are present.
    if any(k[0]=='citation' for k in fields(data)):
        root=d.settings.element; n=root.find(qn('w:updateFields'))
        if n is not None: n.set(qn('w:val'),'false')
    elif c['update_fields']:
        root=d.settings.element; n=root.find(qn('w:updateFields'))
        if n is None: n=el('updateFields'); root.append(n)
        n.set(qn('w:val'),'true')
    report['tables_fitted']=fit_tables(d,skip) if c['table_fit'] else 0
    report['images_resized']=resize_images(d,skip) if c['image_resize'] else 0
    b=io.BytesIO(); d.save(b); out=b.getvalue(); guard=integrity(data,out,c['image_resize'])
    if not guard['passed']: raise ValueError('Integrity check failed: '+', '.join(guard['missing_types'] or ['citation or image relationship'])+'. Original document remains untouched.')
    report.update(integrity=guard,audit=audit(out,c,outline),captions=dict(counts))
    report['change_report']={k:report.get(k,0) for k in ('heading','body','tables','figures','cleanup_runs','citation_paragraphs_formatted','tables_fitted','images_resized','cover_protected','smart_pages_applied','toc_added')}
    if existing_caption_fields: report['warnings'].append('Existing caption fields preserved; refresh and verify sequence identifiers in Word.')
    return out,report

def pdf_export(data):
    import os
    from zipfile import ZipFile
    limit=int(os.getenv('PDF_MAX_MB','5'))*1024*1024
    if len(data)>limit: raise ValueError('PDF preview exceeds configured input size limit.')
    with ZipFile(io.BytesIO(data)) as z:
        if sum(x.file_size for x in z.infolist())>25*1024*1024: raise ValueError('PDF preview exceeds expanded size limit.')
    exe=shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        for candidate in [r'C:\Program Files\LibreOffice\program\soffice.exe',r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']:
            if Path(candidate).exists(): exe=candidate; break
    if not exe: raise ValueError('Install LibreOffice to enable PDF preview and export.')
    load(data)
    with tempfile.TemporaryDirectory() as tmp:
        root=Path(tmp); source=root/'report.docx'; source.write_bytes(data)
        profile=(root/'profile').as_uri()
        try: subprocess.run([exe,'-env:UserInstallation='+profile,'--headless','--convert-to','pdf','--outdir',str(root),str(source)],capture_output=True,timeout=45,check=True)
        except (subprocess.TimeoutExpired,subprocess.CalledProcessError) as exc: raise ValueError('PDF rendering failed.') from exc
        pdf=root/'report.pdf'
        if not pdf.exists(): raise ValueError('LibreOffice did not produce a PDF.')
        if pdf.stat().st_size>20*1024*1024: raise ValueError('Rendered PDF exceeds 20 MB.')
        return pdf.read_bytes()
