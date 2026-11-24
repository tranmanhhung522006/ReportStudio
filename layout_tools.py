"""Conservative section, table and image helpers for v2.2."""
from docx.oxml.ns import qn
from docx.shared import Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from engine import el, protected

def cover_elements(d,enabled):
    if not enabled or len(d.sections)<2: return set()
    result=set()
    for child in d.element.body:
        result.add(child)
        if child.tag==qn('w:p') and child.pPr is not None and child.pPr.find(qn('w:sectPr')) is not None: break
    return result

def safe_runs(p):
    return bool(p.runs) and all(all(n.tag in (qn('w:rPr'),qn('w:t')) for n in r._r) for r in p.runs)

def fit_tables(d,skip):
    from docx.table import Table
    section=0; count=0
    for child in d.element.body:
        if child.tag==qn('w:tbl') and child not in skip:
            sec=d.sections[min(section,len(d.sections)-1)]
            width=sec.page_width-sec.left_margin-sec.right_margin
            table=Table(child,d._body)
            cols=list(table.columns)
            if not cols or any(c.width is None for c in cols): continue
            total=sum(c.width for c in cols)
            if total<=width or total<=0: continue
            for col in cols: col.width=Emu(int(col.width*width/total))
            seen=set()
            for row in table.rows:
                for cell in row.cells:
                    if cell._tc in seen: continue
                    seen.add(cell._tc)
                    if cell.width: cell.width=Emu(int(cell.width*width/total))
            tblw=child.tblPr.find(qn('w:tblW'))
            if tblw is not None:
                tblw.set(qn('w:type'),'dxa');tblw.set(qn('w:w'),str(int(width/635)))
            table.autofit=False;count+=1
        if child.tag==qn('w:p') and child.pPr is not None and child.pPr.find(qn('w:sectPr')) is not None: section+=1
    return count

def resize_images(d,skip):
    from docx.shape import InlineShape
    section=0;count=0
    for child in d.element.body:
        if child not in skip:
            sec=d.sections[min(section,len(d.sections)-1)]
            width=sec.page_width-sec.left_margin-sec.right_margin
            height=sec.page_height-sec.top_margin-sec.bottom_margin
            for inline in child.iter(qn('wp:inline')):
                shape=InlineShape(inline)
                if not shape.width or not shape.height: continue
                ratio=min(1,width/shape.width,height/shape.height)
                if ratio<1:
                    w,h=shape.width,shape.height
                    shape.width=Emu(max(1,int(w*ratio)));shape.height=Emu(max(1,int(h*ratio)));count+=1
        if child.tag==qn('w:p') and child.pPr is not None and child.pPr.find(qn('w:sectPr')) is not None: section+=1
    return count

def smart_pages(d,c,field,report):
    # Use explicit existing section breaks rather than guessing physical page boundaries.
    if len(d.sections)<3:
        report['warnings'].append('Smart page numbering needs three existing sections: cover, front matter, body. Existing page numbering kept.')
        return
    for i,sec in enumerate(d.sections):
        if i>2: break
        # Avoid destructive replacement of existing non-empty cover footer or linked variants.
        if i==0:
            for footer in (sec.footer,sec.first_page_footer,sec.even_page_footer):
                if not footer.is_linked_to_previous and any(p.text.strip() or protected(p) for p in footer.paragraphs):
                    report['warnings'].append('Cover footer already has content; smart numbering skipped to preserve it.')
                    return
    align={'left':WD_ALIGN_PARAGRAPH.LEFT,'center':WD_ALIGN_PARAGRAPH.CENTER,'right':WD_ALIGN_PARAGRAPH.RIGHT}
    for i,sec in enumerate(d.sections):
        if i==0: continue
        pg=sec._sectPr.find(qn('w:pgNumType'))
        if pg is None: pg=el('pgNumType'); sec._sectPr.append(pg)
        pg.set(qn('w:fmt'),'lowerRoman' if i==1 else 'decimal')
        if i in (1,2): pg.set(qn('w:start'),'1')
        else: pg.attrib.pop(qn('w:start'),None)
        footers=[sec.footer]
        if sec.different_first_page_header_footer: footers.append(sec.first_page_footer)
        if d.settings.odd_and_even_pages_header_footer: footers.append(sec.even_page_footer)
        for footer in footers:
            if i in (1,2) and footer.is_linked_to_previous: footer.is_linked_to_previous=False
            if footer.is_linked_to_previous: continue
            if not any('PAGE' in (n.text or '').upper().split() for n in footer._element.iter(qn('w:instrText'))) and not any('PAGE' in n.get(qn('w:instr'),'').upper().split() for n in footer._element.iter(qn('w:fldSimple'))):
                p=footer.paragraphs[0] if not footer.paragraphs[0].text and not protected(footer.paragraphs[0]) else footer.add_paragraph()
                p.alignment=align[c['page_position']];field(p,'PAGE')
    report['smart_pages_applied']=True
