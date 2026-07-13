from docx import Document
from docx.oxml.ns import qn

doc = Document('nota_dinas.docx')
table = doc.tables[2]
row = table.rows[2]

for ci in [3, 4, 5]:
    cell = row.cells[ci]
    tc = cell._tc
    paras = tc.findall(qn('w:p'))
    print(f'--- Cell [2,{ci}] ---')
    for pi, p in enumerate(paras):
        pPr = p.find(qn('w:pPr'))
        # Look for tabs
        runs = p.findall(qn('w:r'))
        tab_count = sum(1 for r in runs if r.find(qn('w:tab')) is not None)
        texts = [t.text for t in p.findall('.//' + qn('w:t')) if t.text]
        print(f'  P{pi}: texts={texts} tabs={tab_count}')
        
        # alignment
        if pPr is not None:
            jc = pPr.find(qn('w:jc'))
            if jc is not None:
                print("    align=" + jc.get(qn("w:val")))
            ind = pPr.find(qn('w:ind'))
            if ind is not None:
                left = ind.get(qn("w:left"))
                print("    indent_left=" + str(left))
