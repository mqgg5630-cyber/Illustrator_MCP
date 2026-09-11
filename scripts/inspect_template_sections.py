import docx
import xml.etree.ElementTree as ET
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

doc = docx.Document(r'E:\0writing\Lark-Formatter\鲁东大学学术学位论文_Zotero活动引用版_new.docx')
for i, s in enumerate(doc.sections):
    print(f'=== Section {i} ===')
    print('  start_type:', s.start_type)
    for hp in s.header.paragraphs:
        print('  Default Header:', repr(hp.text))
        for r in hp.runs:
            print('    run:', repr(r.text))
        print('    XML snippet:', hp._p.xml[:200])

body = doc._body._body
body_sect = body.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
if body_sect is not None:
    print('Body sectPr at end:')
    print(ET.tostring(body_sect, encoding='utf-8').decode('utf-8'))
