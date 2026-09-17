import subprocess,xml.etree.ElementTree as ET,json,sys,os
from collections import Counter,defaultdict
from pypdf import PdfReader
from pathlib import Path
root=Path(os.environ['FIDELITY_OUTPUT_DIR'])
for id in sys.argv[1:]:
 def read(suffix):
  p=root/(id+suffix+'.pdf');r=PdfReader(p);ns='{http://www.w3.org/1999/xhtml}';x=ET.fromstring(subprocess.check_output(['pdftotext','-bbox',str(p),'-'],stderr=subprocess.DEVNULL));pages=x.findall(f'.//{ns}page');w=[(i,e.text,{k:float(v) for k,v in e.attrib.items()}) for i,q in enumerate(pages) for e in q.findall(f'.//{ns}word')]
  refs={p.indirect_reference.idnum:i for i,p in enumerate(r.pages)}
  links=[]
  for i,p in enumerate(r.pages):
   for ar in p.get('/Annots',[]):
    a=ar.get_object()
    if a.get('/Subtype')!='/Link':continue
    ac=a['/A'].get_object() if '/A' in a else {};d=a.get('/Dest',ac.get('/D'))
    dest=(refs.get(d[0].idnum),*[str(x) for x in d[1:]]) if isinstance(d,list) else str(d)
    links.append((i,str(ac.get('/URI','')),dest))
  return w,links,len(r.pages)
 a,al,ap=read('-after');b,bl,bp=read('-final');same=Counter((i,t)for i,t,r in a)==Counter((i,t)for i,t,r in b)
 remaining=defaultdict(list)
 for i,t,r in b:remaining[(i,t)].append(r)
 delta=0
 for i,t,r in a:
  choices=remaining[(i,t)]
  if not choices:delta=float('inf');break
  nearest=min(range(len(choices)),key=lambda n:sum(abs(r[k]-choices[n][k]) for k in r))
  s=choices.pop(nearest);delta=max(delta,max(abs(r[k]-s[k]) for k in r))
 result={'id':id,'pages':bp,'samePageCount':ap==bp,'sameTextAndPageAssignment':same,'words':len(b),'maxWordCoordinateDeltaPt':delta,'sameLinks':al==bl,'links':len(bl)}
 print(json.dumps(result));(root/(id+'-compression-check.json')).write_text(json.dumps(result,indent=2))
 assert same and ap==bp and delta<.002 and al==bl
