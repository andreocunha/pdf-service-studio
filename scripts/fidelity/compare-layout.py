import json,sys,collections,os
from pathlib import Path
id=sys.argv[1];root=Path(os.environ['FIDELITY_OUTPUT_DIR']);a=json.loads((root/(id+'-editor.json')).read_text())['snapshot'];b=json.loads((root/(id+'-export.json')).read_text())['snapshot']
x={z['id']:(i,z) for i,p in enumerate(a['pages']) for z in p['blocks']};y={z['id']:(i,z) for i,p in enumerate(b['pages']) for z in p['blocks']};errs=collections.Counter();details=[];maxdelta=collections.defaultdict(float)
if [[z['id'] for z in p['blocks']] for p in a['pages']] != [[z['id'] for z in p['blocks']] for p in b['pages']]: errs['blockOrder']+=1
if [(p['width'],p['height']) for p in a['pages']] != [(p['width'],p['height']) for p in b['pages']]: errs['pageGeometry']+=1
def compare_rect(r,s):
 for prop in ['x','y','width','height']:maxdelta[prop]=max(maxdelta[prop],abs(r[prop]-s[prop]))
for p,q in zip(a['pages'],b['pages']):
 for kind in ['images','navigation']:
  if len(p[kind])!=len(q[kind]):errs[kind+'Count']+=1
  for u,v in zip(p[kind],q[kind]):
   if {k:x for k,x in u.items() if k!='rect'}!={k:x for k,x in v.items() if k!='rect'}:errs[kind+'Content']+=1
   compare_rect(u['rect'],v['rect'])
for k,(pi,blk) in x.items():
 if k not in y:errs['missingBlocks']+=1;continue
 qi,c=y[k]
 compare_rect(blk['rect'],c['rect'])
 if pi!=qi:errs['changedPages']+=1
 if len(blk['nodes'])!=len(c['nodes']):errs['nodeCounts']+=1
 for u,v in zip(blk['nodes'],c['nodes']):
  for prop in ['text','font','size','weight','lineHeight','transform']:
   if u[prop]!=v[prop]:errs[prop]+=1;details.append([k,prop,u[prop],v[prop]])
  if len(u['rects'])!=len(v['rects']):errs['lineCounts']+=1;details.append([k,'lineCounts',u['text'][:60],len(u['rects']),len(v['rects'])])
  for r,s in zip(u['rects'],v['rects']):compare_rect(r,s)
report={'id':id,'editorPages':len(a['pages']),'exportPages':len(b['pages']),'blocks':[len(x),len(y)],'differences':dict(errs),'maxRectDeltaPx':dict(maxdelta),'samples':details[:10]}
(root/(id+'-comparison.json')).write_text(json.dumps(report,indent=2));print(json.dumps(report,indent=2))

if report['differences'] or len(a['pages']) != len(b['pages']) or set(x) != set(y) or any(d > 0.001 for d in maxdelta.values()):
 raise SystemExit(1)
print('PASS: same pages, blocks, text, fonts, wrapping and text rectangles (0.001 CSS px tolerance).')
