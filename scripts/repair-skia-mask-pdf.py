#!/usr/bin/env python3
"""Offline repair for already-exported PDFs. This is not the production renderer.

Recognizes and validates Skia's two-pass SrcIn wrapper, extracts only its
mask, and replaces the wrapper with a vector clip. Original content streams,
font programs, images, page geometry and annotations are kept. Never use on a
signed PDF: any modification invalidates its signatures. Validate output with
check-pdf-regression.py and both Quartz and Poppler before delivery.
"""
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject as N, DictionaryObject as D, DecodedStreamObject, RectangleObject
from pathlib import Path
from PIL import Image, ImageFilter
import argparse
import tempfile
import re, subprocess, json
from datetime import datetime, timezone
parser = argparse.ArgumentParser(description="Repair Skia hard-edge mask wrappers in an existing PDF without reflowing text. Requires pypdf, Pillow and pdftocairo.")
parser.add_argument('source', type=Path)
parser.add_argument('output', type=Path)
args = parser.parse_args()
source, output = args.source, args.output
if source.resolve() == output.resolve():
    raise SystemExit('Use a different output path; the reference PDF must remain untouched.')
work = tempfile.TemporaryDirectory(prefix='lex-pdf-masks-')
scratch = Path(work.name)
r=PdfReader(source);w=PdfWriter(clone_from=r)
w.pdf_header=r.pdf_header
w.add_metadata({'/Producer':str(r.metadata.producer or '')+'; Lex mask compatibility', '/ModDate':datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%S+00'00'")})
mask_writer=PdfWriter();targets=[]
pattern=re.compile(rb'^0 0 0 RG\s+0 0 0 rg\s+(/G\d+) gs\s+(/G\d+) gs\s+(/X\d+) Do\s+(/G\d+) gs\s+(/G\d+) gs\s+\3 Do\s+\4 gs\s*$')
for obj in w._objects:
 if not isinstance(obj,dict) or obj.get('/Subtype')!='/Form':continue
 match=pattern.match(obj.get_data())
 if not match:continue
 base,inv,content,reset,mask=[x.decode() for x in match.groups()]
 ext=obj['/Resources']['/ExtGState']
 assert ext[reset]['/SMask']=='/None'
 inverse=ext[inv]['/SMask']; sm=ext[mask]['/SMask']
 assert inverse['/S']=='/Alpha' and sm['/S']=='/Alpha' and '/TR' not in sm
 assert re.sub(rb'\s+',b' ',inverse['/TR'].get_object().get_data()).strip()==b'{1 exch sub}'
 group=sm['/G']; box=list(group['/BBox'])
 assert box[:2]==[0,0] and list(inverse['/G']['/BBox'])==box
 expected=f'0 0 0 RG 0 0 0 rg /G3 gs 0 0 {box[2]} {box[3]} re f'
 assert re.sub(rb'\s+',b' ',inverse['/G'].get_data()).decode().strip()==expected
 p=mask_writer.add_blank_page(width=float(box[2]),height=float(box[3]))
 p[N('/Resources')]=D({N('/XObject'):D({N('/MaskShape'):group.clone(mask_writer).indirect_reference})})
 stream=DecodedStreamObject();stream.set_data(b'/MaskShape Do\n');p[N('/Contents')]=mask_writer._add_object(stream)
 targets.append((obj,base,content,reset,box))
if not targets: raise SystemExit('No supported Skia mask wrappers found.')
mask_writer.write(str(scratch / 'masks.pdf'))
subprocess.run(['pdftocairo','-png','-transp','-r','72',str(scratch / 'masks.pdf'),str(scratch / 'mask')],check=True)
for i,(obj,base,content,reset,box) in enumerate(targets,1):
 im=Image.open(str(scratch / f'mask-{i:0{len(str(len(targets)))}d}.png')).convert('RGBA');alpha=im.getchannel('A');pix=alpha.load();iw,ih=im.size
 # Refuse soft fades: only a narrow anti-aliasing band may be thresholded.
 lo=alpha.filter(ImageFilter.MinFilter(7));hi=alpha.filter(ImageFilter.MaxFilter(7))
 if any(8<a<247 and low>8 and high<247 for a,low,high in zip(alpha.getdata(),lo.getdata(),hi.getdata())):
  raise SystemExit('Soft mask detected; refusing to change its appearance.')
 active={};rects=[]
 for y in range(ih):
  nxt={};x=0
  while x<iw:
   if pix[x,y]<128:x+=1;continue
   left=x;x+=1
   while x<iw and pix[x,y]>=128:x+=1
   key=(left,x);run=active.pop(key,[left,x,y,y]);run[3]=y+1;nxt[key]=run
  rects.extend(active.values());active=nxt
 rects.extend(active.values())
 sx=float(box[2])/iw;sy=float(box[3])/ih
 paths=''.join(f'{left*sx:.4f} {(ih-bottom)*sy:.4f} {(right-left)*sx:.4f} {(bottom-top)*sy:.4f} re\n' for left,right,top,bottom in rects)
 replacement=DecodedStreamObject();replacement.update({k:v for k,v in obj.items() if k not in ['/Filter','/DecodeParms','/Length']})
 replacement.set_data(f'q\n0 0 0 RG 0 0 0 rg\n{base} gs\n{reset} gs\n{paths}W n\n{content} Do\nQ\n'.encode())
 w._objects[obj.indirect_reference.idnum-1]=replacement.flate_encode()
w.write(output)
print(json.dumps({'masks':len(targets),'before':source.stat().st_size,'after':output.stat().st_size,'pages':len(w.pages)}))

work.cleanup()
