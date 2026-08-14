from __future__ import annotations
import hashlib,json,posixpath,zipfile
from pathlib import Path
from typing import Any
from lxml import etree
from docx import Document
from ..core.errors import ArtifactError
from ..core.hashes import file_sha256,object_sha256
from ..core.validation import validate_package

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';PR='http://schemas.openxmlformats.org/package/2006/relationships';NS={'w':W,'r':R,'pr':PR};DOCUMENT='word/document.xml'
STORY_TYPES={f'{R}/header':'header',f'{R}/footer':'footer',f'{R}/footnotes':'footnote',f'{R}/endnotes':'endnote'}

def _rels_name(part):
 d,n=posixpath.split(part);return posixpath.join(d,'_rels',n+'.rels')
def _relations(parts,part):
 name=_rels_name(part)
 if name not in parts:return {}
 root=etree.fromstring(parts[name]);return {x.get('Id'):{'type':x.get('Type'),'part':posixpath.normpath(posixpath.join(posixpath.dirname(part),x.get('Target')))} for x in root.xpath('./pr:Relationship',namespaces=NS) if x.get('TargetMode')!='External'}
def _stories(parts):
 result={DOCUMENT:{'type':'document','sections':[]}};rels=_relations(parts,DOCUMENT);root=etree.fromstring(parts[DOCUMENT]);sections=root.xpath('.//w:sectPr',namespaces=NS)
 for index,section in enumerate(sections):
  for ref in section.xpath('./w:headerReference | ./w:footerReference',namespaces=NS):
   rel=rels.get(ref.get(f'{{{R}}}id'))
   if rel and rel['type'] in STORY_TYPES:result.setdefault(rel['part'],{'type':STORY_TYPES[rel['type']],'sections':[]})['sections'].append(index)
 for rel in rels.values():
  if STORY_TYPES.get(rel['type']) in {'footnote','endnote'}:result.setdefault(rel['part'],{'type':STORY_TYPES[rel['type']],'sections':[]})
 return result
def _text(node):
 result=[]
 for item in node.xpath('.//w:t | .//w:br',namespaces=NS):result.append('\n' if item.tag==f'{{{W}}}br' else item.text or '')
 return ''.join(result)
def _style(p):
 v=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=NS);return v[0] if v else 'Normal'
def _id(source,part,kind,path,payload):return 'tx_'+hashlib.sha256(f'{source}|{part}|{kind}|{path}|{payload}'.encode()).hexdigest()[:24]
def inspect_docx(path:Path)->dict[str,Any]:
 report=validate_package(path)
 if report['status']!='valid':raise ArtifactError('validation_failure',report.get('error','invalid DOCX package'))
 source=file_sha256(path)
 with zipfile.ZipFile(path) as z:parts={x.filename:z.read(x.filename) for x in z.infolist()}
 stories_meta=_stories(parts);elements=[];stories={}
 for part,meta in sorted(stories_meta.items()):
  if part not in parts:continue
  root=etree.fromstring(parts[part]);story_elements=[]
  paragraphs=root.xpath('.//w:p[not(ancestor::w:tc)]',namespaces=NS)
  for i,p in enumerate(paragraphs):
   text=_text(p)
   if not text:continue
   style=_style(p);kind='heading' if style.lower().startswith('heading') else 'paragraph';ident=_id(source,part,kind,str(i),text)
   item={'id':ident,'kind':kind,'text':text,'style':style,'story':meta['type'],'story_part':part,'sections':sorted(set(meta['sections']))};elements.append(item);story_elements.append(ident)
  for ti,table in enumerate(root.xpath('.//w:tbl',namespaces=NS)):
   rows=table.xpath('./w:tr',namespaces=NS);matrix=[]
   for row in rows:matrix.append([_text(cell) for cell in row.xpath('./w:tc',namespaces=NS)])
   tid=_id(source,part,'table',str(ti),json.dumps(matrix,ensure_ascii=False));table_item={'id':tid,'kind':'table','rows':matrix,'row_count':len(rows),'story':meta['type'],'story_part':part,'sections':sorted(set(meta['sections']))};elements.append(table_item);story_elements.append(tid)
   for ri,(row,values) in enumerate(zip(rows,matrix)):
    rid=_id(source,part,'row',f'{ti}/{ri}','|'.join(values));r={'id':rid,'kind':'row','cells':values,'table_id':tid,'row_index':ri,'story':meta['type'],'story_part':part};elements.append(r)
    for ci,(cell,value) in enumerate(zip(row.xpath('./w:tc',namespaces=NS),values)):
     cid=_id(source,part,'cell',f'{ti}/{ri}/{ci}',value);elements.append({'id':cid,'kind':'cell','text':value,'table_id':tid,'row_id':rid,'row_index':ri,'cell_index':ci,'story':meta['type'],'story_part':part})
  stories[meta['type'] if meta['type']=='document' else f"{meta['type']}:{part}"]={'type':meta['type'],'part':part,'sections':sorted(set(meta['sections'])),'element_ids':story_elements}
 doc=Document(path);metadata={k:getattr(doc.core_properties,k) for k in ('title','subject','author','keywords','comments') if getattr(doc.core_properties,k,None)};sections=[{'index':i,'start_type':str(s.start_type)} for i,s in enumerate(doc.sections)];styles=sorted({x.get('style') for x in elements if x.get('style')})
 snapshot={'source_sha256':source,'document_id':'doc_'+source[:20],'metadata':metadata,'sections':sections,'stories':stories,'elements':elements,'styles':styles};snapshot['snapshot_sha256']=object_sha256(snapshot);return snapshot
