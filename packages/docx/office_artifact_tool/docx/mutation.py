from __future__ import annotations

import copy
import difflib
import json
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

from ..core.errors import ArtifactError
from ..core.hashes import file_sha256
from .inspect import NS, W, _id, _stories, _style, _text

_ALLOWED={'replace_text','insert_paragraph_after','set_cell_text','clone_row_after','reorder_rows','delete_row'}
def _read(path):
 with zipfile.ZipFile(path) as z:return z.infolist(),{x.filename:z.read(x.filename) for x in z.infolist()}
def _resolve(parts,source):
 roots={};objects={};stories=_stories(parts)
 for part in iter(stories):
  if part not in parts:continue
  root=etree.fromstring(parts[part]);roots[part]=root
  for i,p in enumerate(root.xpath('.//w:p[not(ancestor::w:tc)]',namespaces=NS)):
   text=_text(p)
   if text:objects[_id(source,part,'heading' if _style(p).lower().startswith('heading') else 'paragraph',str(i),text)]={'element':p,'kind':'heading' if _style(p).lower().startswith('heading') else 'paragraph','part':part,'text':text}
  for ti,table in enumerate(root.xpath('.//w:tbl',namespaces=NS)):
   rows=table.xpath('./w:tr',namespaces=NS);matrix=[[_text(c) for c in r.xpath('./w:tc',namespaces=NS)] for r in rows];tid=_id(source,part,'table',str(ti),json.dumps(matrix,ensure_ascii=False));objects[tid]={'element':table,'kind':'table','part':part,'rows':matrix}
   for ri,(row,vals) in enumerate(zip(rows,matrix)):
    rid=_id(source,part,'row',f'{ti}/{ri}','|'.join(vals));objects[rid]={'element':row,'kind':'row','part':part,'table_id':tid,'cells':vals}
    for ci,(cell,val) in enumerate(zip(row.xpath('./w:tc',namespaces=NS),vals)):
     cid=_id(source,part,'cell',f'{ti}/{ri}/{ci}',val);objects[cid]={'element':cell,'kind':'cell','part':part,'table_id':tid,'row_id':rid,'text':val}
 return roots,objects
def _replace(p,old,new):
 nodes=p.xpath('.//w:t',namespaces=NS);joined=''.join(x.text or '' for x in nodes)
 if joined.count(old)!=1:raise ArtifactError('validation_failure','old text cardinality')
 result=joined.replace(old,new,1);original=[node.text or '' for node in nodes];offsets=[];cursor=0
 for text in original:offsets.append((cursor,cursor+len(text)));cursor+=len(text)
 distributed=['' for _ in nodes]
 for kind,i1,i2,j1,j2 in difflib.SequenceMatcher(a=joined,b=result,autojunk=False).get_opcodes():
  if kind=='equal':
   for index,(start,end) in enumerate(offsets):
    left=max(start,i1);right=min(end,i2)
    if left<right:distributed[index]+=joined[left:right]
  elif kind in {'replace','insert'} and j1<j2:
   owner=next((index for index,(start,end) in enumerate(offsets) if start<=i1<end),len(nodes)-1)
   distributed[owner]+=result[j1:j2]
 for node,text in zip(nodes,distributed):node.text=text
def _set_cell(cell,text):
 paragraphs=cell.xpath('./w:p',namespaces=NS)
 nested_tables=cell.xpath('./w:tbl',namespaces=NS)
 if len(paragraphs)==1 and not nested_tables and paragraphs[0].xpath('.//w:t',namespaces=NS):
  old=_text(cell)
  if old:
   _replace(paragraphs[0],old,text)
   for node in list(paragraphs[0].xpath('.//w:t',namespaces=NS)):
    if '\n' not in (node.text or ''):continue
    run=node.getparent();pieces=(node.text or '').split('\n');node.text=pieces[0];position=run.index(node)
    for piece in pieces[1:]:
     position+=1;run.insert(position,etree.Element(f'{{{W}}}br'))
     position+=1;continuation=etree.Element(f'{{{W}}}t');continuation.text=piece;run.insert(position,continuation)
   return
 for child in list(cell):
  if child.tag!=f'{{{W}}}tcPr':cell.remove(child)
 p=etree.SubElement(cell,f'{{{W}}}p');r=etree.SubElement(p,f'{{{W}}}r')
 lines=text.split('\n')
 for index,line in enumerate(lines):
  if index:etree.SubElement(r,f'{{{W}}}br')
  t=etree.SubElement(r,f'{{{W}}}t');t.text=line

def mutate(source:Path,plan:dict[str,Any],candidate:Path)->dict:
 infos,parts=_read(source);roots,objects=_resolve(parts,file_sha256(source));prepared=[]
 if set(plan)!={'schema','source_sha256','snapshot_sha256','operations','plan_sha256'} or plan.get('schema')!=1:return {'status':'refused','reason':'validation_failure','details':'invalid plan envelope'}
 deleted={x.get('target_id') for x in plan['operations'] if isinstance(x,dict) and x.get('type')=='delete_row'}
 structural=[];anchors=[]
 for op in plan['operations']:
  if not isinstance(op,dict):raise ArtifactError('validation_failure','invalid operation')
  if op.get('type')=='reorder_rows':structural.append(op.get('table_id'))
  elif op.get('type')=='delete_row':structural.append(objects.get(op.get('target_id'),{}).get('table_id'))
  if op.get('type') in {'insert_paragraph_after','clone_row_after'}:anchors.append(op.get('target_id'))
  item=objects.get(op.get('target_id'),{})
  if op.get('target_id') in deleted and op.get('type')!='delete_row' or item.get('row_id') in deleted:raise ArtifactError('unsafe_plan','deleted target reused')
 if len(anchors)!=len(set(anchors)):raise ArtifactError('unsafe_plan','multiple insertions after one anchor')
 structural=[x for x in structural if x]
 if len(structural)!=len(set(structural)):raise ArtifactError('unsafe_plan','multiple structural edits in one table')
 for op in plan['operations']:
  kind=op.get('type')
  if kind not in _ALLOWED:raise ArtifactError('unsupported_capability')
  if kind=='reorder_rows':
   table=objects.get(op.get('table_id'));ids=op.get('row_ids')
   expected=[k for k,v in objects.items() if v.get('kind')=='row' and v.get('table_id')==op.get('table_id')]
   if not table or table['kind']!='table' or not isinstance(ids,list) or len(ids)!=len(expected) or set(ids)!=set(expected):raise ArtifactError('unsafe_plan','row_ids must be exact permutation')
   prepared.append((kind,table,[objects[x]['element'] for x in ids]));continue
  target=objects.get(op.get('target_id'))
  if not target:raise ArtifactError('ambiguous_target','unknown transaction target')
  if kind=='replace_text':
   if target['kind'] not in {'paragraph','heading'} or set(op)!={'type','target_id','old','new'} or not isinstance(op['old'],str) or not op['old'] or not isinstance(op['new'],str) or target['text'].count(op['old'])!=1:raise ArtifactError('validation_failure','invalid replace_text')
  elif kind=='insert_paragraph_after':
   if target['kind'] not in {'paragraph','heading'} or set(op)!={'type','target_id','text','copy_properties'} or not isinstance(op['text'],str) or not op['text'] or not isinstance(op['copy_properties'],bool):raise ArtifactError('validation_failure','invalid insert')
  elif kind=='set_cell_text':
   if target['kind']!='cell' or set(op)!={'type','target_id','text'} or not isinstance(op['text'],str):raise ArtifactError('validation_failure','invalid cell edit')
  elif kind=='clone_row_after':
   vals=op.get('cell_texts')
   if target['kind']!='row' or set(op)!={'type','target_id','cell_texts'} or not isinstance(vals,list) or not all(isinstance(x,str) for x in vals) or len(vals)!=len(target['element'].xpath('./w:tc',namespaces=NS)):raise ArtifactError('validation_failure','invalid row clone')
  elif kind=='delete_row' and (target['kind']!='row' or set(op)!={'type','target_id'}):raise ArtifactError('validation_failure','invalid row deletion')
  prepared.append((kind,target,op))
 changed_elements=[];changed_parts=set()
 for kind,target,data in prepared:
  if kind=='replace_text':_replace(target['element'],data['old'],data['new'])
  elif kind=='insert_paragraph_after':
   p=etree.Element(f'{{{W}}}p');props=target['element'].find(f'{{{W}}}pPr')
   if data['copy_properties'] and props is not None:p.append(copy.deepcopy(props))
   r=etree.SubElement(p,f'{{{W}}}r');t=etree.SubElement(r,f'{{{W}}}t');t.text=data['text'];target['element'].addnext(p)
  elif kind=='set_cell_text':_set_cell(target['element'],data['text'])
  elif kind=='clone_row_after':
   row=copy.deepcopy(target['element'])
   for cell,text in zip(row.xpath('./w:tc',namespaces=NS),data['cell_texts']):_set_cell(cell,text)
   target['element'].addnext(row)
  elif kind=='delete_row':target['element'].getparent().remove(target['element'])
  elif kind=='reorder_rows':
   for row in data:target['element'].append(row)
  changed_parts.add(target['part']);changed_elements.append({'operation':kind,'target_id':data.get('target_id') if isinstance(data,dict) else None})
 for part in changed_parts:parts[part]=etree.tostring(roots[part],xml_declaration=True,encoding='UTF-8',standalone=True)
 with zipfile.ZipFile(candidate,'w') as out:
  for info in infos:out.writestr(copy.copy(info),parts[info.filename])
 return {'changed_parts':sorted(changed_parts),'changed_elements':changed_elements}
