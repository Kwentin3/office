from __future__ import annotations
import math
from typing import Any
from .errors import ArtifactError
from .hashes import object_sha256

PRIMITIVES={'replace_text','insert_paragraph_after','set_cell_text','clone_row_after','reorder_rows','delete_row'}

def _refuse(reason,details):return {'status':'refused','reason':reason,'details':details}
def build_plan(snapshot:dict[str,Any],request:dict[str,Any])->dict[str,Any]:
    if not isinstance(snapshot,dict) or not isinstance(request,dict):return _refuse('validation_failure','invalid snapshot or request')
    elements={x['id']:x for x in snapshot.get('elements',[]) if isinstance(x,dict) and isinstance(x.get('id'),str)};ops=[];computed={}
    if 'operations' in request:ops=request['operations'] if isinstance(request['operations'],list) else []
    elif 'intents' in request:
        if not isinstance(request['intents'],list):return _refuse('validation_failure','intents must be a list')
        for intent in request['intents']:
            selector=intent.get('selector',{});operation=intent.get('operation',{})
            if set(selector) != {'kind','text'}:return _refuse('ambiguous_target','selector requires exact kind and text')
            matches=[x for x in elements.values() if x.get('kind')==selector['kind'] and x.get('text')==selector['text']]
            if len(matches)!=1:return _refuse('ambiguous_target',f'{len(matches)} targets matched')
            ops.append({**operation,'target_id':matches[0]['id']})
    elif 'transform' in request:
        transform=request['transform']
        if not isinstance(transform,dict):return _refuse('validation_failure','invalid transform')
        kind=transform.get('type')
        if kind=='sort_rows':
            row_ids=transform.get('row_ids',[]);keys_by_row_id=transform.get('keys_by_row_id')
            if not isinstance(row_ids,list) or set(row_ids)-set(elements) or not all(elements[x].get('kind')=='row' for x in row_ids):return _refuse('ambiguous_target','unknown transform row')
            if keys_by_row_id is not None:
                if not isinstance(keys_by_row_id,dict) or set(keys_by_row_id)!=set(row_ids):return _refuse('validation_failure','keys_by_row_id must match row_ids exactly')
                try:ordered=sorted(row_ids,key=lambda x:keys_by_row_id[x],reverse=bool(transform.get('descending')))
                except Exception:return _refuse('validation_failure','sort key invalid')
            else:
                keys=transform.get('keys',{})
                try:ordered=sorted(row_ids,key=lambda x:keys[elements[x]['cells'][0]],reverse=bool(transform.get('descending')))
                except Exception:return _refuse('validation_failure','sort key missing')
            all_rows=list(transform.get('prefix_row_ids',[]))+ordered+list(transform.get('suffix_row_ids',[]));ops=[{'type':'reorder_rows','table_id':transform.get('table_id'),'row_ids':all_rows}];computed={'row_order':ordered}
        elif kind=='table_totals':
            totals=[]
            for row in transform.get('rows',[]):
                total=row['quantity']*row['unit_price']
                if not math.isfinite(total):return _refuse('validation_failure','non-finite total')
                totals.append(total);ops.append({'type':'set_cell_text','target_id':row['total_target_id'],'text':str(total)})
            grand=sum(totals)
            if not math.isfinite(grand):return _refuse('validation_failure','non-finite grand total')
            ops.append({'type':'set_cell_text','target_id':transform['grand_total_target_id'],'text':str(grand)});computed={'line_totals':totals,'grand_total':grand}
        elif kind=='fill_missing':
            for item in transform.get('items',[]):
                if item.get('value') in {'',None}:ops.append({'type':'set_cell_text','target_id':item['target_id'],'text':str(transform['replacement'])})
            computed={'matched':len(ops)}
        elif kind=='bulk_replace':
            ops=[{'type':'set_cell_text','target_id':x['target_id'],'text':str(x['value'])} for x in transform.get('items',[])];computed={'matched':len(ops)}
        else:return _refuse('unsupported_capability','unsupported transform')
    else:return _refuse('validation_failure','request must contain operations, intents or transform')
    if not ops or not all(isinstance(x,dict) and x.get('type') in PRIMITIVES for x in ops):return _refuse('unsupported_capability','unsupported or empty operation list')
    deleted={x.get('target_id') for x in ops if x['type']=='delete_row'}
    structural=[];anchors=[]
    for op in ops:
        if op['type']=='reorder_rows':structural.append(op.get('table_id'))
        elif op['type']=='delete_row':
            item=elements.get(op.get('target_id'),{});structural.append(item.get('table_id'))
        if op['type'] in {'insert_paragraph_after','clone_row_after'}:anchors.append(op.get('target_id'))
    if len(anchors)!=len(set(anchors)):return _refuse('unsafe_plan','multiple insertions after one anchor')
    if len([x for x in structural if x])!=len(set(x for x in structural if x)):return _refuse('unsafe_plan','multiple structural edits in one table')
    for op in ops:
        target=op.get('target_id')
        if target in deleted and op['type']!='delete_row':return _refuse('unsafe_plan','deleted target reused')
        item=elements.get(target,{})
        if item.get('row_id') in deleted:return _refuse('unsafe_plan','target inside deleted row')
    plan={'schema':1,'source_sha256':snapshot['source_sha256'],'snapshot_sha256':snapshot['snapshot_sha256'],'operations':ops};plan['plan_sha256']=object_sha256(plan)
    return {'status':'ok','plan':plan,'computed':computed}
