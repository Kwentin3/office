from __future__ import annotations

import copy
import math
import re
from functools import lru_cache
from typing import Any

from docx import Document
from docx.enum.style import WD_STYLE_TYPE

from ..docx.presentation import MANAGED_PARAGRAPH_BLOCK_STYLES, MANAGED_TABLE_STYLES
from .errors import ArtifactError

_IDENTIFIER = re.compile(r'^[A-Za-z][A-Za-z0-9_.-]{0,63}$')
_CREATE_KEYS={'heading':{'type','block_id','level','text'},'paragraph':{'type','block_id','text','style'},'numbered_list':{'type','block_id','items'},'bulleted_list':{'type','block_id','items'},'table':{'type','block_id','style','rows'}}
_OPERATION_KEYS={'replace_text':{'type','target_id','old','new'},'insert_paragraph_after':{'type','target_id','text','copy_properties'},'set_cell_text':{'type','target_id','text'},'clone_row_after':{'type','target_id','cell_texts'},'reorder_rows':{'type','table_id','row_ids'},'delete_row':{'type','target_id'}}
_TRANSFORM_KEYS={'sort_rows':{'type','table_id','row_ids','keys_by_row_id','keys','descending','prefix_row_ids','suffix_row_ids'},'table_totals':{'type','rows','grand_total_target_id'},'fill_missing':{'type','items','replacement'},'bulk_replace':{'type','items'}}


def _is_xml_text(value: str) -> bool:
    return all(
        code in (0x09, 0x0A, 0x0D)
        or 0x20 <= code <= 0xD7FF
        or 0xE000 <= code <= 0xFFFD
        or 0x10000 <= code <= 0x10FFFF
        for code in map(ord, value)
    )


def _is_safely_renderable_table_scalar(value: Any) -> bool:
    if isinstance(value, str):
        return _is_xml_text(value)
    if isinstance(value, float):
        return math.isfinite(value)
    if isinstance(value, (int, bool)):
        try:
            str(value)
        except (ValueError, OverflowError):
            return False
        return True
    return False


@lru_cache(maxsize=1)
def _default_style_names() -> tuple[frozenset[str], frozenset[str]]:
    document = Document()
    paragraph = frozenset(style.name for style in document.styles if style.type == WD_STYLE_TYPE.PARAGRAPH)
    table = frozenset(style.name for style in document.styles if style.type == WD_STYLE_TYPE.TABLE)
    return paragraph, table


def closed_keys(value:dict[str,Any],allowed:set[str],required:set[str],name:str)->None:
    if set(value)-allowed or not required<=set(value):raise ArtifactError('validation_failure',f'invalid {name} fields')

def validate_create_model(model:Any)->dict[str,Any]:
    if not isinstance(model,dict):raise ArtifactError('validation_failure','create model must be object')
    closed_keys(model,{'document_id','metadata','blocks'},{'blocks'},'create model')
    if 'document_id' in model and (not isinstance(model['document_id'],str) or not _IDENTIFIER.fullmatch(model['document_id'])):raise ArtifactError('validation_failure','invalid document_id')
    if 'metadata' in model:
        if not isinstance(model['metadata'],dict) or set(model['metadata'])-{'title','subject','author','keywords','comments'} or not all(isinstance(x,str) and _is_xml_text(x) for x in model['metadata'].values()):raise ArtifactError('validation_failure','invalid metadata')
    if not isinstance(model['blocks'],list):raise ArtifactError('validation_failure','blocks must be list')
    if len(model['blocks'])>1000:raise ArtifactError('unsafe_plan','create block budget exceeded')
    block_ids:set[str]=set()
    for block in model['blocks']:
        if not isinstance(block,dict) or block.get('type') not in _CREATE_KEYS:raise ArtifactError('unsupported_capability','unsupported create block')
        kind=block['type'];required={'type','text'} if kind in {'heading','paragraph'} else {'type','items'} if kind.endswith('_list') else {'type','rows'}
        closed_keys(block,_CREATE_KEYS[kind],required,kind)
        if 'block_id' in block:
            block_id=block['block_id']
            if not isinstance(block_id,str) or not _IDENTIFIER.fullmatch(block_id) or block_id in block_ids:raise ArtifactError('validation_failure','block_id must be a unique stable identifier')
            block_ids.add(block_id)
        if kind == 'heading':
            level = block.get('level', 1)
            if (
                not isinstance(block['text'], str)
                or not _is_xml_text(block['text'])
                or isinstance(level, bool)
                or not isinstance(level, int)
                or not 1 <= level <= 9
            ):
                raise ArtifactError('validation_failure', 'invalid heading')
        elif kind == 'paragraph':
            style = block.get('style')
            paragraph_styles, _ = _default_style_names()
            if (
                not isinstance(block['text'], str)
                or not _is_xml_text(block['text'])
                or (style is not None and (not isinstance(style, str) or style not in paragraph_styles))
            ):
                raise ArtifactError('validation_failure', 'invalid paragraph')
            if style is not None and style not in MANAGED_PARAGRAPH_BLOCK_STYLES:
                raise ArtifactError(
                    'unsupported_capability',
                    'style is outside the managed DOCX presentation contract',
                )
        elif kind.endswith('_list'):
            if not isinstance(block['items'],list) or not block['items'] or len(block['items'])>10000 or not all(isinstance(x,str) and _is_xml_text(x) for x in block['items']):raise ArtifactError('validation_failure','invalid list items')
        elif kind=='table':
            rows=block['rows'];style=block.get('style','Table Grid');_, table_styles = _default_style_names()
            if not isinstance(style,str) or style not in table_styles:raise ArtifactError('validation_failure','invalid table style')
            if style not in MANAGED_TABLE_STYLES:raise ArtifactError('unsupported_capability','style is outside the managed DOCX presentation contract')
            if not isinstance(rows,list) or not rows or len(rows)>10000 or not all(isinstance(row,list) and row and len(row)<=1000 for row in rows):raise ArtifactError('validation_failure','invalid table rows')
            if sum(len(row) for row in rows)>100000 or not all(_is_safely_renderable_table_scalar(value) for row in rows for value in row):raise ArtifactError('validation_failure','table cells must be safely renderable finite XML-compatible scalar values')
    return copy.deepcopy(model)

def validate_plan_request(request:Any)->None:
    if not isinstance(request,dict) or len(set(request)&{'operations','intents','transform'})!=1 or set(request)-{'operations','intents','transform'}:raise ArtifactError('validation_failure','request must contain exactly one mode')
    if 'operations' in request:
        if not isinstance(request['operations'],list) or not request['operations']:raise ArtifactError('validation_failure','operations must be nonempty list')
        if len(request['operations'])>1000:raise ArtifactError('unsafe_plan','operation budget exceeded')
        for op in request['operations']:
            if not isinstance(op,dict) or op.get('type') not in _OPERATION_KEYS:raise ArtifactError('unsupported_capability','unsupported operation')
            kind=op['type'];required=_OPERATION_KEYS[kind]
            closed_keys(op,required,required,kind)
            if kind=='replace_text' and (not all(isinstance(op[key],str) for key in ('target_id','old','new')) or not op['old']):raise ArtifactError('validation_failure','invalid replace_text values')
            if kind=='insert_paragraph_after' and (not isinstance(op['target_id'],str) or not isinstance(op['text'],str) or not op['text'] or not isinstance(op['copy_properties'],bool)):raise ArtifactError('validation_failure','invalid insert values')
            if kind=='set_cell_text' and (not isinstance(op['target_id'],str) or not isinstance(op['text'],str)):raise ArtifactError('validation_failure','invalid cell values')
            if kind=='clone_row_after' and (not isinstance(op['target_id'],str) or not isinstance(op['cell_texts'],list) or not all(isinstance(value,str) for value in op['cell_texts'])):raise ArtifactError('validation_failure','invalid clone values')
            if kind=='reorder_rows' and (not isinstance(op['table_id'],str) or not isinstance(op['row_ids'],list) or not all(isinstance(value,str) for value in op['row_ids'])):raise ArtifactError('validation_failure','invalid reorder values')
            if kind=='delete_row' and not isinstance(op['target_id'],str):raise ArtifactError('validation_failure','invalid delete target')
    elif 'intents' in request:
        if not isinstance(request['intents'],list) or not request['intents']:raise ArtifactError('validation_failure','intents must be nonempty list')
        if len(request['intents'])>1000:raise ArtifactError('unsafe_plan','intent budget exceeded')
        for intent in request['intents']:
            if not isinstance(intent,dict):raise ArtifactError('validation_failure','intent must be object')
            closed_keys(intent,{'selector','operation'},{'selector','operation'},'intent')
            selector=intent['selector']
            if not isinstance(selector,dict):raise ArtifactError('validation_failure','selector must be object')
            closed_keys(selector,{'kind','text'},{'kind','text'},'selector')
            operation=intent['operation']
            if not isinstance(operation,dict) or operation.get('type') not in _OPERATION_KEYS:raise ArtifactError('unsupported_capability','unsupported operation')
            allowed=_OPERATION_KEYS[operation['type']]-{'target_id'};required=allowed
            closed_keys(operation,allowed,required,operation['type'])
    else:
        transform=request['transform']
        if not isinstance(transform,dict) or transform.get('type') not in _TRANSFORM_KEYS:raise ArtifactError('unsupported_capability','unsupported transform')
        kind=transform['type']
        if kind=='sort_rows':
            if ('keys_by_row_id' in transform)==('keys' in transform):raise ArtifactError('validation_failure','sort_rows requires exactly one key map')
            required={'type','table_id','row_ids','descending','prefix_row_ids','suffix_row_ids','keys_by_row_id' if 'keys_by_row_id' in transform else 'keys'}
        else:required=_TRANSFORM_KEYS[kind]
        closed_keys(transform,_TRANSFORM_KEYS[kind],required,kind)
        if kind=='sort_rows':
            row_ids=transform['row_ids'];prefix=transform['prefix_row_ids'];suffix=transform['suffix_row_ids'];keys=transform.get('keys_by_row_id',transform.get('keys'))
            if not all(isinstance(value,list) and all(isinstance(x,str) for x in value) for value in (row_ids,prefix,suffix)):raise ArtifactError('validation_failure','sort row IDs must be string lists')
            if not isinstance(transform['table_id'],str) or not isinstance(transform['descending'],bool) or not isinstance(keys,dict) or not all(isinstance(k,str) and isinstance(v,(str,int,float)) and not isinstance(v,bool) for k,v in keys.items()):raise ArtifactError('validation_failure','invalid sort inputs')
            if len(row_ids)>10000:raise ArtifactError('unsafe_plan','transform row budget exceeded')
        if kind in {'table_totals','fill_missing','bulk_replace'}:
            items=transform['rows'] if kind=='table_totals' else transform['items']
            if not isinstance(items,list) or len(items)>10000:raise ArtifactError('unsafe_plan','transform item budget exceeded')
            if not all(isinstance(item,dict) for item in items):raise ArtifactError('validation_failure','transform items must be objects')
            required_item={'quantity','unit_price','total_target_id'} if kind=='table_totals' else {'target_id','value'}
            if any(set(item)!=required_item for item in items):raise ArtifactError('validation_failure','invalid transform item fields')
            if kind=='table_totals' and not all(isinstance(item['quantity'],(int,float)) and not isinstance(item['quantity'],bool) and math.isfinite(item['quantity']) and isinstance(item['unit_price'],(int,float)) and not isinstance(item['unit_price'],bool) and math.isfinite(item['unit_price']) and isinstance(item['total_target_id'],str) for item in items):raise ArtifactError('validation_failure','invalid total inputs')
            if kind in {'fill_missing','bulk_replace'} and not all(isinstance(item['target_id'],str) and isinstance(item['value'],str) for item in items):raise ArtifactError('validation_failure','invalid transform item values')
            if kind=='fill_missing' and not isinstance(transform['replacement'],str):raise ArtifactError('validation_failure','replacement must be string')
