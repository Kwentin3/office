from __future__ import annotations
from typing import Any


def _signature(item:dict[str,Any]):
    kind=item.get('kind')
    if kind in {'paragraph','heading'}:return kind,item.get('story'),item.get('story_part'),item.get('style'),item.get('text')
    if kind=='cell':return kind,item.get('story'),item.get('table_id'),item.get('row_index'),item.get('cell_index'),item.get('text')
    if kind=='row':return kind,item.get('story'),item.get('table_id'),tuple(item.get('cells',[]))
    if kind=='table':return kind,item.get('story'),tuple(tuple(x) for x in item.get('rows',[]))
    return kind,

def _table_ordinal(elements:list[dict[str,Any]], item:dict[str,Any]) -> int | None:
    tables=[x for x in elements if x.get('kind')=='table' and x.get('story_part')==item.get('story_part')]
    ids=[x.get('id') for x in tables]
    try:return ids.index(item.get('table_id') if item.get('kind')!='table' else item.get('id'))
    except ValueError:return None


def _target_table(before:list[dict[str,Any]], after:list[dict[str,Any]], item:dict[str,Any]) -> dict[str,Any] | None:
    ordinal=_table_ordinal(before,item)
    tables=[x for x in after if x.get('kind')=='table' and x.get('story_part')==item.get('story_part')]
    return tables[ordinal] if ordinal is not None and ordinal<len(tables) else None


def _same_part_ordinal(elements:list[dict[str,Any]], item:dict[str,Any], kinds:set[str]) -> int | None:
    candidates=[x for x in elements if x.get('kind') in kinds and x.get('story_part')==item.get('story_part')]
    try:return [x.get('id') for x in candidates].index(item.get('id'))
    except ValueError:return None


def semantic_postconditions(before:dict[str,Any],after:dict[str,Any],operations:list[dict[str,Any]])->dict[str,Any]:
    old_by_id={x.get('id'):x for x in before.get('elements',[])};new=after.get('elements',[]);failures=[]
    for op in operations:
        kind=op['type'];old=old_by_id.get(op.get('target_id'))
        if kind=='replace_text':
            expected=old.get('text','').replace(op['old'],op['new'],1) if old else None
            ordinal=_same_part_ordinal(before.get('elements',[]),old,{old.get('kind')}) if old else None
            candidates=[x for x in new if x.get('kind')==old.get('kind') and x.get('story_part')==old.get('story_part')] if old else []
            target=candidates[ordinal] if ordinal is not None and ordinal<len(candidates) else None
            if not target or target.get('style')!=old.get('style') or target.get('text')!=expected:failures.append(kind)
        elif kind=='insert_paragraph_after':
            matches=[x for x in new if x.get('kind') in {'paragraph','heading'} and x.get('story')==old.get('story') and x.get('story_part')==old.get('story_part') and x.get('text')==op['text']] if old else []
            if len(matches)!=1:failures.append(kind)
        elif kind=='set_cell_text':
            old_table=old_by_id.get(old.get('table_id')) if old else None;new_table=_target_table(before.get('elements',[]),new,old) if old else None
            if old and old_table and new_table:
                rows=new_table.get('rows',[]);row_index=old.get('row_index');cell_index=old.get('cell_index')
                if not isinstance(row_index,int) or not isinstance(cell_index,int) or row_index>=len(rows) or cell_index>=len(rows[row_index]) or rows[row_index][cell_index]!=op['text']:failures.append(kind)
            else:failures.append(kind)
        elif kind=='clone_row_after':
            if not any(x.get('kind')=='row' and x.get('story')==old.get('story') and x.get('cells')==op['cell_texts'] for x in new):failures.append(kind)
        elif kind=='delete_row':
            if old and any(x.get('kind')=='row' and x.get('story')==old.get('story') and x.get('cells')==old.get('cells') for x in new):failures.append(kind)
        elif kind=='reorder_rows':
            old_rows=[old_by_id.get(x) for x in op['row_ids']]
            expected=[x.get('cells') for x in old_rows if x]
            tables=[x for x in new if x.get('kind')=='table' and x.get('rows')==expected]
            if len(tables)!=1:failures.append(kind)
    return {'status':'valid' if not failures else 'invalid','error':'semantic_postcondition_failed' if failures else None,'failed_operations':failures}
