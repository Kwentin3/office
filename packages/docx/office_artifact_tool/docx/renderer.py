from __future__ import annotations
from io import BytesIO
from pathlib import Path
from typing import Any
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches
from ..core.errors import ArtifactError

_ALLOWED={'heading','paragraph','numbered_list','bulleted_list','table'}

def render(model:dict[str,Any],output:Path)->None:
    if not isinstance(model,dict) or not isinstance(model.get('blocks'),list):raise ArtifactError('validation_failure','model.blocks must be a list')
    doc=Document();meta=model.get('metadata',{})
    if not isinstance(meta,dict):raise ArtifactError('validation_failure','metadata must be an object')
    for key in ('title','subject','author','keywords','comments'):
        if key in meta and isinstance(meta[key],str):setattr(doc.core_properties,key,meta[key])
    for block in model['blocks']:
        if not isinstance(block,dict) or block.get('type') not in _ALLOWED:raise ArtifactError('unsupported_capability','unsupported create block')
        kind=block['type']
        if kind=='heading':
            text=block.get('text');level=block.get('level',1)
            if not isinstance(text,str) or not isinstance(level,int) or not 1<=level<=9:raise ArtifactError('validation_failure','invalid heading')
            doc.add_heading(text,level=level)
        elif kind=='paragraph':
            text=block.get('text');style=block.get('style')
            if not isinstance(text,str) or style is not None and not isinstance(style,str):raise ArtifactError('validation_failure','invalid paragraph')
            try:doc.add_paragraph(text,style=style)
            except KeyError as e:raise ArtifactError('validation_failure','unknown paragraph style') from e
        elif kind in {'numbered_list','bulleted_list'}:
            items=block.get('items');style='List Number' if kind=='numbered_list' else 'List Bullet'
            if not isinstance(items,list) or not items or not all(isinstance(x,str) for x in items):raise ArtifactError('validation_failure','invalid list')
            for item in items:doc.add_paragraph(item,style=style)
        elif kind=='table':
            rows=block.get('rows');style=block.get('style','Table Grid')
            if not isinstance(rows,list) or not rows or not all(isinstance(r,list) and r for r in rows):raise ArtifactError('validation_failure','invalid table')
            width=max(map(len,rows))
            table=doc.add_table(rows=len(rows),cols=width)
            try:table.style=style
            except KeyError as e:raise ArtifactError('validation_failure','unknown table style') from e
            for r,row in enumerate(rows):
                for c,value in enumerate(row):table.cell(r,c).text=str(value)
    doc.save(output)
