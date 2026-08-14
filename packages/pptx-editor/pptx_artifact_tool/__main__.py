from __future__ import annotations
import json,sys
from pathlib import Path
from .api import PptxArtifactTool,_refuse

def main():
 try:
  payload=json.load(sys.stdin)
  if not isinstance(payload,dict) or not isinstance(payload.get('action'),str):raise ValueError('validation_failure')
  tool=PptxArtifactTool(payload.get('workdir','.pptx-artifact-work'));action=payload['action']
  if action=='inspect':result=tool.inspect(Path(payload['source']),view=payload.get('view','summary'),slide_id=payload.get('slide_id'),query=payload.get('query'))
  elif action=='plan':result=tool.plan(payload['snapshot'],payload['request'])
  elif action=='apply':result=tool.apply(Path(payload['source']),payload['plan'],Path(payload['output']))
  elif action=='create':result=tool.create(Path(payload['template']),payload['model'],Path(payload['output']))
  elif action=='validate':result=tool.validate(Path(payload['source']),before=Path(payload['before']) if payload.get('before') else None)
  else:result=_refuse('unsupported_capability')
 except (ValueError,KeyError,TypeError,json.JSONDecodeError):result=_refuse('validation_failure')
 print(json.dumps(result,ensure_ascii=False,separators=(',',':')))
if __name__=='__main__':main()
