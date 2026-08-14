from __future__ import annotations
import argparse,json,sys
from pathlib import Path
from . import DocxArtifactTool

def load(path):return json.loads(Path(path).read_text())
def main(argv=None):
 p=argparse.ArgumentParser(prog='office-artifact-tool');p.add_argument('--workspace',default='.office-artifact-work');sub=p.add_subparsers(dest='action',required=True)
 c=sub.add_parser('create');c.add_argument('--model',required=True);c.add_argument('--output',required=True)
 i=sub.add_parser('inspect');i.add_argument('--source',required=True);i.add_argument('--view',default='full',choices=['full','search']);i.add_argument('--query')
 n=sub.add_parser('plan');n.add_argument('--snapshot',required=True);n.add_argument('--request',required=True)
 a=sub.add_parser('apply');a.add_argument('--source',required=True);a.add_argument('--plan',required=True);a.add_argument('--output',required=True)
 v=sub.add_parser('validate');v.add_argument('--source',required=True);v.add_argument('--before');v.add_argument('--expectations')
 args=p.parse_args(argv);tool=DocxArtifactTool(args.workspace)
 if args.action=='create':result=tool.create(load(args.model),args.output)
 elif args.action=='inspect':result=tool.inspect(args.source,args.view,args.query)
 elif args.action=='plan':result=tool.plan(load(args.snapshot),load(args.request))
 elif args.action=='apply':result=tool.apply(args.source,load(args.plan),args.output)
 else:result=tool.validate(args.source,args.before,load(args.expectations) if args.expectations else None)
 print(json.dumps(result,ensure_ascii=False));return 0 if result.get('status') in {'ok','valid','refused'} else 1
if __name__=='__main__':raise SystemExit(main())
