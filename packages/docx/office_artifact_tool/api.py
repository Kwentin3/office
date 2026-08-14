from __future__ import annotations
import json,os,tempfile,time
from pathlib import Path
from typing import Any
from .core.errors import ArtifactError,refusal
from .core.contracts import validate_create_model,validate_plan_request
from .core.hashes import file_sha256,object_sha256
from .core.plans import build_plan
from .core.semantic import semantic_postconditions
from .core.transaction import atomic_candidate
from .core.validation import validate_package
from .docx.inspect import inspect_docx
from .docx.mutation import mutate
from .docx.renderer import render

class DocxArtifactTool:
    def __init__(self,workspace:Path|str):
        self.workspace=Path(workspace);self.workspace.mkdir(parents=True,exist_ok=True)
    def create(self,model:dict[str,Any],output:Path|str)->dict[str,Any]:
        output=Path(output);start=time.perf_counter()
        try:
            if output.suffix.lower()!='.docx':raise ArtifactError('validation_failure','output must be .docx')
            validate_create_model(model)
            report=atomic_candidate(output,lambda p:render(model,p),lambda p:validate_package(p))
            return {'status':'ok','output':str(output),'sha256':file_sha256(output),'validation':report,'latency_ms':round((time.perf_counter()-start)*1000,3)}
        except ArtifactError as e:return refusal(e.reason,e.details)
        except Exception as e:return refusal('validation_failure',str(e))
    def inspect(self,source:Path|str,view:str='full',query:str|None=None)->dict[str,Any]:
        source=Path(source);start=time.perf_counter()
        try:
            snap=inspect_docx(source)
            if view=='search':
                if not isinstance(query,str) or not query:return refusal('validation_failure','query_required')
                q=query.casefold();matched=[x for x in snap['elements'] if q in json.dumps(x,ensure_ascii=False).casefold()];row_ids={x.get('row_id',x.get('id')) for x in matched if x.get('kind') in {'row','cell'}};table_ids={x.get('table_id') for x in matched if x.get('table_id')};snap['elements']=[x for x in snap['elements'] if x in matched or x.get('id') in row_ids or x.get('row_id') in row_ids or x.get('table_id') in table_ids and x.get('kind') in {'table','row','cell'}]
            elif view!='full':return refusal('unsupported_capability','unsupported inspect view')
            payload={**snap,'status':'ok','view':view};encoded=json.dumps(payload,ensure_ascii=False,separators=(',',':')).encode();payload['telemetry']={'tool_calls':1,'context_bytes':len(encoded),'approx_tokens':(len(encoded)+3)//4,'latency_ms':round((time.perf_counter()-start)*1000,3)};return payload
        except ArtifactError as e:return refusal(e.reason,e.details)
        except Exception:return refusal('validation_failure','invalid DOCX package')
    def plan(self,snapshot:dict[str,Any],request:dict[str,Any])->dict[str,Any]:
        try:
            if not isinstance(snapshot,dict):return refusal('validation_failure','invalid snapshot')
            check={k:v for k,v in snapshot.items() if k not in {'status','view','telemetry','snapshot_sha256'}}
            if snapshot.get('snapshot_sha256')!=object_sha256(check):return refusal('validation_failure','snapshot fingerprint mismatch')
            validate_plan_request(request)
            return build_plan(snapshot,request)
        except ArtifactError as e:return refusal(e.reason,e.details)
        except Exception:return refusal('validation_failure','invalid snapshot or request')
    def apply(self,source:Path|str,plan:dict[str,Any],output:Path|str)->dict[str,Any]:
        source=Path(source);output=Path(output);start=time.perf_counter();source_snapshot=None
        if output.suffix.lower()!='.docx':return refusal('validation_failure','output must be .docx')
        if source.resolve()==output.resolve():return refusal('unsafe_plan','source and output must differ')
        if not isinstance(plan,dict):return refusal('validation_failure','invalid plan')
        operations=plan.get('operations')
        if not isinstance(operations,list) or not operations:return refusal('validation_failure','invalid plan operations')
        if len(operations)>1000:return refusal('unsafe_plan','operation budget exceeded')
        try:validate_plan_request({'operations':operations})
        except ArtifactError as e:return refusal(e.reason,e.details)
        try:
            if not source.is_file():return refusal('validation_failure','source missing or not a file')
            fd,name=tempfile.mkstemp(prefix='.source-snapshot.',suffix='.docx',dir=self.workspace);source_snapshot=Path(name)
            with os.fdopen(fd,'wb') as target,source.open('rb') as original:
                for block in iter(lambda:original.read(1024*1024),b''):target.write(block)
        except Exception:
            if source_snapshot is not None:source_snapshot.unlink(missing_ok=True)
            return refusal('validation_failure','source snapshot failed')
        state={}
        try:
            source=source_snapshot
            source_hash=file_sha256(source)
            if plan.get('source_sha256')!=source_hash:return refusal('stale_snapshot','source fingerprint changed')
            if plan.get('plan_sha256')!=object_sha256({k:v for k,v in plan.items() if k!='plan_sha256'}):return refusal('validation_failure','plan fingerprint mismatch')
            before=self.inspect(source)
            def build(candidate):
                result=mutate(source,plan,candidate)
                if result.get('status')=='refused':raise ArtifactError(result['reason'],result['details'])
                state.update(result)
            def accept(candidate):
                report=validate_package(candidate,before=source,expected_changed_members=state.get('changed_parts',[]))
                if report.get('status')!='valid':return report
                after_candidate=self.inspect(candidate)
                semantic=semantic_postconditions(before,after_candidate,plan['operations'])
                report['semantic']=semantic
                if semantic['status']!='valid':report['status']='invalid';report['error']=semantic['error']
                return report
            report=atomic_candidate(output,build,accept)
            after=self.inspect(output)
            if after['status']!='ok':output.unlink(missing_ok=True);return refusal('validation_failure','semantic post-inspection failed')
            audit={'source_sha256':source_hash,'snapshot_sha256':plan['snapshot_sha256'],'plan_sha256':plan['plan_sha256'],'changed_parts':state['changed_parts'],'validation_status':report['status'],'output_sha256':file_sha256(output)}
            return {'status':'ok','output':str(output),'validation':report,'diff':{'changed_members':report['changed_members'],'changed_elements':state['changed_elements']},'audit':audit,'latency_ms':round((time.perf_counter()-start)*1000,3)}
        except ArtifactError as e:return refusal(e.reason,e.details)
        except Exception as e:return refusal('validation_failure',str(e))
        finally:
            if source_snapshot is not None:source_snapshot.unlink(missing_ok=True)
    def validate(self,path:Path|str,before:Path|str|None=None,expectations:dict[str,Any]|None=None)->dict[str,Any]:
        expected=(expectations or {}).get('changed_members')
        return validate_package(Path(path),Path(before) if before else None,expected)
