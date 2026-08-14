from __future__ import annotations
import json,subprocess,tempfile,unittest
from pathlib import Path
PY=Path('/workspace/.venv-docx-study/bin/python');ROOT=Path(__file__).resolve().parents[1]
class CliTests(unittest.TestCase):
 def test_json_cli_create_inspect_plan_apply_validate(self):
  with tempfile.TemporaryDirectory() as d:
   d=Path(d);model=d/'model.json';model.write_text(json.dumps({'blocks':[{'type':'paragraph','text':'Alpha'}]}));src=d/'source.docx'
   def run(*args):
    p=subprocess.run([str(PY),'-m','office_artifact_tool','--workspace',str(d/'work'),*map(str,args)],cwd=ROOT,env={'PYTHONPATH':str(ROOT)},text=True,capture_output=True);self.assertEqual(p.returncode,0,p.stderr);return json.loads(p.stdout)
   self.assertEqual(run('create','--model',model,'--output',src)['status'],'ok');snap=run('inspect','--source',src);self.assertEqual(snap['status'],'ok');snapshot=d/'snapshot.json';snapshot.write_text(json.dumps(snap));target=next(x for x in snap['elements'] if x['kind']=='paragraph');request=d/'request.json';request.write_text(json.dumps({'operations':[{'type':'replace_text','target_id':target['id'],'old':'Alpha','new':'Beta'}]}));planned=run('plan','--snapshot',snapshot,'--request',request);self.assertEqual(planned['status'],'ok');plan=d/'plan.json';plan.write_text(json.dumps(planned['plan']));out=d/'out.docx';self.assertEqual(run('apply','--source',src,'--plan',plan,'--output',out)['status'],'ok');self.assertEqual(run('validate','--source',out)['status'],'valid')
if __name__=='__main__':unittest.main()
