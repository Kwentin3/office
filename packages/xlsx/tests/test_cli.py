from __future__ import annotations
import json,subprocess,tempfile,unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1];PY='/workspace/.venv-docx-study/bin/python'
class CliTests(unittest.TestCase):
 def call(self,payload):
  proc=subprocess.run([PY,'-m','xlsx_artifact_tool'],input=json.dumps(payload,ensure_ascii=False),text=True,capture_output=True,cwd=ROOT,env={'PYTHONPATH':str(ROOT)});self.assertEqual(proc.returncode,0,proc.stderr);return json.loads(proc.stdout)
 def test_json_cli_create_inspect_plan_apply_validate(self):
  with tempfile.TemporaryDirectory() as d:
   root=Path(d);src=root/'source.xlsx';out=root/'output.xlsx'
   created=self.call({'action':'create','model':{'sheets':[{'name':'Data','cells':{'A1':{'value':'Item','style':'header'},'A2':{'value':'A'}}}]},'output':str(src)});self.assertEqual(created['status'],'ok')
   snap=self.call({'action':'inspect','source':str(src),'view':'region','sheet':'Data','range':'A1:A2'});target=next(x for x in snap['cells'] if x['coordinate']=='A2')
   planned=self.call({'action':'plan','snapshot':snap,'request':{'operations':[{'type':'set_cell_value','target_id':target['id'],'value':'B','expected_kind':'value'}]}});self.assertEqual(planned['status'],'ok')
   applied=self.call({'action':'apply','source':str(src),'plan':planned['plan'],'output':str(out)});self.assertEqual(applied['status'],'ok')
   valid=self.call({'action':'validate','source':str(out),'before':str(src)});self.assertEqual(valid['status'],'valid')
   self.assertEqual(self.call({'action':'inspect','source':str(out),'view':'region','sheet':'Data','range':'A2:A2'})['cells'][0]['value'],'B')
 def test_cli_malformed_request_is_typed_refusal(self):
  result=self.call({'action':'inspect'});self.assertEqual(result['status'],'refused');self.assertEqual(result['reason'],'validation_failure')
if __name__=='__main__':unittest.main()
