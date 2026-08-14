from __future__ import annotations
import hashlib,tempfile,unittest
from pathlib import Path
from office_artifact_tool import DocxArtifactTool

class MvpApiTests(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name);self.tool=DocxArtifactTool(self.root/'work')
 def tearDown(self):self.tmp.cleanup()
 def model(self):
  return {'metadata':{'title':'Quarterly note','author':'Test'},'blocks':[{'type':'heading','level':1,'text':'Quarterly Note'},{'type':'paragraph','text':'Payment is due in 30 days.','style':'Normal'},{'type':'numbered_list','items':['First','Second']},{'type':'table','rows':[['Item','Qty','Price','Total'],['A','2','5',''],['B','1','7',''],['Grand','','','']]}]}
 def test_create_inspect_snapshot_has_semantics_and_transaction_ids(self):
  out=self.root/'created.docx';created=self.tool.create(self.model(),out);self.assertEqual(created['status'],'ok');self.assertTrue(out.exists())
  snap=self.tool.inspect(out);self.assertEqual(snap['status'],'ok');self.assertEqual(snap['metadata']['title'],'Quarterly note');self.assertTrue(snap['sections']);self.assertIn('document',snap['stories'])
  self.assertTrue(any(x['kind']=='heading' and x['text']=='Quarterly Note' for x in snap['elements']));self.assertTrue(any(x['kind']=='cell' for x in snap['elements']))
  self.assertTrue(all(x['id'].startswith('tx_') for x in snap['elements']));self.assertNotIn(b'tx_',out.read_bytes())
 def test_plan_apply_validate_returns_diff_and_never_changes_source(self):
  src=self.root/'source.docx';self.tool.create(self.model(),src);source_hash=hashlib.sha256(src.read_bytes()).hexdigest();snap=self.tool.inspect(src)
  target=next(x for x in snap['elements'] if x['kind']=='paragraph' and '30 days' in x['text'])
  planned=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'30','new':'45'}]});self.assertEqual(planned['status'],'ok')
  out=self.root/'changed.docx';result=self.tool.apply(src,planned['plan'],out);self.assertEqual(result['status'],'ok');self.assertEqual(hashlib.sha256(src.read_bytes()).hexdigest(),source_hash)
  after=self.tool.inspect(out);self.assertTrue(any('45 days' in x.get('text','') for x in after['elements']));self.assertIn('word/document.xml',result['validation']['changed_members']);self.assertTrue(result['diff']['changed_elements'])
  self.assertEqual(result['audit']['source_sha256'],source_hash);self.assertEqual(result['audit']['output_sha256'],hashlib.sha256(out.read_bytes()).hexdigest())
 def test_ambiguous_selector_and_stale_snapshot_are_typed_refusals(self):
  model={'blocks':[{'type':'paragraph','text':'Duplicate'},{'type':'paragraph','text':'Duplicate'}]};src=self.root/'a.docx';self.tool.create(model,src);snap=self.tool.inspect(src)
  ambiguous=self.tool.plan(snap,{'intents':[{'selector':{'kind':'paragraph','text':'Duplicate'},'operation':{'type':'replace_text','old':'Duplicate','new':'Unique'}}]});self.assertEqual((ambiguous['status'],ambiguous['reason']),('refused','ambiguous_target'))
  src.write_bytes(src.read_bytes()+b'x');out=self.root/'b.docx';target=next(x for x in snap['elements'] if x['kind']=='paragraph');p=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'Duplicate','new':'X'}]})
  stale=self.tool.apply(src,p['plan'],out);self.assertEqual((stale['status'],stale['reason']),('refused','stale_snapshot'));self.assertFalse(out.exists())
 def test_semantic_postcondition_failure_blocks_publication(self):
  from unittest.mock import patch
  src=self.root/'source.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'Alpha'}]},src);snap=self.tool.inspect(src);target=next(x for x in snap['elements'] if x['kind']=='paragraph');plan=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'Alpha','new':'Beta'}]})['plan'];out=self.root/'out.docx'
  with patch('office_artifact_tool.api.semantic_postconditions',return_value={'status':'invalid','error':'semantic_postcondition_failed'}):result=self.tool.apply(src,plan,out)
  self.assertEqual((result['status'],result['reason']),('refused','validation_failure'));self.assertFalse(out.exists())
 def test_semantic_postconditions_are_target_specific_not_global_presence(self):
  from office_artifact_tool.core.semantic import semantic_postconditions
  before={'elements':[{'id':'p1','kind':'paragraph','text':'Alpha','story':'document','story_part':'word/document.xml'}]};after={'elements':[{'id':'new1','kind':'paragraph','text':'Alpha','story':'document','story_part':'word/document.xml'},{'id':'new2','kind':'paragraph','text':'Beta','story':'footer','story_part':'word/footer1.xml'}]}
  report=semantic_postconditions(before,after,[{'type':'replace_text','target_id':'p1','old':'Alpha','new':'Beta'}]);self.assertEqual(report['status'],'invalid')
 def test_replace_text_postcondition_is_bound_to_target_ordinal(self):
  from office_artifact_tool.core.semantic import semantic_postconditions
  before={'elements':[{'id':'p1','kind':'paragraph','text':'Alpha','style':'Normal','story':'document','story_part':'word/document.xml'},{'id':'p2','kind':'paragraph','text':'Gamma','style':'Normal','story':'document','story_part':'word/document.xml'}]}
  after={'elements':[{'id':'n1','kind':'paragraph','text':'Alpha','style':'Normal','story':'document','story_part':'word/document.xml'},{'id':'n2','kind':'paragraph','text':'Beta','style':'Normal','story':'document','story_part':'word/document.xml'}]}
  report=semantic_postconditions(before,after,[{'type':'replace_text','target_id':'p1','old':'Alpha','new':'Beta'}]);self.assertEqual(report['status'],'invalid')
 def test_table_semantic_postconditions_are_bound_to_target_table(self):
  from office_artifact_tool.core.semantic import semantic_postconditions
  before={'elements':[{'id':'t1','kind':'table','story':'document','story_part':'word/document.xml','rows':[['A']]},{'id':'c1','kind':'cell','story':'document','story_part':'word/document.xml','table_id':'t1','row_index':0,'cell_index':0,'text':'A'},{'id':'t2','kind':'table','story':'document','story_part':'word/document.xml','rows':[['X']]},{'id':'c2','kind':'cell','story':'document','story_part':'word/document.xml','table_id':'t2','row_index':0,'cell_index':0,'text':'X'}]}
  after={'elements':[{'id':'n1','kind':'table','story':'document','story_part':'word/document.xml','rows':[['A']]},{'id':'nc1','kind':'cell','story':'document','story_part':'word/document.xml','table_id':'n1','row_index':0,'cell_index':0,'text':'A'},{'id':'n2','kind':'table','story':'document','story_part':'word/document.xml','rows':[['B']]},{'id':'nc2','kind':'cell','story':'document','story_part':'word/document.xml','table_id':'n2','row_index':0,'cell_index':0,'text':'B'}]}
  report=semantic_postconditions(before,after,[{'type':'set_cell_text','target_id':'c1','text':'B'}]);self.assertEqual(report['status'],'invalid')
 def test_create_and_plan_enforce_closed_typed_contracts(self):
  bad_create=self.tool.create({'blocks':[{'type':'paragraph','text':'A','raw_xml':'<w:p/>'}]},self.root/'bad.docx');self.assertEqual((bad_create['status'],bad_create['reason']),('refused','validation_failure'));self.assertFalse((self.root/'bad.docx').exists())
  src=self.root/'source.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'Alpha'}]},src);snap=self.tool.inspect(src);target=next(x for x in snap['elements'] if x['kind']=='paragraph')
  mixed=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'Alpha','new':'Beta'}],'transform':{'type':'bulk_replace','items':[]}});self.assertEqual((mixed['status'],mixed['reason']),('refused','validation_failure'))
  extra=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'Alpha','new':'Beta','raw_xml':'x'}]});self.assertEqual((extra['status'],extra['reason']),('refused','validation_failure'))
 def test_malformed_nested_requests_and_non_scalar_cells_are_typed_refusals(self):
  bad_cell=self.tool.create({'blocks':[{'type':'table','rows':[['A',{'nested':'object'}]]}]},self.root/'bad-cell.docx');self.assertEqual((bad_cell['status'],bad_cell['reason']),('refused','validation_failure'));self.assertFalse((self.root/'bad-cell.docx').exists())
  src=self.root/'nested.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);snap=self.tool.inspect(src)
  cases=[
   {'intents':[{'selector':'not-an-object','operation':{'type':'replace_text','old':'A','new':'B'}}]},
   {'transform':{'type':'table_totals','rows':[{'quantity':2}],'grand_total_target_id':'tx_missing'}},
   {'transform':{'type':'fill_missing','items':[None],'replacement':'X'}},
   {'transform':{'type':'sort_rows','table_id':'tx_missing','row_ids':1,'keys_by_row_id':{},'descending':False,'prefix_row_ids':[],'suffix_row_ids':[]}},
   {'transform':{'type':'fill_missing','items':[{'target_id':'tx_missing','value':['not','scalar']}],'replacement':'X'}},
   {'transform':{'type':'bulk_replace','items':[{'target_id':'tx_missing','value':['not','scalar']}]}},
  ]
  for request in cases:
   with self.subTest(request=request):
    result=self.tool.plan(snap,request);self.assertEqual((result['status'],result['reason']),('refused','validation_failure'))
 def test_request_cardinality_budgets_refuse_before_work(self):
  src=self.root/'budget.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);snap=self.tool.inspect(src);target=next(x for x in snap['elements'] if x['kind']=='paragraph');op={'type':'replace_text','target_id':target['id'],'old':'A','new':'B'}
  result=self.tool.plan(snap,{'operations':[op for _ in range(1001)]});self.assertEqual((result['status'],result['reason']),('refused','unsafe_plan'))
  blocks=[{'type':'paragraph','text':'A'} for _ in range(1001)];result=self.tool.create({'blocks':blocks},self.root/'too-many.docx');self.assertEqual((result['status'],result['reason']),('refused','unsafe_plan'));self.assertFalse((self.root/'too-many.docx').exists())
 def test_create_schema_declares_runtime_budgets_and_accepts_schema_valid_ragged_table(self):
  import json
  schema=json.loads((Path(__file__).resolve().parents[1]/'agent/schemas/create.schema.json').read_text())
  self.assertEqual(schema['properties']['blocks']['maxItems'],1000);self.assertEqual(schema['$defs']['numbered']['properties']['items']['maxItems'],10000);self.assertEqual(schema['$defs']['table']['properties']['rows']['maxItems'],10000);self.assertEqual(schema['$defs']['table']['properties']['rows']['items']['maxItems'],1000)
  out=self.root/'ragged.docx';result=self.tool.create({'blocks':[{'type':'table','rows':[['A'],['B','C']]}]},out);self.assertEqual(result['status'],'ok');snap=self.tool.inspect(out);rows=[x for x in snap['elements'] if x['kind']=='row'];self.assertEqual(rows[0]['cells'],['A','']);self.assertEqual(rows[1]['cells'],['B','C'])
 def test_apply_malformed_value_is_typed_refusal_and_cleans_private_snapshot(self):
  from office_artifact_tool.core.hashes import object_sha256
  source=self.root/'malformed-value.docx';self.tool.create({'blocks':[{'type':'table','rows':[['A']]}]},source);snap=self.tool.inspect(source);target=next(x for x in snap['elements'] if x['kind']=='cell');plan=self.tool.plan(snap,{'operations':[{'type':'set_cell_text','target_id':target['id'],'text':'B'}]})['plan'];plan['operations'][0]['text']={1};out=self.root/'must-not-publish.docx';before=set((self.root/'work').glob('.source-snapshot.*.docx'))
  result=self.tool.apply(source,plan,out)
  self.assertEqual((result['status'],result['reason']),('refused','validation_failure'));self.assertFalse(out.exists());self.assertEqual(set((self.root/'work').glob('.source-snapshot.*.docx')),before)
 def test_apply_rechecks_operation_budget_for_forged_plan(self):
  from office_artifact_tool.core.hashes import object_sha256
  src=self.root/'forged-budget.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);snap=self.tool.inspect(src);target=next(x for x in snap['elements'] if x['kind']=='paragraph');base=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'A','new':'A'}]})['plan'];base['operations']=base['operations']*1001;base['plan_sha256']=object_sha256({k:v for k,v in base.items() if k!='plan_sha256'});out=self.root/'forged-budget-out.docx'
  result=self.tool.apply(src,base,out);self.assertEqual((result['status'],result['reason']),('refused','unsafe_plan'));self.assertFalse(out.exists())
  malformed=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'A','new':'A'}]})['plan'];malformed['operations'][0]['raw_xml']='x';malformed['plan_sha256']=object_sha256({k:v for k,v in malformed.items() if k!='plan_sha256'});result=self.tool.apply(src,malformed,out);self.assertEqual((result['status'],result['reason']),('refused','validation_failure'));self.assertFalse(out.exists())
 def test_public_methods_return_typed_results_for_malformed_paths(self):
  missing=self.root/'missing.docx';bad_output=self.root/'bad-output.docx'
  for result in (self.tool.inspect(missing),self.tool.apply(missing,{},bad_output)):
   with self.subTest(result=result):self.assertEqual((result['status'],result['reason']),('refused','validation_failure'))
  directory=self.root/'directory.docx';directory.mkdir();result=self.tool.apply(directory,{'operations':[{'type':'delete_row','target_id':'x'}]},bad_output);self.assertEqual((result['status'],result['reason']),('refused','validation_failure'))
  self.assertEqual(self.tool.validate(missing)['status'],'invalid')
 def test_apply_is_bound_to_private_source_snapshot(self):
  from unittest.mock import patch
  from office_artifact_tool import api as api_module
  from docx import Document
  source=self.root/'race.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'Stable'}]},source);snap=self.tool.inspect(source);target=next(x for x in snap['elements'] if x['kind']=='paragraph');plan=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'Stable','new':'Reviewed'}]})['plan'];out=self.root/'race-out.docx'
  real_hash=api_module.file_sha256;raced={'done':False}
  def race_after_hash(path):
   digest=real_hash(path)
   if Path(path)==source and not raced['done']:
    raced['done']=True;document=Document(source);document.paragraphs[0].text='RACED';document.save(source)
   return digest
  with patch('office_artifact_tool.api.file_sha256',side_effect=race_after_hash):result=self.tool.apply(source,plan,out)
  self.assertEqual(result['status'],'ok');self.assertEqual(Document(out).paragraphs[0].text,'Reviewed')
 def test_snapshot_metadata_cannot_self_authorize_a_plan(self):
  source=self.root/'snapshot-authority.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'Alpha'}]},source);snap=self.tool.inspect(source);target=next(x for x in snap['elements'] if x['kind']=='paragraph');forged=dict(snap);forged['source_sha256']='0'*64;forged['snapshot_sha256']='0'*64;result=self.tool.plan(forged,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'Alpha','new':'Beta'}]});self.assertEqual((result['status'],result['reason']),('refused','validation_failure'))
 def test_plan_malformed_snapshot_is_a_typed_refusal(self):
  result=self.tool.plan({}, {'operations':[{'type':'delete_row','target_id':'tx_missing'}]});self.assertEqual((result['status'],result['reason']),('refused','validation_failure'))
 def test_public_create_and_apply_refuse_non_docx_output(self):
  bad_create=self.root/'created.txt';result=self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},bad_create);self.assertEqual((result['status'],result['reason']),('refused','validation_failure'));self.assertFalse(bad_create.exists())
  source=self.root/'source.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},source);snap=self.tool.inspect(source);target=next(x for x in snap['elements'] if x['kind']=='paragraph');plan=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'A','new':'B'}]})['plan'];bad_apply=self.root/'applied.txt';result=self.tool.apply(source,plan,bad_apply);self.assertEqual((result['status'],result['reason']),('refused','validation_failure'));self.assertFalse(bad_apply.exists())
 def test_create_refuses_non_finite_table_numbers(self):
  for index,value in enumerate((float('nan'),float('inf'),float('-inf'))):
   output=self.root/f'non-finite-{index}.docx';result=self.tool.create({'blocks':[{'type':'table','rows':[['Value',value]]}]},output);self.assertEqual((result['status'],result['reason']),('refused','validation_failure'));self.assertFalse(output.exists())
 def test_table_totals_refuses_non_finite_inputs_and_results(self):
  source=self.root/'totals.docx';self.tool.create({'blocks':[{'type':'table','rows':[['Qty','Price','Total'],['1','1','']]}]},source);snap=self.tool.inspect(source);target=next(x for x in snap['elements'] if x['kind']=='cell' and x['row_index']==1 and x['cell_index']==2)
  for quantity,price in ((float('inf'),1),(1e308,1e308)):
   result=self.tool.plan(snap,{'transform':{'type':'table_totals','rows':[{'quantity':quantity,'unit_price':price,'total_target_id':target['id']}],'grand_total_target_id':target['id']}});self.assertEqual((result['status'],result['reason']),('refused','validation_failure'))
if __name__=='__main__':unittest.main()
