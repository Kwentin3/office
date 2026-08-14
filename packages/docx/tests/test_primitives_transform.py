from __future__ import annotations
import tempfile,unittest
from pathlib import Path
from office_artifact_tool import DocxArtifactTool
class PrimitiveTransformTests(unittest.TestCase):
 def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name);self.tool=DocxArtifactTool(self.root/'work')
 def tearDown(self):self.tmp.cleanup()
 def source(self):
  p=self.root/'table.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'Anchor'},{'type':'table','rows':[['Item','Qty','Price','Total','Owner'],['B','1','7','',''],['A','2','5','','Ann'],['Grand','','','','']]}]},p);return p
 def test_six_primitives_apply_in_bounded_combinations(self):
  src=self.source();snap=self.tool.inspect(src);els=snap['elements'];anchor=next(x for x in els if x['kind']=='paragraph' and x['text']=='Anchor');cell=next(x for x in els if x['kind']=='cell' and x['text']=='B');row=next(x for x in els if x['kind']=='row' and x['cells'][0]=='A')
  req={'operations':[{'type':'insert_paragraph_after','target_id':anchor['id'],'text':'Added','copy_properties':True},{'type':'set_cell_text','target_id':cell['id'],'text':'Bee'},{'type':'clone_row_after','target_id':row['id'],'cell_texts':['C','3','2','6','']}]};plan=self.tool.plan(snap,req);self.assertEqual(plan['status'],'ok');out=self.root/'six.docx';self.assertEqual(self.tool.apply(src,plan['plan'],out)['status'],'ok')
  texts=[x.get('text') for x in self.tool.inspect(out)['elements']];self.assertIn('Added',texts);self.assertIn('Bee',texts)
 def test_transform_computation_emits_typed_plan_then_kernel_applies(self):
  src=self.source();snap=self.tool.inspect(src);table=next(x for x in snap['elements'] if x['kind']=='table');rows=[x for x in snap['elements'] if x['kind']=='row' and x['table_id']==table['id']]
  planned=self.tool.plan(snap,{'transform':{'type':'sort_rows','table_id':table['id'],'row_ids':[r['id'] for r in rows[1:-1]],'keys':{'B':7,'A':10},'descending':True,'prefix_row_ids':[rows[0]['id']],'suffix_row_ids':[rows[-1]['id']]}});self.assertEqual(planned['status'],'ok');self.assertEqual(planned['computed']['row_order'],[rows[2]['id'],rows[1]['id']])
  out=self.root/'sorted.docx';self.assertEqual(self.tool.apply(src,planned['plan'],out)['status'],'ok');table_after=next(x for x in self.tool.inspect(out)['elements'] if x['kind']=='table');self.assertEqual(table_after['rows'][1][0],'A')
 def test_sort_rows_accepts_unambiguous_keys_by_row_id(self):
  src=self.source();snap=self.tool.inspect(src);table=next(x for x in snap['elements'] if x['kind']=='table');rows=[x for x in snap['elements'] if x['kind']=='row' and x['table_id']==table['id']];request={'transform':{'type':'sort_rows','table_id':table['id'],'row_ids':[r['id'] for r in rows[1:-1]],'keys_by_row_id':{rows[1]['id']:7,rows[2]['id']:10},'descending':True,'prefix_row_ids':[rows[0]['id']],'suffix_row_ids':[rows[-1]['id']]}};planned=self.tool.plan(snap,request);self.assertEqual(planned['status'],'ok');self.assertEqual(planned['computed']['row_order'],[rows[2]['id'],rows[1]['id']])
 def test_conflicting_structural_plan_is_refused(self):
  src=self.source();snap=self.tool.inspect(src);row=next(x for x in snap['elements'] if x['kind']=='row' and x['cells'][0]=='A');cell=next(x for x in snap['elements'] if x['kind']=='cell' and x['row_id']==row['id'])
  r=self.tool.plan(snap,{'operations':[{'type':'delete_row','target_id':row['id']},{'type':'set_cell_text','target_id':cell['id'],'text':'X'}]});self.assertEqual((r['status'],r['reason']),('refused','unsafe_plan'))
if __name__=='__main__':unittest.main()
