from __future__ import annotations
import tempfile,unittest
from pathlib import Path
from xlsx_artifact_tool import XlsxArtifactTool

class LifecycleTests(unittest.TestCase):
 def setUp(self):
  self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name);self.tool=XlsxArtifactTool(self.root/'work')
 def tearDown(self):self.temp.cleanup()
 def test_create_inspect_plan_apply_validate_exact_cell_and_preservation(self):
  source=self.root/'source.xlsx'
  created=self.tool.create({'sheets':[{'name':'Summary','cells':{'A1':{'value':'Total','style':'header'},'B1':{'formula':'=SUM(Data!B2:B3)','style':'currency'}},'freeze_panes':'A2'},{'name':'Data','cells':{'A1':{'value':'Item','style':'header'},'B1':{'value':'Amount','style':'header'},'A2':{'value':'A'},'B2':{'value':2,'style':'currency'},'A3':{'value':'B'},'B3':{'value':3,'style':'currency'}},'auto_filter':'A1:B3'}]},source)
  self.assertEqual(created['status'],'ok')
  snapshot=self.tool.inspect(source,view='region',sheet='Data',range_ref='A1:B3');self.assertEqual(snapshot['status'],'ok')
  target=next(x for x in snapshot['cells'] if x['coordinate']=='B2');self.assertTrue(target['id'].startswith('tx_'))
  planned=self.tool.plan(snapshot,{'operations':[{'type':'set_cell_value','target_id':target['id'],'value':5,'expected_kind':'value'}]});self.assertEqual(planned['status'],'ok')
  output=self.root/'output.xlsx';applied=self.tool.apply(source,planned['plan'],output);self.assertEqual(applied['status'],'ok');self.assertTrue(output.exists())
  after=self.tool.inspect(output,view='region',sheet='Data',range_ref='A1:B3');self.assertEqual(next(x for x in after['cells'] if x['coordinate']=='B2')['value'],5)
  summary=self.tool.inspect(output,view='region',sheet='Summary',range_ref='A1:B1');self.assertEqual(next(x for x in summary['cells'] if x['coordinate']=='B1')['formula'],'=SUM(Data!B2:B3)')
  report=self.tool.validate(output,before=source);self.assertEqual(report['status'],'valid');self.assertEqual(report['unexpected_changed_members'],[])
  self.assertEqual(self.tool.inspect(source,view='region',sheet='Data',range_ref='B2:B2')['cells'][0]['value'],2)

 def test_formula_clear_and_semantic_style_are_exact_and_preserve_other_cells(self):
  source=self.root/'source.xlsx';self.tool.create({'sheets':[{'name':'Data','cells':{'A1':{'value':'Amount','style':'header'},'A2':{'value':10,'style':'currency'},'B2':{'value':'remove','style':'text'},'C2':{'value':0.5}}}]},source)
  snapshot=self.tool.inspect(source,view='region',sheet='Data',range_ref='A1:C2');by_coordinate={x['coordinate']:x for x in snapshot['cells']}
  request={'operations':[{'type':'set_cell_formula','target_id':by_coordinate['C2']['id'],'formula':'=A2*2','expected_kind':'value'},{'type':'clear_cell','target_id':by_coordinate['B2']['id'],'expected_kind':'value'},{'type':'set_cell_style','target_id':by_coordinate['A2']['id'],'style':'percent'}]}
  planned=self.tool.plan(snapshot,request);self.assertEqual(planned['status'],'ok')
  output=self.root/'output.xlsx';applied=self.tool.apply(source,planned['plan'],output);self.assertEqual(applied['status'],'ok')
  after=self.tool.inspect(output,view='region',sheet='Data',range_ref='A1:C2');cells={x['coordinate']:x for x in after['cells']}
  self.assertEqual(cells['C2']['formula'],'=A2*2');self.assertNotIn('B2',cells);self.assertEqual(cells['A2']['number_format'],'0.00%');self.assertEqual(cells['A1']['value'],'Amount')
  self.assertEqual(applied['formula_recalculation'],'required')

 def test_create_preserves_declared_layout_and_inspect_reports_it(self):
  source=self.root/'layout.xlsx';created=self.tool.create({'sheets':[{'name':'Data','cells':{'A1':{'value':'Report','style':'header'},'A2':{'value':'A'},'B2':{'value':10,'style':'currency'}},'column_widths':{'A':24,'B':14},'row_heights':{'1':28},'freeze_panes':'A2','auto_filter':'A1:B2','merged_ranges':['A1:B1']},{'name':'Hidden','state':'hidden','cells':{'A1':{'value':'x'}}}]},source);self.assertEqual(created['status'],'ok')
  summary=self.tool.inspect(source,view='summary');self.assertEqual([x['name'] for x in summary['sheets']],['Data','Hidden']);data=summary['sheets'][0];self.assertEqual(data['merged_ranges'],['A1:B1']);self.assertEqual(data['column_widths'],{'A':24.0,'B':14.0});self.assertEqual(data['row_heights'],{'1':28.0});self.assertEqual(data['freeze_panes'],'A2');self.assertEqual(data['auto_filter'],'A1:B2');self.assertEqual(summary['sheets'][1]['state'],'hidden')
 def test_reorder_rows_moves_values_formulas_and_styles_as_bound_rows(self):
  source=self.root/'rows.xlsx';self.tool.create({'sheets':[{'name':'Data','cells':{'A1':{'value':'Item','style':'header'},'B1':{'value':'Total','style':'header'},'A2':{'value':'A'},'B2':{'formula':'=10*2','style':'currency'},'A3':{'value':'B'},'B3':{'formula':'=20*2','style':'currency'},'A4':{'value':'C'},'B4':{'formula':'=30*2','style':'currency'}}}]},source)
  snap=self.tool.inspect(source,view='region',sheet='Data',range_ref='A2:B4');self.assertEqual(len(snap['rows']),3)
  planned=self.tool.plan(snap,{'operations':[{'type':'reorder_rows','region_id':snap['region_id'],'row_ids':[snap['rows'][2]['id'],snap['rows'][0]['id'],snap['rows'][1]['id']]}]});self.assertEqual(planned['status'],'ok')
  output=self.root/'rows-out.xlsx';result=self.tool.apply(source,planned['plan'],output);self.assertEqual(result['status'],'ok')
  after=self.tool.inspect(output,view='region',sheet='Data',range_ref='A2:B4');self.assertEqual([[c.get('value',c.get('formula')) for c in r['cells']] for r in after['rows']],[['C','=30*2'],['A','=10*2'],['B','=20*2']]);self.assertTrue(all(r['cells'][1]['number_format']=='#,##0.00' for r in after['rows']))
  self.assertEqual(self.tool.inspect(source,view='region',sheet='Data',range_ref='A2:B4')['rows'][0]['cells'][0]['value'],'A')

 def test_summary_and_search_are_bounded_and_agent_friendly(self):
  source=self.root/'views.xlsx';self.tool.create({'sheets':[{'name':'Summary','cells':{'A1':{'value':'Dashboard','style':'header'}}},{'name':'Data','cells':{'A1':{'value':'Item','style':'header'},'B1':{'value':'Status','style':'header'},'A2':{'value':'Alpha'},'B2':{'value':'Pending'},'A3':{'value':'Beta'},'B3':{'value':'Done'}}}]},source)
  summary=self.tool.inspect(source,view='summary');self.assertEqual(summary['status'],'ok');self.assertEqual([x['name'] for x in summary['sheets']],['Summary','Data']);self.assertTrue(all('used_range' in x for x in summary['sheets']))
  search=self.tool.inspect(source,view='search',query='pending');self.assertEqual(search['status'],'ok');self.assertEqual(len(search['matches']),1);self.assertEqual(search['matches'][0]['coordinate'],'B2');self.assertEqual(search['matches'][0]['row_context'],['Alpha','Pending'])

 def test_append_rows_uses_exact_region_and_explicit_template_style(self):
  source=self.root/'append.xlsx';self.tool.create({'sheets':[{'name':'Data','cells':{'A1':{'value':'Item','style':'header'},'B1':{'value':'Amount','style':'header'},'A2':{'value':'A'},'B2':{'value':10,'style':'currency'}}}]},source)
  snap=self.tool.inspect(source,view='region',sheet='Data',range_ref='A1:B2');template=snap['rows'][1]
  request={'operations':[{'type':'append_rows','region_id':snap['region_id'],'copy_from_row_id':template['id'],'rows':[['B',20],['C',{'formula':'=10+20'}]]}]}
  planned=self.tool.plan(snap,request);self.assertEqual(planned['status'],'ok');out=self.root/'append-out.xlsx';result=self.tool.apply(source,planned['plan'],out);self.assertEqual(result['status'],'ok')
  after=self.tool.inspect(out,view='region',sheet='Data',range_ref='A1:B4');self.assertEqual([[c.get('value',c.get('formula')) for c in r['cells']] for r in after['rows'][2:]],[['B',20],['C','=10+20']]);self.assertTrue(all(r['cells'][1]['number_format']=='#,##0.00' for r in after['rows'][2:]))

 def test_transforms_compile_to_exact_primitives_and_apply(self):
  source=self.root/'transforms.xlsx';self.tool.create({'sheets':[{'name':'Data','cells':{'A1':{'value':'Item','style':'header'},'B1':{'value':'Qty','style':'header'},'C1':{'value':'Price','style':'header'},'D1':{'value':'Total','style':'header'},'A2':{'value':'B'},'B2':{'value':2},'C2':{'value':5},'A3':{'value':'A'},'C3':{'value':3}}}]},source)
  snap=self.tool.inspect(source,view='region',sheet='Data',range_ref='A2:D3');rows=snap['rows'];blank_b3=rows[1]['cells'][1];d2=rows[0]['cells'][3];d3=rows[1]['cells'][3]
  request={'transforms':[{'type':'fill_missing','target_ids':[blank_b3['id']],'value':4},{'type':'table_totals','rows':[{'quantity_id':rows[0]['cells'][1]['id'],'unit_price_id':rows[0]['cells'][2]['id'],'target_id':d2['id']},{'quantity_id':blank_b3['id'],'unit_price_id':rows[1]['cells'][2]['id'],'target_id':d3['id']}]},{'type':'sort_rows','region_id':snap['region_id'],'keys_by_row_id':{rows[0]['id']:'B',rows[1]['id']:'A'},'descending':False}]}
  planned=self.tool.plan(snap,request);self.assertEqual(planned['status'],'ok');self.assertEqual([x['type'] for x in planned['plan']['operations']],['set_cell_value','set_cell_value','set_cell_value','reorder_rows'])
  output=self.root/'transforms-out.xlsx';result=self.tool.apply(source,planned['plan'],output);self.assertEqual(result['status'],'ok')
  after=self.tool.inspect(output,view='region',sheet='Data',range_ref='A2:D3');values=[[c.get('value',c.get('formula')) for c in row['cells']] for row in after['rows']];self.assertEqual(values,[['A',4,3,12],['B',2,5,10]])

if __name__=='__main__':unittest.main()
