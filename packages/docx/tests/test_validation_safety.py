from __future__ import annotations
import hashlib,json,stat,tempfile,unittest,zipfile
from pathlib import Path
from lxml import etree
from office_artifact_tool import DocxArtifactTool
from office_artifact_tool.core.transaction import atomic_candidate
class ValidationSafetyTests(unittest.TestCase):
 def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name);self.tool=DocxArtifactTool(self.root/'work')
 def tearDown(self):self.tmp.cleanup()
 def test_invalid_package_and_source_equals_output_are_refused(self):
  bad=self.root/'bad.docx';bad.write_bytes(b'not zip');self.assertEqual(self.tool.inspect(bad)['reason'],'validation_failure')
  src=self.root/'a.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);snap=self.tool.inspect(src);target=next(x for x in snap['elements'] if x['kind']=='paragraph');plan=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'A','new':'B'}]})['plan'];r=self.tool.apply(src,plan,src);self.assertEqual(r['reason'],'unsafe_plan')
 def test_relationship_and_content_type_validation(self):
  src=self.root/'a.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);report=self.tool.validate(src);self.assertEqual(report['status'],'valid');self.assertTrue(report['relationships_valid']);self.assertTrue(report['content_types_valid'])
 def test_directory_entries_do_not_require_content_types(self):
  src=self.root/'a.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);with_dirs=self.root/'with-dirs.docx'
  with zipfile.ZipFile(src) as zin,zipfile.ZipFile(with_dirs,'w') as zout:
   for directory in ('word/','word/_rels/','_rels/','docProps/'):zout.writestr(zipfile.ZipInfo(directory),b'')
   for info in zin.infolist():zout.writestr(info,zin.read(info.filename))
  report=self.tool.validate(with_dirs);self.assertEqual(report['status'],'valid');self.assertTrue(report['content_types_valid']);self.assertEqual(self.tool.inspect(with_dirs)['status'],'ok')
 def test_set_cell_text_encodes_newlines_as_word_breaks(self):
  src=self.root/'multiline.docx';self.tool.create({'blocks':[{'type':'table','rows':[['Old']]}]},src);snap=self.tool.inspect(src);cell=next(x for x in snap['elements'] if x['kind']=='cell');plan=self.tool.plan(snap,{'operations':[{'type':'set_cell_text','target_id':cell['id'],'text':'Line 1\nLine 2'}]})['plan'];out=self.root/'multiline-out.docx';result=self.tool.apply(src,plan,out);self.assertEqual(result['status'],'ok')
  with zipfile.ZipFile(out) as archive:
   root=etree.fromstring(archive.read('word/document.xml'));self.assertEqual(len(root.xpath('.//w:br',namespaces={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})),1);self.assertFalse(any('\n' in (x.text or '') for x in root.xpath('.//w:t',namespaces={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})))
 def test_collateral_report_lists_unchanged_members(self):
  src=self.root/'a.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);snap=self.tool.inspect(src);target=next(x for x in snap['elements'] if x['kind']=='paragraph');plan=self.tool.plan(snap,{'operations':[{'type':'replace_text','target_id':target['id'],'old':'A','new':'B'}]})['plan'];out=self.root/'b.docx';result=self.tool.apply(src,plan,out);report=result['validation'];self.assertIn('word/document.xml',report['changed_members']);self.assertIn('word/styles.xml',report['unchanged_members']);self.assertFalse(report['unexpected_changed_members'])
 def test_atomic_candidate_keeps_exclusive_private_inode(self):
  output=self.root/'atomic.docx';observed={}
  def build(candidate):
   observed['exists']=candidate.exists();observed['mode']=candidate.stat().st_mode & 0o777;candidate.write_bytes(b'ok')
  report=atomic_candidate(output,build,lambda p:{'status':'valid'});self.assertEqual(report['status'],'valid');self.assertTrue(observed['exists']);self.assertEqual(observed['mode'] & 0o077,0);self.assertEqual(output.read_bytes(),b'ok')
 def test_corrupt_relationship_target_is_invalid(self):
  src=self.root/'a.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},src);bad=self.root/'bad.docx'
  with zipfile.ZipFile(src) as zin,zipfile.ZipFile(bad,'w') as zout:
   for info in zin.infolist():
    data=zin.read(info.filename)
    if info.filename=='word/_rels/document.xml.rels':data=data.replace(b'Target="styles.xml"',b'Target="missing.xml"')
    zout.writestr(info,data)
  self.assertEqual(self.tool.validate(bad)['status'],'invalid')
 def test_duplicate_traversal_and_nonregular_members_are_refused(self):
  valid=self.root/'valid.docx';self.tool.create({'blocks':[{'type':'paragraph','text':'A'}]},valid)
  cases=[('duplicate','word/document.xml',None),('traversal','../escape.xml',None),('symlink','word/styles.xml',stat.S_IFLNK|0o777),('fifo','word/styles.xml',stat.S_IFIFO|0o600)]
  for label,member,mode in cases:
   bad=self.root/f'unsafe-{label}.docx'
   with zipfile.ZipFile(valid) as source,zipfile.ZipFile(bad,'w') as output:
    for info in source.infolist():
     if mode is not None and info.filename==member:continue
     output.writestr(info,source.read(info.filename))
    info=zipfile.ZipInfo(member)
    if mode is not None:info.external_attr=mode<<16
    payload=source.read(member) if mode is not None else b'x'
    output.writestr(info,payload)
   self.assertEqual(self.tool.inspect(bad)['reason'],'validation_failure')
 def test_apply_rechecks_conflicts_even_for_direct_forged_plan(self):
  src=self.root/'a.docx';self.tool.create({'blocks':[{'type':'table','rows':[['H'],['A'],['B']]}]},src);snap=self.tool.inspect(src);row=next(x for x in snap['elements'] if x['kind']=='row' and x['cells']==['A']);cell=next(x for x in snap['elements'] if x['kind']=='cell' and x['row_id']==row['id'])
  plan={'schema':1,'source_sha256':snap['source_sha256'],'snapshot_sha256':snap['snapshot_sha256'],'operations':[{'type':'delete_row','target_id':row['id']},{'type':'set_cell_text','target_id':cell['id'],'text':'X'}]};plan['plan_sha256']=hashlib.sha256(json.dumps(plan,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest();out=self.root/'out.docx'
  result=self.tool.apply(src,plan,out);self.assertEqual((result['status'],result['reason']),('refused','unsafe_plan'));self.assertFalse(out.exists())
if __name__=='__main__':unittest.main()
