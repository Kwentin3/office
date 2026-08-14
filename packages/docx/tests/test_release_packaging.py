from __future__ import annotations
import importlib.util,tempfile,tarfile,unittest
from pathlib import Path
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('package_release',ROOT/'package_release.py');release=importlib.util.module_from_spec(spec);spec.loader.exec_module(release)

class ReleasePackagingTests(unittest.TestCase):
 def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)/'project';self.root.mkdir();(self.root/'keep.txt').write_text('keep')
 def tearDown(self):self.tmp.cleanup()
 def test_build_tree_is_excluded_from_release_payload(self):
  stale=self.root/'build/lib/pkg/stale.py';stale.parent.mkdir(parents=True);stale.write_text('stale')
  with patch.object(release,'ROOT',self.root):
   release.write_manifest();archive=self.root.parent/'release.tar.gz';release.build(archive)
  with tarfile.open(archive,'r:gz') as tar:names=tar.getnames()
  self.assertIn('project/keep.txt',names);self.assertFalse(any('/build/' in f'/{name}/' for name in names))
 def test_unsafe_member_fails_closed(self):
  external=self.root.parent/'external.txt';external.write_text('external');(self.root/'external-link').symlink_to(external);published=self.root.parent/'published.tar.gz'
  with patch.object(release,'ROOT',self.root),patch.object(release,'ARCHIVE',published):
   release.write_manifest();archive=self.root.parent/'unsafe.tar.gz';release.build(archive)
   with self.assertRaises(SystemExit):release.verify(archive)
   with self.assertRaises(SystemExit):release.main()
  self.assertFalse(published.exists())

if __name__=='__main__':unittest.main()
