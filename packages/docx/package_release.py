from __future__ import annotations
import gzip,hashlib,json,os,tarfile,tempfile
from pathlib import Path,PurePosixPath
ROOT=Path(__file__).resolve().parent
ARCHIVE=ROOT.parent/'office-artifact-tool-final.tar.gz'
EXCLUDED_NAMES={'__pycache__','.pytest_cache','.git','.mypy_cache','.ruff_cache','work','build','dist','.eggs'}
EXCLUDED_SUFFIXES={'.pyc','.pyo','.log'}
GENERATED={'FILE_MANIFEST.json','ARCHIVE_VERIFICATION.json'}

def sha(path:Path)->str:
 digest=hashlib.sha256()
 with path.open('rb') as stream:
  for block in iter(lambda:stream.read(1024*1024),b''):digest.update(block)
 return digest.hexdigest()
def include(path:Path)->bool:
 rel=path.relative_to(ROOT)
 if any(part in EXCLUDED_NAMES for part in rel.parts) or path.suffix in EXCLUDED_SUFFIXES or rel.as_posix() in GENERATED:return False
 if 'agent/evaluation/model-runs' in rel.as_posix() and path.suffix=='.docx':return False
 return path.is_file()
def files():return sorted((p for p in ROOT.rglob('*') if include(p)),key=lambda p:p.relative_to(ROOT).as_posix())
def write_manifest():
 records=[{'path':p.relative_to(ROOT).as_posix(),'bytes':p.stat().st_size,'sha256':sha(p)} for p in files()]
 (ROOT/'FILE_MANIFEST.json').write_text(json.dumps({'schema':1,'scope':'all packaged files except FILE_MANIFEST.json and ARCHIVE_VERIFICATION.json','files':records},indent=2)+'\n')
def build(path:Path):
 payload=files()+[ROOT/'FILE_MANIFEST.json']
 with path.open('wb') as raw,gzip.GzipFile(filename='',mode='wb',fileobj=raw,mtime=0) as gz,tarfile.open(fileobj=gz,mode='w',format=tarfile.PAX_FORMAT) as tar:
  for source in sorted(payload,key=lambda p:p.relative_to(ROOT).as_posix()):
   arc=PurePosixPath(ROOT.name)/source.relative_to(ROOT)
   info=tar.gettarinfo(str(source),str(arc));info.uid=0;info.gid=0;info.uname='';info.gname='';info.mtime=0;info.mode=0o644
   with source.open('rb') as stream:tar.addfile(info,stream)
def verify(path:Path):
 with tarfile.open(path,'r:gz') as tar:
  members=tar.getmembers();names=[x.name for x in members];unsafe=[x.name for x in members if x.name.startswith('/') or '..' in PurePosixPath(x.name).parts or x.issym() or x.islnk() or not x.isfile()]
 result={'sha256':sha(path),'bytes':path.stat().st_size,'members':len(members),'duplicate_members':len(names)-len(set(names)),'unsafe_members':unsafe}
 if result['duplicate_members'] or unsafe:raise SystemExit(f'unsafe archive: {result}')
 return result
def main():
 write_manifest()
 with tempfile.TemporaryDirectory(dir=ROOT.parent) as temp:
  one=Path(temp)/'one.tar.gz';two=Path(temp)/'two.tar.gz';build(one);build(two)
  if sha(one)!=sha(two):raise SystemExit('nondeterministic archive')
  verified=verify(one)
  os.replace(one,ARCHIVE)
 result={**verified,'deterministic_double_build':True,'archive':str(ARCHIVE)}
 (ROOT/'ARCHIVE_VERIFICATION.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps(result))
if __name__=='__main__':main()
