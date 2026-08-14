from __future__ import annotations
import hashlib,posixpath,stat,zipfile
from pathlib import Path,PurePosixPath
from lxml import etree
from .errors import ArtifactError

PR='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'

def _resolve(source_part:str,target:str)->str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part),target))

def _source_for_rels(name:str)->str:
    directory,filename=posixpath.split(name)
    if directory=='_rels':return filename[:-5]
    parent=posixpath.dirname(directory);return posixpath.join(parent,filename[:-5])

def validate_package(path:Path,before:Path|None=None,expected_changed_members:list[str]|None=None)->dict:
    base={'application_compatibility':'not_executed','visual_fidelity':'not_executed'}
    try:
        if not path.is_file() or path.suffix.lower()!='.docx':raise ArtifactError('validation_failure','not a DOCX file')
        with zipfile.ZipFile(path) as archive:
            infos=archive.infolist();names=[x.filename for x in infos];name_set=set(names)
            if len(names)!=len(name_set):raise ArtifactError('validation_failure','duplicate ZIP member')
            total=sum(x.file_size for x in infos)
            if len(infos)>5000 or total>64*1024*1024:raise ArtifactError('validation_failure','package budget exceeded')
            for info in infos:
                pure=PurePosixPath(info.filename)
                mode=(info.external_attr>>16)&0xffff
                member_type=stat.S_IFMT(mode)
                if info.filename.startswith('/') or '..' in pure.parts or '\\' in info.filename or not (info.is_dir() or member_type in (0,stat.S_IFREG)):raise ArtifactError('validation_failure','unsafe ZIP member')
                if info.file_size>16*1024*1024 or info.file_size>1024*1024 and info.file_size/max(info.compress_size,1)>200:raise ArtifactError('validation_failure','ZIP member budget exceeded')
            required={'[Content_Types].xml','_rels/.rels','word/document.xml'}
            if not required<=name_set:raise ArtifactError('validation_failure','required package member missing')
            for name in names:
                if name.endswith('.xml') or name.endswith('.rels'):etree.fromstring(archive.read(name))
            relationships_valid=True
            for name in [x for x in names if x.endswith('.rels')]:
                root=etree.fromstring(archive.read(name));source=_source_for_rels(name)
                for rel in root.findall(f'{{{PR}}}Relationship'):
                    if rel.get('TargetMode')=='External':continue
                    target=_resolve(source,rel.get('Target',''))
                    if target not in name_set:relationships_valid=False
            ct=etree.fromstring(archive.read('[Content_Types].xml'));defaults={x.get('Extension') for x in ct.findall(f'{{{CT}}}Default')};overrides={x.get('PartName','').lstrip('/') for x in ct.findall(f'{{{CT}}}Override')}
            def extension(name:str)->str:
                filename=posixpath.basename(name)
                return filename.rsplit('.',1)[-1] if '.' in filename else ''
            content_types_valid=all(name in overrides or extension(name) in defaults for name in names if name!='[Content_Types].xml' and not name.endswith('/'))
            hashes={x.filename:hashlib.sha256(archive.read(x.filename)).hexdigest() for x in infos}
        changed=[];unchanged=sorted(hashes)
        if before:
            with zipfile.ZipFile(before) as archive:old={x.filename:hashlib.sha256(archive.read(x.filename)).hexdigest() for x in archive.infolist()}
            changed=sorted(k for k in old|hashes if old.get(k)!=hashes.get(k));unchanged=sorted(k for k in old.keys()&hashes.keys() if old[k]==hashes[k])
        unexpected=sorted(set(changed)-set(expected_changed_members if expected_changed_members is not None else changed))
        valid=relationships_valid and content_types_valid and not unexpected
        return {'status':'valid' if valid else 'invalid','package_valid':True,'xml_valid':True,'relationships_valid':relationships_valid,'content_types_valid':content_types_valid,'changed_members':changed,'unchanged_members':unchanged,'unexpected_changed_members':unexpected,**base}
    except ArtifactError as e:return {'status':'invalid','error':e.details,**base}
    except Exception:return {'status':'invalid','error':'invalid_package',**base}
