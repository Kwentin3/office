from __future__ import annotations
import os,tempfile
from pathlib import Path
from typing import Callable


def atomic_candidate(output: Path, build: Callable[[Path],None], accept: Callable[[Path],dict]) -> dict:
    output.parent.mkdir(parents=True,exist_ok=True)
    fd,name=tempfile.mkstemp(prefix=f'.{output.name}.candidate.',suffix='.docx',dir=output.parent);os.close(fd)
    candidate=Path(name)
    try:
        build(candidate)
        report=accept(candidate)
        if report.get('status')!='valid':raise ValueError(report.get('error','validation_failure'))
        os.replace(candidate,output)
        return report
    finally:candidate.unlink(missing_ok=True)
