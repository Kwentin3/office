"""Provider-neutral host adapter example for an Open WebUI backend.

The host owns authentication, user/request IDs, upload storage, quotas, cleanup,
and attachment registration. This module only demonstrates safe workspace and
package boundaries.
"""

from __future__ import annotations

import os
import re
import shutil
import stat
import weakref
from contextlib import ExitStack
from pathlib import Path

from office_artifact_tool import DocxArtifactTool
from pptx_artifact_tool import PptxArtifactTool
from xlsx_artifact_tool import XlsxArtifactTool

_SAFE_ID = re.compile(r"^[A-Za-z0-9_-]{1,80}$")
_ALLOWED_SUFFIXES = {".docx", ".xlsx", ".pptx"}
_DIRECTORY_FLAGS = (
    os.O_RDONLY
    | getattr(os, "O_DIRECTORY", 0)
    | getattr(os, "O_NOFOLLOW", 0)
    | getattr(os, "O_CLOEXEC", 0)
)


def _same_directory(open_fd: int, parent_fd: int, name: str) -> bool:
    opened = os.fstat(open_fd)
    try:
        current = os.stat(name, dir_fd=parent_fd, follow_symlinks=False)
    except OSError:
        return False
    return stat.S_ISDIR(current.st_mode) and (opened.st_dev, opened.st_ino) == (
        current.st_dev,
        current.st_ino,
    )


class OfficeService:
    def __init__(self, root: str | Path) -> None:
        configured = Path(root)
        if configured.is_symlink():
            raise ValueError("workspace root must not be a symlink")
        configured.mkdir(mode=0o700, parents=True, exist_ok=True)
        try:
            self._root_fd = os.open(configured, _DIRECTORY_FLAGS)
        except OSError as exc:
            raise ValueError("workspace root must be a non-symlink directory") from exc
        self._finalizer = weakref.finalize(self, os.close, self._root_fd)
        self.root = configured.resolve()

    def close(self) -> None:
        if self._finalizer.alive:
            self._finalizer()
        self._root_fd = -1

    def __enter__(self) -> "OfficeService":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    @staticmethod
    def _open_directory(parent_fd: int, name: str) -> int:
        if Path(name).name != name:
            raise ValueError("workspace directory name must be a basename")
        try:
            os.mkdir(name, mode=0o700, dir_fd=parent_fd)
        except FileExistsError:
            pass
        try:
            directory_fd = os.open(name, _DIRECTORY_FLAGS, dir_fd=parent_fd)
        except OSError as exc:
            raise ValueError("workspace directory must not be a symlink") from exc
        if not _same_directory(directory_fd, parent_fd, name):
            os.close(directory_fd)
            raise ValueError("workspace changed during directory open")
        return directory_fd

    def _request_fd(self, request_id: str) -> int:
        if not isinstance(request_id, str) or not _SAFE_ID.fullmatch(request_id):
            raise ValueError("invalid server-issued request id")
        if self._root_fd < 0:
            raise RuntimeError("office service is closed")
        return self._open_directory(self._root_fd, request_id)

    def request_root(self, request_id: str) -> Path:
        request_fd = self._request_fd(request_id)
        os.close(request_fd)
        return self.root / request_id

    def stage_upload(self, request_id: str, source: str | Path, suffix: str) -> Path:
        source_path = Path(source)
        if suffix not in _ALLOWED_SUFFIXES:
            raise ValueError("unsupported Office suffix")
        source_fd = request_fd = input_fd = output_fd = -1
        destination_name = f"source{suffix}"
        destination_created = False
        try:
            try:
                source_fd = os.open(
                    source_path,
                    os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_CLOEXEC", 0),
                )
            except OSError as exc:
                raise ValueError("upload source must be a readable regular file") from exc
            if not stat.S_ISREG(os.fstat(source_fd).st_mode):
                raise ValueError("upload source must be a regular file")
            request_fd = self._request_fd(request_id)
            input_fd = self._open_directory(request_fd, "input")
            output_fd = os.open(
                destination_name,
                os.O_WRONLY
                | os.O_CREAT
                | os.O_EXCL
                | getattr(os, "O_NOFOLLOW", 0)
                | getattr(os, "O_CLOEXEC", 0),
                0o600,
                dir_fd=input_fd,
            )
            destination_created = True
            with ExitStack() as stack:
                reader = stack.enter_context(os.fdopen(source_fd, "rb"))
                source_fd = -1
                writer = stack.enter_context(os.fdopen(output_fd, "wb"))
                output_fd = -1
                shutil.copyfileobj(reader, writer, length=1024 * 1024)
            if not _same_directory(input_fd, request_fd, "input"):
                raise ValueError("workspace changed during staging")
        except Exception:
            if destination_created and input_fd >= 0:
                try:
                    os.unlink(destination_name, dir_fd=input_fd)
                except FileNotFoundError:
                    pass
            raise
        finally:
            for descriptor in (source_fd, output_fd, input_fd, request_fd):
                if descriptor >= 0:
                    os.close(descriptor)
        return self.root / request_id / "input" / destination_name

    def _directory_path(self, request_id: str, *names: str) -> Path:
        descriptors: list[int] = []
        try:
            parent_fd = self._request_fd(request_id)
            descriptors.append(parent_fd)
            for name in names:
                parent_fd = self._open_directory(parent_fd, name)
                descriptors.append(parent_fd)
            return self.root.joinpath(request_id, *names)
        finally:
            for descriptor in reversed(descriptors):
                os.close(descriptor)

    def output_path(self, request_id: str, filename: str) -> Path:
        name = Path(filename)
        if name.name != filename or name.suffix.lower() not in _ALLOWED_SUFFIXES:
            raise ValueError("output must be a basename with a supported suffix")
        return self._directory_path(request_id, "output") / filename

    def _domain_workdir(self, request_id: str, domain: str) -> Path:
        return self._directory_path(request_id, "internal", domain)

    def docx_tool(self, request_id: str) -> DocxArtifactTool:
        return DocxArtifactTool(self._domain_workdir(request_id, "docx"))

    def xlsx_tool(self, request_id: str) -> XlsxArtifactTool:
        return XlsxArtifactTool(self._domain_workdir(request_id, "xlsx"))

    def pptx_editor(self, request_id: str) -> PptxArtifactTool:
        return PptxArtifactTool(self._domain_workdir(request_id, "pptx-editor"))
