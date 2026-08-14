from __future__ import annotations

import hashlib
import os
import re
import secrets
import shutil
import signal
import stat
import subprocess
import time
import zipfile
from pathlib import Path, PurePosixPath
from typing import BinaryIO

from .contracts import ArtifactType, RefusalReason, RuntimeIdentity, WitnessRefusal, WitnessResult

_FORMATS: dict[ArtifactType, str] = {"docx": ".docx", "xlsx": ".xlsx", "pptx": ".pptx"}
_CONVERSIONS: dict[ArtifactType, str] = {
    "docx": "pdf:writer_pdf_Export",
    "xlsx": "xlsx:Calc MS Excel 2007 XML",
    "pptx": "pdf:impress_pdf_Export",
}
_MAX_SOURCE_BYTES = 256 * 1024 * 1024
_MAX_OUTPUT_BYTES = 512 * 1024 * 1024
_IMAGE_DIGEST = re.compile(r"sha256:[0-9a-f]{64}\Z")


def _runtime_identity(value: RuntimeIdentity | None) -> RuntimeIdentity:
    if value is None:
        return {"application_version": "not_observed", "image_digest": "not_observed"}
    if not isinstance(value, dict) or set(value) != {"application_version", "image_digest"}:
        raise ValueError("invalid runtime identity")
    version = value.get("application_version")
    digest = value.get("image_digest")
    if version == "not_observed" and digest == "not_observed":
        return {"application_version": version, "image_digest": digest}
    if (
        not isinstance(version, str)
        or not 1 <= len(version) <= 128
        or any(ord(character) < 0x20 or ord(character) == 0x7F for character in version)
        or not isinstance(digest, str)
        or not _IMAGE_DIGEST.fullmatch(digest)
    ):
        raise ValueError("invalid runtime identity")
    return {"application_version": version, "image_digest": digest}


def _refusal(reason: RefusalReason, details: str = "") -> WitnessRefusal:
    return {"schema_version": 1, "status": "refused", "reason": reason, "details": details}


def _stream_sha256(stream: BinaryIO) -> tuple[str, int]:
    digest = hashlib.sha256()
    total = 0
    for block in iter(lambda: stream.read(1024 * 1024), b""):
        digest.update(block)
        total += len(block)
    return digest.hexdigest(), total


def _same_file_state(path: Path, descriptor: int, initial: os.stat_result) -> bool:
    try:
        opened = os.fstat(descriptor)
        current = path.lstat()
    except OSError:
        return False
    fields = ("st_dev", "st_ino", "st_mode", "st_size", "st_mtime_ns", "st_ctime_ns")
    expected = tuple(getattr(initial, field) for field in fields)
    return tuple(getattr(opened, field) for field in fields) == expected and tuple(
        getattr(current, field) for field in fields
    ) == expected


def _same_file_state_at(name: str, directory_fd: int, descriptor: int, initial: os.stat_result) -> bool:
    try:
        opened = os.fstat(descriptor)
        current = os.stat(name, dir_fd=directory_fd, follow_symlinks=False)
    except OSError:
        return False
    fields = ("st_dev", "st_ino", "st_mode", "st_size", "st_mtime_ns", "st_ctime_ns")
    expected = tuple(getattr(initial, field) for field in fields)
    return tuple(getattr(opened, field) for field in fields) == expected and tuple(
        getattr(current, field) for field in fields
    ) == expected


def _open_private_directory(path: Path) -> tuple[int, os.stat_result]:
    flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0) | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_CLOEXEC", 0)
    descriptor = os.open(path, flags)
    try:
        opened = os.fstat(descriptor)
        current = path.lstat()
        if (
            not stat.S_ISDIR(opened.st_mode)
            or stat.S_ISLNK(current.st_mode)
            or (opened.st_dev, opened.st_ino) != (current.st_dev, current.st_ino)
        ):
            raise ValueError("expected private output directory")
        return descriptor, opened
    except Exception:
        os.close(descriptor)
        raise


def _directory_matches(path: Path, descriptor: int, initial: os.stat_result) -> bool:
    try:
        opened = os.fstat(descriptor)
        current = path.lstat()
    except OSError:
        return False
    return (
        stat.S_ISDIR(opened.st_mode)
        and stat.S_ISDIR(current.st_mode)
        and not stat.S_ISLNK(current.st_mode)
        and (opened.st_dev, opened.st_ino, opened.st_uid)
        == (initial.st_dev, initial.st_ino, initial.st_uid)
        == (current.st_dev, current.st_ino, current.st_uid)
    )


def _open_bounded_regular(path: Path, maximum_bytes: int) -> tuple[int, os.stat_result]:
    before = path.lstat()
    if stat.S_ISLNK(before.st_mode) or not stat.S_ISREG(before.st_mode) or not 0 < before.st_size <= maximum_bytes:
        raise ValueError("expected a bounded regular file")
    flags = os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_NONBLOCK", 0) | getattr(os, "O_CLOEXEC", 0)
    descriptor = os.open(path, flags)
    try:
        opened = os.fstat(descriptor)
        if (
            not stat.S_ISREG(opened.st_mode)
            or not 0 < opened.st_size <= maximum_bytes
            or (opened.st_dev, opened.st_ino) != (before.st_dev, before.st_ino)
        ):
            raise ValueError("expected a stable bounded regular file")
        return descriptor, opened
    except Exception:
        os.close(descriptor)
        raise


def _resolve_executable(value: str | Path) -> Path:
    executable = Path(value)
    if not executable.is_absolute():
        raise FileNotFoundError
    executable = executable.resolve(strict=True)
    info = executable.stat()
    if not stat.S_ISREG(info.st_mode) or not os.access(executable, os.X_OK):
        raise FileNotFoundError
    return executable


def _validate_directory_chain(path: Path) -> os.stat_result:
    if not path.is_absolute():
        raise ValueError("witness workdir must be absolute")
    current = Path(path.anchor)
    for part in path.parts[1:]:
        current /= part
        info = current.lstat()
        if stat.S_ISLNK(info.st_mode) or not stat.S_ISDIR(info.st_mode):
            raise ValueError("invalid witness workdir")
    return path.lstat()


def _validate_existing_directory_chain(path: Path) -> None:
    current = Path(path.anchor)
    for part in path.parts[1:]:
        current /= part
        try:
            info = current.lstat()
        except FileNotFoundError:
            return
        if stat.S_ISLNK(info.st_mode) or not stat.S_ISDIR(info.st_mode):
            raise ValueError("invalid witness workdir")


def _create_private_workspace(root: Path, identity: tuple[int, int]) -> tuple[int, Path]:
    flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0) | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_CLOEXEC", 0)
    root_fd = os.open(root, flags)
    opened = os.fstat(root_fd)
    if (opened.st_dev, opened.st_ino) != identity:
        os.close(root_fd)
        raise ValueError("witness workdir identity changed")
    name = ".office-witness." + secrets.token_hex(16)
    try:
        os.mkdir(name, mode=0o700, dir_fd=root_fd)
        created = os.stat(name, dir_fd=root_fd, follow_symlinks=False)
        workspace = root / name
        current = workspace.lstat()
        if (created.st_dev, created.st_ino) != (current.st_dev, current.st_ino):
            raise ValueError("witness workdir changed during workspace creation")
        return root_fd, workspace
    except Exception:
        os.close(root_fd)
        raise


def _validate_source_clone(path: Path, artifact_type: ArtifactType) -> None:
    required = {
        "docx": {"[Content_Types].xml", "word/document.xml"},
        "xlsx": {"[Content_Types].xml", "xl/workbook.xml"},
        "pptx": {"[Content_Types].xml", "ppt/presentation.xml"},
    }[artifact_type]
    if not zipfile.is_zipfile(path):
        raise ValueError("invalid Office package")
    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        names = [info.filename for info in infos]
        total = sum(info.file_size for info in infos)
        unsafe = (
            len(infos) > 10_000
            or total > _MAX_SOURCE_BYTES
            or len(names) != len(set(names))
            or not required.issubset(names)
            or any(name.startswith("/") or "\\" in name or ".." in PurePosixPath(name).parts for name in names)
        )
        if unsafe:
            raise ValueError("invalid Office package")


def _run_process(
    argv: list[str],
    *,
    cwd: Path,
    env: dict[str, str],
    timeout: float,
) -> int:
    process = subprocess.Popen(
        argv,
        cwd=cwd,
        env=env,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        shell=False,
        close_fds=True,
        start_new_session=True,
    )
    try:
        return process.wait(timeout=timeout)
    finally:
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
        deadline = time.monotonic() + 0.5
        while time.monotonic() < deadline:
            try:
                os.killpg(process.pid, 0)
            except ProcessLookupError:
                break
            time.sleep(0.01)
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        if process.poll() is None:
            process.wait()


def _inspect_output_at(
    directory: Path,
    directory_fd: int,
    directory_info: os.stat_result,
    name: str,
    artifact_type: ArtifactType,
) -> tuple[str, int]:
    if "/" in name or name in {"", ".", ".."} or not _directory_matches(directory, directory_fd, directory_info):
        raise ValueError("expected stable private output directory")
    flags = os.O_RDONLY | os.O_NONBLOCK | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_CLOEXEC", 0)
    descriptor = os.open(name, flags, dir_fd=directory_fd)
    try:
        info = os.fstat(descriptor)
        if not stat.S_ISREG(info.st_mode) or info.st_size > _MAX_OUTPUT_BYTES:
            raise ValueError("expected bounded regular file")
        with os.fdopen(descriptor, "rb", closefd=False) as stream:
            if artifact_type in {"docx", "pptx"}:
                if stream.read(5) != b"%PDF-":
                    raise ValueError("invalid witness PDF")
                stream.seek(0)
            else:
                with zipfile.ZipFile(stream) as archive:
                    if "xl/workbook.xml" not in archive.namelist():
                        raise ValueError("invalid witness XLSX")
                stream.seek(0)
            digest, bytes_read = _stream_sha256(stream)
            if (
                bytes_read != info.st_size
                or not _same_file_state_at(name, directory_fd, descriptor, info)
                or not _directory_matches(directory, directory_fd, directory_info)
            ):
                raise ValueError("expected stable witness output")
            return digest, bytes_read
    finally:
        os.close(descriptor)


def _inspect_output(path: Path, artifact_type: ArtifactType) -> tuple[str, int]:
    directory_fd, directory_info = _open_private_directory(path.parent)
    try:
        return _inspect_output_at(path.parent, directory_fd, directory_info, path.name, artifact_type)
    finally:
        os.close(directory_fd)


def _source_matches(path: Path, original: os.stat_result, expected_sha256: str) -> bool:
    try:
        descriptor, current = _open_bounded_regular(path, _MAX_SOURCE_BYTES)
    except (OSError, ValueError):
        return False
    try:
        if (
            current.st_dev != original.st_dev
            or current.st_ino != original.st_ino
            or current.st_size != original.st_size
            or current.st_mtime_ns != original.st_mtime_ns
            or current.st_ctime_ns != original.st_ctime_ns
        ):
            return False
        with os.fdopen(descriptor, "rb", closefd=False) as stream:
            digest, bytes_read = _stream_sha256(stream)
        return bytes_read == current.st_size and digest == expected_sha256 and _same_file_state(path, descriptor, current)
    finally:
        os.close(descriptor)


class ApplicationWitness:
    """Observe a private clone with a trusted LibreOffice executable."""

    def __init__(
        self,
        workdir: str | Path,
        *,
        executable: str | Path,
        runtime_identity: RuntimeIdentity | None = None,
    ):
        self.workdir = Path(workdir)
        if not self.workdir.is_absolute():
            raise ValueError("witness workdir must be absolute")
        _validate_existing_directory_chain(self.workdir)
        self.workdir.mkdir(parents=True, mode=0o700, exist_ok=True)
        info = _validate_directory_chain(self.workdir)
        if info.st_uid != os.geteuid() or info.st_mode & 0o077:
            raise ValueError("witness workdir must be private and owned by the current OS identity")
        self._workdir_identity = (info.st_dev, info.st_ino)
        self.executable = executable
        self.runtime_identity = _runtime_identity(runtime_identity)

    def observe(
        self,
        source: str | Path,
        artifact_type: str,
        *,
        timeout_seconds: float = 60.0,
    ) -> WitnessResult:
        started = time.perf_counter()
        workspace: Path | None = None
        root_fd = -1
        source_fd = -1
        output_fd = -1
        try:
            try:
                runtime_identity = _runtime_identity(self.runtime_identity)
            except ValueError:
                return _refusal("validation_failure", "configured runtime identity is invalid")
            source = Path(source)
            if artifact_type not in _FORMATS or source.suffix.lower() != _FORMATS[artifact_type]:
                return _refusal("validation_failure", "artifact type and suffix must match")
            typed_artifact: ArtifactType = artifact_type  # type: ignore[assignment]
            if (
                not isinstance(timeout_seconds, (int, float))
                or isinstance(timeout_seconds, bool)
                or not 1 <= timeout_seconds <= 300
            ):
                return _refusal("validation_failure", "timeout must be between 1 and 300 seconds")
            try:
                executable_path = _resolve_executable(self.executable)
            except FileNotFoundError:
                return _refusal("application_unavailable", "configured LibreOffice executable is unavailable")
            root_info = _validate_directory_chain(self.workdir)
            if (root_info.st_dev, root_info.st_ino) != self._workdir_identity:
                return _refusal("validation_failure", "witness workdir identity changed")
            source_fd, source_stat = _open_bounded_regular(source, _MAX_SOURCE_BYTES)
            root_fd, workspace = _create_private_workspace(self.workdir, self._workdir_identity)
            input_dir = workspace / "input"
            output_dir = workspace / "output"
            profile_dir = workspace / "profile"
            home_dir = workspace / "home"
            for directory in (input_dir, output_dir, profile_dir, home_dir):
                directory.mkdir(mode=0o700)
            output_fd, output_info = _open_private_directory(output_dir)
            clone = input_dir / ("artifact" + _FORMATS[typed_artifact])
            clone_fd = os.open(
                clone,
                os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_NOFOLLOW", 0),
                0o600,
            )
            digest = hashlib.sha256()
            with os.fdopen(source_fd, "rb", closefd=True) as original, os.fdopen(clone_fd, "wb") as target:
                source_fd = -1
                for block in iter(lambda: original.read(1024 * 1024), b""):
                    digest.update(block)
                    target.write(block)
            source_sha256 = digest.hexdigest()
            _validate_source_clone(clone, typed_artifact)
            environment = {
                "HOME": str(home_dir),
                "LANG": "C.UTF-8",
                "LC_ALL": "C.UTF-8",
                "PATH": os.defpath,
                "TMPDIR": str(workspace),
            }
            argv = [
                str(executable_path),
                f"-env:UserInstallation={profile_dir.as_uri()}",
                "--headless",
                "--nologo",
                "--nodefault",
                "--norestore",
                "--convert-to",
                _CONVERSIONS[typed_artifact],
                "--outdir",
                str(output_dir),
                str(clone),
            ]
            returncode = _run_process(
                argv,
                cwd=workspace,
                env=environment,
                timeout=float(timeout_seconds),
            )
            if returncode != 0:
                return _refusal("application_failure", "LibreOffice exited unsuccessfully")
            expected_name = clone.stem + (".xlsx" if typed_artifact == "xlsx" else ".pdf")
            try:
                output_sha256, output_bytes = _inspect_output_at(
                    output_dir,
                    output_fd,
                    output_info,
                    expected_name,
                    typed_artifact,
                )
            except FileNotFoundError:
                return _refusal("application_failure", "LibreOffice did not produce the expected output")
            if not _source_matches(source, source_stat, source_sha256):
                return _refusal("stale_snapshot", "source changed during application observation")
            try:
                shutil.rmtree(workspace)
            except OSError:
                return _refusal("validation_failure", "private workspace cleanup failed")
            workspace = None
            return {
                "schema_version": 1,
                "status": "ok",
                "artifact_type": typed_artifact,
                "source_sha256": source_sha256,
                "source_bytes": source_stat.st_size,
                "source_unchanged": True,
                "witness": {
                    "application": "LibreOffice",
                    "version": runtime_identity["application_version"],
                    "runtime_identity": dict(runtime_identity),
                    "claim": "libreoffice_private_clone_observed",
                    "operation": "recalculation_roundtrip" if typed_artifact == "xlsx" else "pdf_render",
                    "process_exit": "pass",
                    "output_validation": "pass",
                    "output_sha256": output_sha256,
                    "output_bytes": output_bytes,
                    "repair_dialog": "not_observable_headless",
                    "formula_recalculation": "requested_not_semantically_verified"
                    if typed_artifact == "xlsx"
                    else "not_applicable",
                    "application_normalized_clone": output_sha256 != source_sha256
                    if typed_artifact == "xlsx"
                    else "not_applicable",
                    "process_isolation": "trusted_executable_not_sandboxed",
                },
                "private_workspace_artifacts_retained": False,
                "microsoft_office_equivalence": "not_claimed",
                "latency_ms": round((time.perf_counter() - started) * 1000, 3),
            }
        except subprocess.TimeoutExpired:
            return _refusal("application_timeout", "LibreOffice witness timed out")
        except (OSError, TypeError, ValueError, zipfile.BadZipFile):
            return _refusal("validation_failure", "witness could not validate the private clone")
        finally:
            if source_fd >= 0:
                os.close(source_fd)
            if output_fd >= 0:
                os.close(output_fd)
            if workspace is not None:
                try:
                    shutil.rmtree(workspace)
                except OSError:
                    pass
            if root_fd >= 0:
                os.close(root_fd)
