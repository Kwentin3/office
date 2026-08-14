"""Run a private LibreOffice XLSX round-trip and verify declared caches."""

from __future__ import annotations

import argparse
import json
import os
import secrets
import shutil
import subprocess
from pathlib import Path

from office_application_witness.api import (
    _inspect_output,
    _open_bounded_regular,
    _resolve_executable,
    _run_process,
    _runtime_identity,
    _validate_directory_chain,
    _validate_existing_directory_chain,
    _validate_source_clone,
)
from verify_xlsx_recalculation import GateRefusal, verify

_MAX_SOURCE_BYTES = 256 * 1024 * 1024


def _prepare_root(path: Path) -> os.stat_result:
    if not path.is_absolute():
        raise ValueError
    _validate_existing_directory_chain(path)
    path.mkdir(parents=True, mode=0o700, exist_ok=True)
    info = _validate_directory_chain(path)
    if info.st_uid != os.geteuid() or info.st_mode & 0o077:
        raise ValueError
    return info


def run(args: argparse.Namespace) -> dict:
    root = Path(args.workdir)
    root_info = _prepare_root(root)
    executable = _resolve_executable(args.executable)
    identity = _runtime_identity(
        {
            "application_version": args.runtime_version,
            "image_digest": args.runtime_image_digest,
        }
    )
    source = Path(args.source)
    descriptor = -1
    root_fd = -1
    workspace: Path | None = None
    workspace_name: str | None = None
    try:
        root_flags = os.O_RDONLY | getattr(os, "O_DIRECTORY", 0) | getattr(os, "O_NOFOLLOW", 0) | getattr(os, "O_CLOEXEC", 0)
        root_fd = os.open(root, root_flags)
        opened_root = os.fstat(root_fd)
        if (opened_root.st_dev, opened_root.st_ino) != (root_info.st_dev, root_info.st_ino):
            raise ValueError
        baseline_entries = set(os.listdir(root_fd))
        descriptor, _source_info = _open_bounded_regular(source, _MAX_SOURCE_BYTES)
        workspace_name = ".xlsx-recalculation." + secrets.token_hex(16)
        os.mkdir(workspace_name, mode=0o700, dir_fd=root_fd)
        created = os.stat(workspace_name, dir_fd=root_fd, follow_symlinks=False)
        workspace = root / workspace_name
        current = workspace.lstat()
        if (created.st_dev, created.st_ino) != (current.st_dev, current.st_ino):
            raise GateRefusal("cleanup_failure")
        input_dir = workspace / "input"
        output_dir = workspace / "output"
        profile_dir = workspace / "profile"
        home_dir = workspace / "home"
        for directory in (input_dir, output_dir, profile_dir, home_dir):
            directory.mkdir(mode=0o700)
        clone = input_dir / "artifact.xlsx"
        clone_fd = os.open(clone, os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_NOFOLLOW", 0), 0o600)
        with os.fdopen(descriptor, "rb", closefd=True) as original, os.fdopen(clone_fd, "wb") as target:
            descriptor = -1
            for block in iter(lambda: original.read(1024 * 1024), b""):
                target.write(block)
        _validate_source_clone(clone, "xlsx")
        environment = {
            "HOME": str(home_dir),
            "LANG": "C.UTF-8",
            "LC_ALL": "C.UTF-8",
            "PATH": os.defpath,
            "TMPDIR": str(workspace),
        }
        argv = [
            str(executable),
            f"-env:UserInstallation={profile_dir.as_uri()}",
            "--headless",
            "--nologo",
            "--nodefault",
            "--norestore",
            "--convert-to",
            "xlsx:Calc MS Excel 2007 XML",
            "--outdir",
            str(output_dir),
            str(clone),
        ]
        returncode = _run_process(
            argv,
            cwd=workspace,
            env=environment,
            timeout=float(args.timeout_seconds),
        )
        if returncode != 0:
            raise GateRefusal("application_failure")
        normalized = output_dir / "artifact.xlsx"
        _inspect_output(normalized, "xlsx")
        report = verify(clone, normalized, Path(args.cases))
        output_sha256 = report.pop("normalized_snapshot_sha256")
        output_bytes = report.pop("normalized_snapshot_bytes")
        report.update(
            {
                "runtime_identity": identity,
                "process_exit": "pass",
                "normalized_output_sha256": output_sha256,
                "normalized_output_bytes": output_bytes,
                "normalized_output_retained": False,
                "microsoft_office_equivalence_claimed": False,
            }
        )
        shutil.rmtree(workspace_name, dir_fd=root_fd)
        root_current = root.lstat()
        if (
            workspace_name in os.listdir(root_fd)
            or set(os.listdir(root_fd)) != baseline_entries
            or (root_current.st_dev, root_current.st_ino) != (root_info.st_dev, root_info.st_ino)
        ):
            raise GateRefusal("cleanup_failure")
        workspace = None
        workspace_name = None
        return report
    finally:
        if descriptor >= 0:
            os.close(descriptor)
        if workspace_name is not None and root_fd >= 0:
            shutil.rmtree(workspace_name, dir_fd=root_fd, ignore_errors=True)
        if root_fd >= 0:
            os.close(root_fd)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--cases", required=True)
    parser.add_argument("--workdir", required=True)
    parser.add_argument("--executable", required=True)
    parser.add_argument("--runtime-version", required=True)
    parser.add_argument("--runtime-image-digest", required=True)
    parser.add_argument("--timeout-seconds", type=float, default=60.0)
    args = parser.parse_args()
    try:
        if not 1 <= args.timeout_seconds <= 300:
            raise ValueError
        report = run(args)
    except subprocess.TimeoutExpired:
        report = {"schema_version": 1, "status": "refused", "reason": "application_timeout"}
    except GateRefusal as exc:
        report = {"schema_version": 1, "status": "refused", "reason": exc.reason}
    except (OSError, ValueError, TypeError, KeyError):
        report = {"schema_version": 1, "status": "refused", "reason": "invalid_input"}
    print(json.dumps(report, ensure_ascii=False, separators=(",", ":")))
    return 0 if report["status"] == "ok" else 2


if __name__ == "__main__":
    raise SystemExit(main())
