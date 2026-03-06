#!/usr/bin/env python3
"""Batch unzip ZIP files in a folder.

Default behavior:
- Input dir: current directory
- Scan: recursive
- Archive pattern: *.zip
- Output: each archive extracted to a sibling folder with the same name

Usage:
    python unzip_in_folder.py
    python unzip_in_folder.py /path/to/folder
    python unzip_in_folder.py . --no-recursive
    python unzip_in_folder.py . --overwrite
    python unzip_in_folder.py . --no-flatten
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Batch unzip ZIP files in a folder")
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=".",
        help="Directory to scan for ZIP files (default: current directory).",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        default=True,
        help="Scan subdirectories recursively (default: enabled).",
    )
    parser.add_argument(
        "--no-recursive",
        dest="recursive",
        action="store_false",
        help="Scan only the top-level directory.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite extracted folder if it already exists.",
    )
    parser.add_argument(
        "--flatten",
        action="store_true",
        default=True,
        help="Move extracted files from nested folders to the output folder (default: enabled).",
    )
    parser.add_argument(
        "--no-flatten",
        dest="flatten",
        action="store_false",
        help="Keep extracted directory structure.",
    )
    return parser.parse_args()


def iter_zip_files(root: Path, recursive: bool) -> list[Path]:
    pattern = "**/*.zip" if recursive else "*.zip"
    return sorted([p for p in root.glob(pattern) if p.is_file()])


def _is_within(base_dir: Path, target_path: Path) -> bool:
    base = base_dir.resolve()
    target = target_path.resolve()
    return str(target).startswith(str(base) + "/") or target == base


def safe_extract(zip_path: Path, target_dir: Path) -> None:
    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.infolist():
            member_path = target_dir / member.filename
            if not _is_within(target_dir, member_path):
                raise RuntimeError(
                    f"Unsafe ZIP entry detected in {zip_path.name}: {member.filename}"
                )
        zf.extractall(target_dir)


def _next_available_path(dst: Path) -> Path:
    if not dst.exists():
        return dst

    stem = dst.stem
    suffix = dst.suffix
    parent = dst.parent
    index = 1

    while True:
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def flatten_extracted_files(output_dir: Path, overwrite: bool) -> int:
    moved = 0
    nested_files = [
        p for p in output_dir.rglob("*") if p.is_file() and p.parent != output_dir
    ]

    for src in sorted(nested_files):
        dst = output_dir / src.name
        if dst.exists():
            if overwrite:
                dst.unlink()
            else:
                dst = _next_available_path(dst)

        shutil.move(str(src), str(dst))
        moved += 1

    # Clean up directories left after moving files.
    nested_dirs = [p for p in output_dir.rglob("*") if p.is_dir()]
    for folder in sorted(nested_dirs, key=lambda p: len(p.parts), reverse=True):
        try:
            folder.rmdir()
        except OSError:
            # Ignore non-empty directories.
            pass

    return moved


def unzip_file(zip_path: Path, overwrite: bool, flatten: bool) -> tuple[str, Path, int]:
    output_dir = zip_path.with_suffix("")

    if output_dir.exists():
        if not overwrite:
            return "skip", output_dir, 0
        shutil.rmtree(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)
    safe_extract(zip_path, output_dir)

    moved_count = flatten_extracted_files(output_dir, overwrite) if flatten else 0
    return "ok", output_dir, moved_count


def main() -> int:
    args = parse_args()
    root = Path(args.input_dir).expanduser().resolve()

    if not root.exists() or not root.is_dir():
        print(f"Error: directory not found: {root}", file=sys.stderr)
        return 1

    zip_files = iter_zip_files(root, args.recursive)
    if not zip_files:
        print("No ZIP files found.")
        return 0

    extracted = 0
    skipped = 0
    failed = 0

    for zip_file in zip_files:
        try:
            status, output_dir, moved_count = unzip_file(
                zip_file, args.overwrite, args.flatten
            )
            if status == "skip":
                print(f"Skip (exists): {output_dir}")
                skipped += 1
            else:
                if args.flatten:
                    print(f"OK: {zip_file} -> {output_dir} (flatten moved={moved_count})")
                else:
                    print(f"OK: {zip_file} -> {output_dir}")
                extracted += 1
        except zipfile.BadZipFile:
            print(f"Fail: {zip_file} (invalid ZIP file)", file=sys.stderr)
            failed += 1
        except Exception as exc:  # noqa: BLE001
            print(f"Fail: {zip_file} ({exc})", file=sys.stderr)
            failed += 1

    print(
        f"Done. extracted={extracted}, skipped={skipped}, failed={failed}, total={len(zip_files)}"
    )
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
