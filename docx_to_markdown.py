#!/usr/bin/env python3
"""Batch convert DOCX files to Markdown.

Usage:
    python docx_to_markdown.py /path/to/folder
    python docx_to_markdown.py . --recursive

Dependencies:
    pip install mammoth
"""

from __future__ import annotations

import argparse
import base64
import io
import re
import sys
import zipfile
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert DOCX files to Markdown files in batch."
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=".",
        help="Directory to scan for DOCX files (default: current directory).",
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
        help="Overwrite existing Markdown files.",
    )
    return parser.parse_args()


def iter_docx_files(root: Path, recursive: bool) -> list[Path]:
    pattern = "**/*.docx" if recursive else "*.docx"
    files = [p for p in root.glob(pattern) if p.is_file()]
    # Skip temporary lock files created by Microsoft Office.
    return [p for p in files if not p.name.startswith("~$")]


_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+yf7sAAAAASUVORK5CYII="
)
_JPEG_1X1 = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wCEAAkGBxAQEBUQEA8VFRUVFRUVFRUVFRUVFRUXFhUWFhUVFRUYHSggGBolGxUVITEhJSkrLi4uFx8zODMsNygtLisBCgoKDg0OFQ8PGisdFR0rKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrKysrK//AABEIAAEAAQMBEQACEQEDEQH/xAAbAAADAQEBAQEAAAAAAAAAAAAABQYDBEcBAv/EABQBAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhADEAAAAdQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAT8AIP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQIBAT8AIP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQMBAT8AIP/Z"
)


def _placeholder_media_bytes(media_name: str) -> bytes:
    lower = media_name.lower()
    if lower.endswith(".jpg") or lower.endswith(".jpeg"):
        return _JPEG_1X1
    return _PNG_1X1


def _repair_docx_media(docx_bytes: bytes) -> tuple[bytes, list[str]]:
    repaired_notes: list[str] = []
    out = io.BytesIO()

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as source, zipfile.ZipFile(
        out, "w"
    ) as target:
        written_names: set[str] = set()
        for info in source.infolist():
            try:
                data = source.read(info.filename)
            except zipfile.BadZipFile:
                if info.filename.startswith("word/media/"):
                    data = _placeholder_media_bytes(info.filename)
                    repaired_notes.append(f"replaced corrupt media: {info.filename}")
                else:
                    raise
            target.writestr(info, data)
            written_names.add(info.filename)

        media_refs: set[str] = set()
        rel_paths = [
            name
            for name in written_names
            if name.startswith("word/_rels/") and name.endswith(".rels")
        ]
        for rel_path in rel_paths:
            try:
                rel_text = source.read(rel_path).decode("utf-8", errors="ignore")
            except zipfile.BadZipFile:
                continue
            for match in re.findall(r'Target="(media/[^"]+)"', rel_text):
                media_refs.add(f"word/{match}")

        for media_path in sorted(media_refs):
            if media_path in written_names:
                continue
            target.writestr(media_path, _placeholder_media_bytes(media_path))
            repaired_notes.append(f"added missing media: {media_path}")

    return out.getvalue(), repaired_notes


def _remove_markdown_images(markdown_text: str) -> str:
    # Drop Markdown and HTML image tags from generated content.
    without_md_images = re.sub(r"!?\[[^\]]*\]\([^\)]*\)", "", markdown_text)
    without_html_images = re.sub(
        r"<img\b[^>]*>", "", without_md_images, flags=re.IGNORECASE
    )
    return without_html_images


def convert_docx_to_md(docx_path: Path) -> str:
    try:
        import mammoth
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency 'mammoth'. Install it with: pip install mammoth"
        ) from exc

    docx_bytes = docx_path.read_bytes()

    repair_notes: list[str] = []
    work_bytes = docx_bytes

    for _ in range(5):
        try:
            result = mammoth.convert_to_markdown(io.BytesIO(work_bytes))
            break
        except zipfile.BadZipFile:
            work_bytes, notes = _repair_docx_media(work_bytes)
            if not notes:
                raise
            repair_notes.extend(notes)
        except KeyError as exc:
            msg = str(exc)
            missing = re.search(r"'(word/media/[^']+)'", msg)
            if not missing:
                raise
            missing_path = missing.group(1)
            with zipfile.ZipFile(io.BytesIO(work_bytes)) as source, io.BytesIO() as out:
                with zipfile.ZipFile(out, "w") as target:
                    present: set[str] = set()
                    for info in source.infolist():
                        target.writestr(info, source.read(info.filename))
                        present.add(info.filename)
                    if missing_path not in present:
                        target.writestr(missing_path, _placeholder_media_bytes(missing_path))
                        repair_notes.append(f"added missing media: {missing_path}")
                work_bytes = out.getvalue()
    else:
        raise RuntimeError("Unable to repair DOCX media entries after multiple attempts")

    markdown_text = result.value
    markdown_text = _remove_markdown_images(markdown_text)
    # Keep conversion warnings inside the output for visibility.
    if result.messages:
        notes = "\n".join(f"- {m.message}" for m in result.messages)
        markdown_text += f"\n\n<!-- conversion notes:\n{notes}\n-->\n"

    if repair_notes:
        dedup_notes = "\n".join(f"- {note}" for note in dict.fromkeys(repair_notes))
        markdown_text += f"\n\n<!-- recovery notes:\n{dedup_notes}\n-->\n"

    return markdown_text


def main() -> int:
    args = parse_args()
    root = Path(args.input_dir).expanduser().resolve()

    if not root.exists() or not root.is_dir():
        print(f"Error: directory not found: {root}", file=sys.stderr)
        return 1

    docx_files = sorted(iter_docx_files(root, args.recursive))
    if not docx_files:
        print("No DOCX files found.")
        return 0

    converted = 0
    skipped = 0
    failed = 0

    for docx_file in docx_files:
        md_file = docx_file.with_suffix(".md")

        if md_file.exists() and not args.overwrite:
            print(f"Skip (exists): {md_file}")
            skipped += 1
            continue

        try:
            md_content = convert_docx_to_md(docx_file)
            md_file.write_text(md_content, encoding="utf-8")
            print(f"OK: {docx_file} -> {md_file}")
            converted += 1
        except Exception as exc:  # noqa: BLE001
            print(f"Fail: {docx_file} ({exc})", file=sys.stderr)
            failed += 1

    print(
        f"Done. converted={converted}, skipped={skipped}, failed={failed}, total={len(docx_files)}"
    )
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
