#!/usr/bin/env python3
"""Classify Markdown files into class folders using content-based detection.

By default, files are copied into class folders under the input directory.
Use --move to move files instead.

Examples:
    python classify_md_by_class.py . --dry-run
    python classify_md_by_class.py . --move
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

DEFAULT_CLASSES = [
    "2025140902",
    "2024091603",
    "2024170401",
    "2024140903",
    "2025140901",
    "2025140903",
    "2024170101",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify Markdown files to class folders by content."
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=".",
        help="Directory containing Markdown files (default: current directory).",
    )
    parser.add_argument(
        "--class",
        dest="classes",
        action="append",
        help="Class code to match. Repeat to add multiple (default is built-in class list).",
    )
    parser.add_argument(
        "--move",
        action="store_true",
        help="Move files instead of copying.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview classification without creating/changing files.",
    )
    return parser.parse_args()


def normalize_for_match(text: str) -> str:
    # Remove all whitespace so variants like "本 2024091603" are still detected.
    return "".join(text.split())


def extract_digits(text: str) -> str:
    return "".join(ch for ch in text if ch.isdigit())


def canonical_class(raw: str) -> str:
    digits = extract_digits(raw)
    return f"本{digits}"


def detect_class(file_name: str, text: str, classes: list[str]) -> list[str]:
    source = f"{file_name}\n{text}"
    hits = [
        code for code in classes if re.search(rf"(?<!\\d){re.escape(code)}(?!\\d)", source)
    ]
    return hits


def iter_md_files(root: Path) -> list[Path]:
    return sorted(
        [
            p
            for p in root.glob("*.md")
            if p.is_file() and p.name != Path(__file__).name
        ]
    )


def classify_file(md_file: Path, target_dir: Path, move: bool, dry_run: bool) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    target_file = target_dir / md_file.name

    if dry_run:
        action = "MOVE" if move else "COPY"
        print(f"{action}: {md_file} -> {target_file}")
        return

    if move:
        shutil.move(str(md_file), str(target_file))
    else:
        shutil.copy2(md_file, target_file)


def main() -> int:
    args = parse_args()
    root = Path(args.input_dir).expanduser().resolve()

    if not root.exists() or not root.is_dir():
        print(f"Error: directory not found: {root}", file=sys.stderr)
        return 1

    classes_input = args.classes if args.classes else DEFAULT_CLASSES
    class_numbers = [extract_digits(c) for c in classes_input]
    class_numbers = [c for c in class_numbers if c]

    if not class_numbers:
        print("Error: no valid numeric class codes provided.", file=sys.stderr)
        return 1

    md_files = iter_md_files(root)
    if not md_files:
        print("No Markdown files found in top-level directory.")
        return 0

    matched = 0
    unmatched = 0
    ambiguous = 0

    for md_file in md_files:
        try:
            content = md_file.read_text(encoding="utf-8")
        except Exception as exc:  # noqa: BLE001
            print(f"READ_FAIL: {md_file} ({exc})", file=sys.stderr)
            unmatched += 1
            continue

        hits = detect_class(md_file.name, content, class_numbers)

        if len(hits) == 1:
            class_code = canonical_class(hits[0])
            classify_file(md_file, root / class_code, move=args.move, dry_run=args.dry_run)
            matched += 1
            continue

        if len(hits) > 1:
            print(f"AMBIGUOUS: {md_file} -> {', '.join(hits)}", file=sys.stderr)
            ambiguous += 1
            continue

        print(f"UNMATCHED: {md_file}")
        unmatched += 1

    print(
        f"Done. matched={matched}, unmatched={unmatched}, ambiguous={ambiguous}, total={len(md_files)}"
    )
    return 0 if ambiguous == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
