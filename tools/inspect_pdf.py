#!/usr/bin/env python3
"""Extract PDF text/metadata and optionally render every page to PNG."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import pdfplumber
import pypdfium2 as pdfium
from pypdf import PdfReader


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("--json", type=Path)
    parser.add_argument("--text", type=Path)
    parser.add_argument("--render-dir", type=Path)
    parser.add_argument("--scale", type=float, default=2.0)
    args = parser.parse_args()

    reader = PdfReader(args.input)
    with pdfplumber.open(args.input) as pdf:
        pages = []
        text_parts = []
        for index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            pages.append({
                "page": index,
                "width_pt": round(page.width, 2),
                "height_pt": round(page.height, 2),
                "text_chars": len(text),
                "text": text,
            })
            text_parts.append(f"\n===== PAGE {index} =====\n{text}")

    result = {
        "file": str(args.input.resolve()),
        "page_count": len(reader.pages),
        "metadata": {str(k): str(v) for k, v in (reader.metadata or {}).items()},
        "pages": pages,
    }
    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.text:
        args.text.parent.mkdir(parents=True, exist_ok=True)
        args.text.write_text("".join(text_parts).lstrip(), encoding="utf-8")

    if args.render_dir:
        args.render_dir.mkdir(parents=True, exist_ok=True)
        pdf = pdfium.PdfDocument(str(args.input))
        for index in range(len(pdf)):
            bitmap = pdf[index].render(scale=args.scale)
            image = bitmap.to_pil()
            image.save(args.render_dir / f"page-{index + 1}.png")
        # A repeated render can shrink the document. Do not leave old pages
        # looking like part of the latest output. Delete only our numbered PNGs.
        target = args.render_dir.resolve()
        for old in args.render_dir.glob('page-*.png'):
            match = re.fullmatch(r'page-(\d+)\.png', old.name)
            if match and int(match[1]) > len(pdf) and old.resolve().parent == target:
                old.unlink()


if __name__ == "__main__":
    main()
