#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Clinical photo de-identification for aesthetic case reports.

Applies a solid, full-width horizontal band over the orbital (eye) region so
patient photographs can be shared in a case deck without exposing a biometric
identifier. Accepts a single file, several files, or a directory (top level
only, so backup folders such as `_unmasked/` are never re-processed).

Examples
--------
# In-place, with untouched originals preserved (recommended before building a deck)
python3 apply_eye_mask.py ./photos --in-place --backup-dir ./photos/_unmasked

# Non-destructive: write *_masked.png into a separate directory
python3 apply_eye_mask.py baseline.png --output-dir ./masked

# Custom geometry for lateral 90-degree or unusual crops
python3 apply_eye_mask.py ./photos --top 0.30 --bottom 0.46 --output-dir ./masked

# Preview the plan without writing anything
python3 apply_eye_mask.py ./photos --in-place --dry-run
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from PIL import Image, ImageDraw

EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}


def parse_hex(value: str) -> tuple[int, int, int]:
    s = value.lstrip("#")
    if len(s) != 6:
        raise ValueError(f"color must be a 6-digit hex, got {value!r}")
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def collect(paths: list[str]) -> list[Path]:
    found: list[Path] = []
    for raw in paths:
        p = Path(raw).expanduser().resolve()
        if p.is_dir():
            found.extend(sorted(f for f in p.iterdir()
                                if f.is_file() and f.suffix.lower() in EXTS))
        elif p.is_file():
            found.append(p)
        else:
            print(f"WARN  not found: {p}", file=sys.stderr)
    return list(dict.fromkeys(found))


def band_box(size: tuple[int, int], top: float, bottom: float) -> tuple[int, int, int, int]:
    w, h = size
    y0 = max(0, min(h - 1, int(round(h * top))))
    y1 = max(y0 + 1, min(h, int(round(h * bottom))))
    return (0, y0, w, y1)


def mask_image(src: Path, dst: Path, top: float, bottom: float,
               color: str, opacity: float, radius: int) -> tuple[int, int, int, int]:
    im = Image.open(src).convert("RGBA")
    box = band_box(im.size, top, bottom)
    overlay = Image.new("RGBA", im.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    rgba = parse_hex(color) + (int(round(255 * opacity)),)
    if radius > 0:
        draw.rounded_rectangle(box, radius=radius, fill=rgba)
    else:
        draw.rectangle(box, fill=rgba)
    out = Image.alpha_composite(im, overlay).convert("RGB")

    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.suffix.lower() in {".jpg", ".jpeg"}:
        out.save(dst, "JPEG", quality=95)
    else:
        out.save(dst, "PNG")
    return box


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Apply a solid full-width eye-mask band to clinical photos.")
    ap.add_argument("inputs", nargs="+", help="image files and/or directories")
    ap.add_argument("--top", type=float, default=0.33,
                    help="band upper edge as a fraction of image height (default 0.33)")
    ap.add_argument("--bottom", type=float, default=0.49,
                    help="band lower edge as a fraction of image height (default 0.49)")
    ap.add_argument("--color", default="#1E2022",
                    help="mask color, 6-digit hex (default #1E2022 deep charcoal)")
    ap.add_argument("--opacity", type=float, default=1.0,
                    help="mask opacity 0-1; keep 1.0 for de-identification (default 1.0)")
    ap.add_argument("--radius", type=int, default=0,
                    help="corner radius in px; keep 0 so both ends bleed off-frame (default 0)")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--in-place", action="store_true",
                      help="overwrite originals after backing them up")
    mode.add_argument("--output-dir", help="write masked copies into this directory")
    ap.add_argument("--suffix", default="_masked",
                    help="filename suffix for copies (default _masked)")
    ap.add_argument("--backup-dir",
                    help="backup location for --in-place (default <first dir>/_unmasked)")
    ap.add_argument("--dry-run", action="store_true", help="print the plan without writing")
    args = ap.parse_args()

    if not (0.0 <= args.top < args.bottom <= 1.0):
        ap.error("require 0 <= --top < --bottom <= 1")
    if not (0.0 < args.opacity <= 1.0):
        ap.error("require 0 < --opacity <= 1")
    if args.opacity < 1.0 or args.radius > 0:
        print("WARN  translucent or rounded bands can leak the orbital contour; "
              "use --opacity 1.0 --radius 0 for true de-identification.", file=sys.stderr)
    parse_hex(args.color)

    files = collect(args.inputs)
    if not files:
        ap.error("no images found")

    backup_root = None
    if args.in_place:
        backup_root = (Path(args.backup_dir).expanduser().resolve()
                       if args.backup_dir else files[0].parent / "_unmasked")

    for src in files:
        if args.in_place:
            dst = src
            backup = backup_root / src.name
            if not backup.exists():
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, backup)
        elif args.output_dir:
            dst = Path(args.output_dir).expanduser().resolve() / f"{src.stem}{args.suffix}{src.suffix}"
        else:
            dst = src.with_name(f"{src.stem}{args.suffix}{src.suffix}")

        if args.dry_run:
            print(f"DRY  {src} -> {dst}")
            continue
        box = mask_image(src, dst, args.top, args.bottom,
                         args.color, args.opacity, args.radius)
        print(f"OK   {dst.name}: band y={box[1]}..{box[3]} / {Image.open(src).size[1]}px "
              f"({args.top:.0%}-{args.bottom:.0%}), color={args.color}")


if __name__ == "__main__":
    main()
