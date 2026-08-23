#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the LibreOffice extension (.oxt).

    python3 build_oxt.py                 # -> dist/budget-fra-csv-1.0.0.oxt
    python3 build_oxt.py --install       # build and install with unopkg

The .oxt is a zip file containing the extension skeleton from ``oxt/`` plus
``budget_core/`` and ``office/budget_office.py`` copied into ``pythonpath/``
(LibreOffice's Python loader adds that folder to ``sys.path``).
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
import zlib

ROOT = os.path.dirname(os.path.abspath(__file__))
OXT_DIR = os.path.join(ROOT, "oxt")
CORE_DIR = os.path.join(ROOT, "budget_core")
OFFICE_FILE = os.path.join(ROOT, "office", "budget_office.py")
DIST_DIR = os.path.join(ROOT, "dist")

NAVY = (0x33, 0x49, 0x60)
ORANGE = (0xF4, 0x65, 0x24)
WHITE = (0xFF, 0xFF, 0xFF)


# ---------------------------------------------------------------------------
# Icons (drawn here so no binaries need to live in the repository)
# ---------------------------------------------------------------------------

def _png(pixels, width, height):
    """Encode RGBA rows into PNG bytes."""
    raw = b"".join(b"\x00" + bytes(row) for row in pixels)

    def chunk(tag, data):
        body = tag + data
        return (struct.pack(">I", len(data)) + body +
                struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF))

    header = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) +
            chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def _inside_rounded(x, y, size, radius):
    """True when (x, y) is inside a rounded square."""
    cx = min(max(x, radius), size - 1 - radius)
    cy = min(max(y, radius), size - 1 - radius)
    return (x - cx) ** 2 + (y - cy) ** 2 <= radius ** 2


def draw_icon(size, background, bars):
    """A small bar chart icon: rounded background plus three bars."""
    radius = max(2, size // 6)
    pixels = []
    bar_specs = [(0.18, 0.55), (0.42, 0.34), (0.66, 0.72)]  # x start, height
    for y in range(size):
        row = []
        for x in range(size):
            if not _inside_rounded(x, y, size, radius):
                row.extend((0, 0, 0, 0))
                continue
            colour = background
            for start, height in bar_specs:
                x0 = int(start * size)
                x1 = x0 + max(2, int(size * 0.16))
                y0 = int(size * (0.82 - height * 0.62))
                y1 = int(size * 0.82)
                if x0 <= x < x1 and y0 <= y < y1:
                    colour = bars
            row.extend((colour[0], colour[1], colour[2], 255))
        pixels.append(row)
    return _png(pixels, size, size)


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def read_version():
    with open(os.path.join(OXT_DIR, "description.xml"), encoding="utf-8") as fp:
        match = re.search(r'<version\s+value="([^"]+)"', fp.read())
    return match.group(1) if match else "0.0.0"


def stage(target):
    """Copy everything the extension needs into ``target``."""
    for name in os.listdir(OXT_DIR):
        source = os.path.join(OXT_DIR, name)
        destination = os.path.join(target, name)
        if os.path.isdir(source):
            shutil.copytree(source, destination)
        else:
            shutil.copy2(source, destination)

    python_path = os.path.join(target, "pythonpath")
    os.makedirs(python_path, exist_ok=True)
    shutil.copytree(CORE_DIR, os.path.join(python_path, "budget_core"),
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    shutil.copy2(OFFICE_FILE, os.path.join(python_path, "budget_office.py"))

    icons = os.path.join(target, "icons")
    os.makedirs(icons, exist_ok=True)
    with open(os.path.join(icons, "budget.png"), "wb") as fp:
        fp.write(draw_icon(42, NAVY, ORANGE))
    with open(os.path.join(icons, "budget_hc.png"), "wb") as fp:
        fp.write(draw_icon(42, (0, 0, 0), WHITE))


def build(output=None):
    version = read_version()
    os.makedirs(DIST_DIR, exist_ok=True)
    output = output or os.path.join(DIST_DIR, "budget-fra-csv-%s.oxt" % version)

    with tempfile.TemporaryDirectory() as tmp:
        staging = os.path.join(tmp, "oxt")
        os.makedirs(staging)
        stage(staging)
        if os.path.exists(output):
            os.remove(output)
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            for base, _dirs, files in os.walk(staging):
                for name in sorted(files):
                    if name.endswith(".pyc"):
                        continue
                    full = os.path.join(base, name)
                    archive.write(full, os.path.relpath(full, staging))
    return output


def install(path):
    for command in (["unopkg", "add", "-f", path],
                    ["/usr/lib/libreoffice/program/unopkg", "add", "-f", path]):
        try:
            subprocess.check_call(command)
            return True
        except FileNotFoundError:
            continue
        except subprocess.CalledProcessError as exc:
            print("unopkg fejlede: %s" % exc, file=sys.stderr)
            return False
    print("unopkg blev ikke fundet - installér .oxt-filen manuelt.", file=sys.stderr)
    return False


def main(argv=None):
    parser = argparse.ArgumentParser(description="Byg LibreOffice-udvidelsen")
    parser.add_argument("-o", "--output", help="sti til .oxt-filen")
    parser.add_argument("--install", action="store_true",
                        help="installér udvidelsen med unopkg bagefter")
    args = parser.parse_args(argv)

    path = build(args.output)
    print("Byggede %s (%.1f KB)" % (path, os.path.getsize(path) / 1024.0))
    if args.install:
        install(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
