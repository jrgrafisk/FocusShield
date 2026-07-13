#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal pure-Python decoder for the AOSP static (Ver2) binary dictionary format.

Used only for VALIDATION: it walks the PtNode trie of a compiled .dict and yields
every terminal word with its frequency, so we can prove that the words we put in
actually came back out of the binary (round-trip check for Phase 7).

Only the features present in our generated dictionaries are supported:
single-byte code points (all Danish letters are U+0020..U+00FF), terminals with a
1-byte frequency, and children addresses. No shortcuts/bigrams/dynamic-update.
"""

import sys
from pathlib import Path

MAGIC = 0x9BC13AFE
PTNODE_TERMINATOR = 0x1F
FLAG_CHILDREN_MASK = 0xC0
FLAG_CHILDREN_NOADDRESS = 0x00
FLAG_CHILDREN_ONEBYTE = 0x40
FLAG_CHILDREN_TWOBYTES = 0x80
FLAG_CHILDREN_THREEBYTES = 0xC0
FLAG_HAS_MULTIPLE_CHARS = 0x20
FLAG_IS_TERMINAL = 0x10
FLAG_HAS_SHORTCUT = 0x08
FLAG_HAS_BIGRAMS = 0x04

# Bigram/shortcut attribute encoding (used only to skip past them).
FLAG_ATTR_HAS_NEXT = 0x80
MASK_BIGRAM_ATTR_ADDRESS_TYPE = 0x30
BIGRAM_ADDR_ONEBYTE = 0x10
BIGRAM_ADDR_TWOBYTES = 0x20
BIGRAM_ADDR_THREEBYTES = 0x30
SHORTCUT_LIST_SIZE_SIZE = 2


class Reader:
    def __init__(self, data: bytes):
        self.d = data
        self.p = 0

    def u8(self):
        v = self.d[self.p]
        self.p += 1
        return v

    def u16(self):
        return (self.u8() << 8) | self.u8()

    def u24(self):
        return (self.u8() << 16) | (self.u8() << 8) | self.u8()


def read_header(data: bytes):
    magic = int.from_bytes(data[0:4], "big")
    if magic != MAGIC:
        raise ValueError(f"bad magic 0x{magic:08X}, expected 0x{MAGIC:08X}")
    version = int.from_bytes(data[4:6], "big")
    flags = int.from_bytes(data[6:8], "big")
    header_size = int.from_bytes(data[8:12], "big")
    # Header attributes (key/value UTF-16BE strings) live in data[12:header_size];
    # we skip them for word extraction but decode for the report.
    attrs = {}
    try:
        # Header attributes are UTF-8 strings separated by 0x1F (unit separator),
        # laid out as key, value, key, value ...
        body = data[12:header_size]
        parts = [p.decode("utf-8", "replace") for p in body.split(b"\x1f") if p]
        for k in range(0, len(parts) - 1, 2):
            attrs[parts[k]] = parts[k + 1]
    except Exception:
        pass
    return version, flags, header_size, attrs


def decode(path: Path):
    data = Path(path).read_bytes()
    version, flags, header_size, attrs = read_header(data)
    words = []

    # Children addresses in Ver2 are stored relative to the file position where
    # the children-address field itself begins: child_abs = addr_field_pos + value.
    # We therefore work in absolute file offsets throughout.
    def parse_node_array(abs_pos, prefix):
        r = Reader(data)
        r.p = abs_pos
        count_first = r.u8()
        if count_first & 0x80:
            count = ((count_first & 0x7F) << 8) | r.u8()
        else:
            count = count_first
        for _ in range(count):
            flagbyte = r.u8()
            # characters
            chars = []
            if flagbyte & FLAG_HAS_MULTIPLE_CHARS:
                while True:
                    c = r.u8()
                    if c == PTNODE_TERMINATOR:
                        break
                    if c < 0x20:  # 3-byte code point (not expected in our data)
                        c = (c << 16) | r.u16()
                    chars.append(chr(c))
            else:
                c = r.u8()
                if c < 0x20:
                    c = (c << 16) | r.u16()
                chars.append(chr(c))
            word_so_far = prefix + "".join(chars)
            freq = None
            if flagbyte & FLAG_IS_TERMINAL:
                freq = r.u8()
            # children address: relative to the position where its bytes start
            ctype = flagbyte & FLAG_CHILDREN_MASK
            child_abs = None
            addr_field_pos = r.p
            if ctype == FLAG_CHILDREN_ONEBYTE:
                child_abs = addr_field_pos + r.u8()
            elif ctype == FLAG_CHILDREN_TWOBYTES:
                child_abs = addr_field_pos + r.u16()
            elif ctype == FLAG_CHILDREN_THREEBYTES:
                child_abs = addr_field_pos + r.u24()
            # Shortcut list: 2-byte total size (incl. the size field) -> skip block.
            if flagbyte & FLAG_HAS_SHORTCUT:
                sc_start = r.p
                sc_size = r.u16()
                r.p = sc_start + sc_size
            # Bigram list: flag-chained entries, each = 1 flags byte + 1/2/3 addr
            # bytes; continue while HAS_NEXT bit set. We only skip them.
            if flagbyte & FLAG_HAS_BIGRAMS:
                while True:
                    bflags = r.u8()
                    atype = bflags & MASK_BIGRAM_ATTR_ADDRESS_TYPE
                    if atype == BIGRAM_ADDR_ONEBYTE:
                        r.p += 1
                    elif atype == BIGRAM_ADDR_TWOBYTES:
                        r.p += 2
                    elif atype == BIGRAM_ADDR_THREEBYTES:
                        r.p += 3
                    if not (bflags & FLAG_ATTR_HAS_NEXT):
                        break
            if freq is not None:
                words.append((word_so_far, freq))
            if child_abs is not None:
                parse_node_array(child_abs, word_so_far)

    parse_node_array(header_size, "")
    return {"version": version, "flags": flags, "header_size": header_size,
            "attrs": attrs, "words": words}


if __name__ == "__main__":
    res = decode(Path(sys.argv[1]))
    out = sys.argv[2] if len(sys.argv) > 2 else None
    print(f"version={res['version']} header_size={res['header_size']} "
          f"words={len(res['words'])}")
    print("attrs:", res["attrs"])
    ws = sorted(w for w, _ in res["words"])
    if out:
        Path(out).write_text("\n".join(ws) + "\n", encoding="utf-8")
        print("wrote", out)
    else:
        for w in ws[:20]:
            print(" ", w)
