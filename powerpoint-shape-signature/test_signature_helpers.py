#!/usr/bin/env python3
"""Tests for S2 packed signature helpers (mirrors VBA logic)."""
from __future__ import annotations
import struct
import base64

DIGITS = "0123456789abcdefghijklmnopqrstuvwxyz"
F_TYPE, F_AUTO, F_PH = 1, 2, 4
F_POS, F_SIZE, F_AR = 8, 16, 32
F_ROT, F_FLIP = 64, 128
F_FILL, F_LINE, F_CAPS = 256, 512, 1024
F_TXH, F_FONT = 2048, 4096
F_ALL = 8191


def fnv1a32_bytes(data: bytes) -> int:
    h = 0x811C9DC5
    for b in data:
        h ^= b
        h = (h * 16777619) & 0xFFFFFFFF
    return h


def b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def b64url_decode(s: str) -> bytes:
    s = s.replace("+", "-").replace("/", "_")
    pad = "=" * ((4 - len(s) % 4) % 4)
    return base64.urlsafe_b64decode(s + pad)


def pack_signature(mask: int, fields: bytes) -> str:
    body = struct.pack("<H", mask & 0xFFFF) + fields
    crc = fnv1a32_bytes(body) & 0xFFFFFF
    blob = body + bytes((crc & 0xFF, (crc >> 8) & 0xFF, (crc >> 16) & 0xFF))
    return "S2|" + b64url_encode(blob)


def unpack_signature(sig: str) -> tuple[int, bytes]:
    assert sig.startswith("S2|")
    blob = b64url_decode(sig[3:])
    assert len(blob) >= 5
    body, crc_bytes = blob[:-3], blob[-3:]
    crc = crc_bytes[0] | (crc_bytes[1] << 8) | (crc_bytes[2] << 16)
    assert (fnv1a32_bytes(body) & 0xFFFFFF) == crc
    mask = struct.unpack_from("<H", body, 0)[0]
    return mask, body[2:]


def test_roundtrip_typical():
    mask = F_TYPE | F_AUTO | F_SIZE | F_AR
    fields = struct.pack("<B", 17) + struct.pack("<h", 1) + struct.pack("<2h", 4000, 1500) + struct.pack("<i", 1600)
    sig = pack_signature(mask, fields)
    assert len(sig) < 100, len(sig)
    m2, f2 = unpack_signature(sig)
    assert m2 == mask
    assert f2 == fields


def test_roundtrip_full():
    mask = F_ALL
    fields = b"".join([
        struct.pack("<B", 17),
        struct.pack("<h", 1),
        struct.pack("<h", -1),
        struct.pack("<2h", 1200, 800),
        struct.pack("<2h", 4000, 1500),
        struct.pack("<i", 2666),
        struct.pack("<h", 0),
        struct.pack("<B", 0),
        struct.pack("<h", 1) + struct.pack("<i", 0xFF8800) + struct.pack("<h", -1),
        struct.pack("<B", 1) + struct.pack("<i", 0) + struct.pack("<h", 100),
        struct.pack("<B", 1),
        struct.pack("<i", 0x12345678),
        struct.pack("<i", 0x11ABCDEF) + struct.pack("<h", 30),
    ])
    sig = pack_signature(mask, fields)
    assert len(sig) < 100, (len(sig), sig)
    m2, f2 = unpack_signature(sig)
    assert m2 == mask
    assert f2 == fields


def test_geo_split_idea():
    # same size different pos => size units equal
    GEO = 10000
    w1 = int(round((200 / 720) * GEO))
    w2 = int(round((200 / 960) * GEO))  # different slide width same pt width => different relative!
    # relative encoding IS slide-size independent for same fraction:
    w3 = int(round((0.25) * GEO))
    w4 = int(round((0.25) * GEO))
    assert w3 == w4 == 2500


def test_tamper_fails():
    mask = F_TYPE
    fields = struct.pack("<B", 17)
    sig = pack_signature(mask, fields)
    bad = sig[:-1] + ("A" if sig[-1] != "A" else "B")
    try:
        unpack_signature(bad)
        raise AssertionError("tamper should fail")
    except AssertionError as e:
        if "tamper" in str(e):
            raise


if __name__ == "__main__":
    test_roundtrip_typical()
    test_roundtrip_full()
    test_geo_split_idea()
    test_tamper_fails()
    print("OK: S2 helper tests passed")
