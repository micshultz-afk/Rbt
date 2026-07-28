#!/usr/bin/env python3
"""Mirror of the VBA S3 signature format: validates layout, length and FNV math.

Run: python3 test_signature_helpers.py
"""
from __future__ import annotations
import base64

# ---- feature mask bits (must match ShapeSignature.bas) ----
R_ZONE, R_SIZE, R_AR, R_TYPE = 1, 2, 4, 8
R_PH, R_FLAGS, R_TXT, R_FONTREL, R_ROT = 16, 32, 64, 128, 256
R_ALL = 511
SIG_VER_NUM = 3


# ---------------- FNV-1a 32 ----------------
def fnv_mul_reference(h: int) -> int:
    return (h * 16777619) & 0xFFFFFFFF


def fnv_mul_split(h: int) -> int:
    """Mirror of VBA FnvMul: 16-bit halves so Double stays exact."""
    hi = h // 65536
    lo = h - hi * 65536
    r = (lo * 16777619) % 4294967296
    t = (hi * 16777619) % 65536
    r = (r + t * 65536) % 4294967296
    return r


def fnv1a_bytes(data: bytes) -> int:
    h = 0x811C9DC5
    for b in data:
        h ^= b
        h = fnv_mul_split(h)
    return h


# ---------------- packing ----------------
def u8(v: int) -> bytes:
    return bytes((v & 0xFF,))


def u16(v: int) -> bytes:
    return bytes((v & 0xFF, (v >> 8) & 0xFF))


def u32(v: int) -> bytes:
    v &= 0xFFFFFFFF
    return bytes((v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF))


def b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def unb64url(s: str) -> bytes:
    s = s.replace("+", "-").replace("/", "_")
    return base64.urlsafe_b64decode(s + "=" * ((4 - len(s) % 4) % 4))


def encode(role: dict) -> str:
    mask = role["mask"]
    has_format = "format" in role
    body = b""
    body += u8(SIG_VER_NUM)
    body += u8(1 | (2 if has_format else 0))
    body += u16(mask)
    body += u8(role.get("nBad", 0))
    body += u8(role.get("nGood", 0))
    body += u8(role.get("thr", 60))

    if mask & R_ZONE:
        body += u8(role["CxC"]) + u8(role["CxH"]) + u8(role["CyC"]) + u8(role["CyH"])
    if mask & R_SIZE:
        body += u8(role["WC"]) + u8(role["WH"]) + u8(role["HC"]) + u8(role["HH"])
    if mask & R_AR:
        body += u8(role["ARC"]) + u8(role["ARH"])
    if mask & R_TYPE:
        body += u32(role["TypeSet"])
    if mask & R_PH:
        body += u16(role["PhSet"])
    if mask & R_FLAGS:
        body += u8(role["FlagsSeen"]) + u8(role["FlagsAll"])
    if mask & R_TXT:
        body += u8(role["ParaC"]) + u8(role["ParaH"]) + u8(role["LenSet"])
    if mask & R_FONTREL:
        body += u8(role["FontRelC"]) + u8(role["FontRelH"])
    if mask & R_ROT:
        body += u8(role["RotC"]) + u8(role["RotH"])

    if has_format:
        f = role["format"]
        body += u32(f["FontFnv"])
        body += u8(f["FontRelC"]) + u8(f["FontRelH"])
        body += u8(f["FontBucketSet"]) + u8(f["FillBucketSet"])
        body += u8(f["StyleFlags"]) + u8(f["AlignSet"])
        body += u8(f["CxC"]) + u8(f["CxH"]) + u8(f["CyC"]) + u8(f["CyH"])

    crc = fnv1a_bytes(body) & 0xFFFFFF
    blob = body + bytes((crc & 0xFF, (crc >> 8) & 0xFF, (crc >> 16) & 0xFF))
    return "S3|" + b64url(blob)


def decode(sig: str) -> dict:
    prefix, payload = sig.split("|", 1)
    assert prefix == "S3"
    blob = unb64url(payload)
    assert len(blob) >= 10
    body, crc_bytes = blob[:-3], blob[-3:]
    crc = crc_bytes[0] | (crc_bytes[1] << 8) | (crc_bytes[2] << 16)
    assert (fnv1a_bytes(body) & 0xFFFFFF) == crc, "CRC mismatch"

    pos = 0

    def rd8() -> int:
        nonlocal pos
        v = body[pos]
        pos += 1
        return v

    def rd16() -> int:
        nonlocal pos
        v = body[pos] | (body[pos + 1] << 8)
        pos += 2
        return v

    def rd32() -> int:
        nonlocal pos
        v = body[pos] | (body[pos + 1] << 8) | (body[pos + 2] << 16) | (body[pos + 3] << 24)
        pos += 4
        return v

    assert rd8() == SIG_VER_NUM
    flags = rd8()
    out: dict = {"mask": rd16()}
    assert out["mask"] != 0 and (out["mask"] & ~R_ALL) == 0
    out["nBad"] = rd8()
    out["nGood"] = rd8()
    out["thr"] = rd8()
    mask = out["mask"]

    if mask & R_ZONE:
        out["CxC"], out["CxH"], out["CyC"], out["CyH"] = rd8(), rd8(), rd8(), rd8()
    if mask & R_SIZE:
        out["WC"], out["WH"], out["HC"], out["HH"] = rd8(), rd8(), rd8(), rd8()
    if mask & R_AR:
        out["ARC"], out["ARH"] = rd8(), rd8()
    if mask & R_TYPE:
        out["TypeSet"] = rd32()
    if mask & R_PH:
        out["PhSet"] = rd16()
    if mask & R_FLAGS:
        out["FlagsSeen"], out["FlagsAll"] = rd8(), rd8()
    if mask & R_TXT:
        out["ParaC"], out["ParaH"], out["LenSet"] = rd8(), rd8(), rd8()
    if mask & R_FONTREL:
        out["FontRelC"], out["FontRelH"] = rd8(), rd8()
    if mask & R_ROT:
        out["RotC"], out["RotH"] = rd8(), rd8()

    if flags & 2:
        out["format"] = {
            "FontFnv": rd32(),
            "FontRelC": rd8(),
            "FontRelH": rd8(),
            "FontBucketSet": rd8(),
            "FillBucketSet": rd8(),
            "StyleFlags": rd8(),
            "AlignSet": rd8(),
            "CxC": rd8(),
            "CxH": rd8(),
            "CyC": rd8(),
            "CyH": rd8(),
        }

    assert pos == len(body), (pos, len(body))
    return out


# ---------------- tests ----------------
def test_fnv_split_matches_reference():
    h = 0x811C9DC5
    for b in b"footer label calibri":
        h ^= b
        assert fnv_mul_split(h) == fnv_mul_reference(h)
        h = fnv_mul_split(h)


def test_footer_signature_is_short():
    """Typical footer class: zone + size + type + flags + text structure."""
    role = {
        "mask": R_ZONE | R_SIZE | R_TYPE | R_PH | R_FLAGS | R_TXT,
        "nBad": 6,
        "nGood": 3,
        "thr": 72,
        "CxC": 40, "CxH": 12, "CyC": 236, "CyH": 8,
        "WC": 60, "WH": 14, "HC": 12, "HH": 4,
        "TypeSet": (1 << 16) | (1 << 13),   # msoTextBox / msoPlaceholder
        "PhSet": (1 << 15) | (1 << 4),
        "FlagsSeen": 0b01100111, "FlagsAll": 0b00000011,
        "ParaC": 1, "ParaH": 1, "LenSet": 0b00000110,
        "format": {
            "FontFnv": 0x1A2B3C4D,
            "FontRelC": 22, "FontRelH": 2,
            "FontBucketSet": 0b00000100,
            "FillBucketSet": 0b00000010,
            "StyleFlags": 0,
            "AlignSet": 0b00000010,
            "CxC": 40, "CxH": 4, "CyC": 236, "CyH": 3,
        },
    }
    sig = encode(role)
    assert len(sig) < 100, (len(sig), sig)
    assert decode(sig) == role
    print(f"  footer class: {len(sig)} chars -> {sig}")


def test_worst_case_all_features():
    role = {
        "mask": R_ALL,
        "nBad": 255, "nGood": 255, "thr": 85,
        "CxC": 128, "CxH": 90, "CyC": 250, "CyH": 90,
        "WC": 200, "WH": 60, "HC": 40, "HH": 60,
        "ARC": 120, "ARH": 38,
        "TypeSet": 0x3FFFFFFF,
        "PhSet": 0xFFFF,
        "FlagsSeen": 0xFF, "FlagsAll": 0xFF,
        "ParaC": 255, "ParaH": 29, "LenSet": 0xFF,
        "FontRelC": 255, "FontRelH": 39,
        "RotC": 180, "RotH": 19,
        "format": {
            "FontFnv": 0xFFFFFFFF,
            "FontRelC": 255, "FontRelH": 199,
            "FontBucketSet": 0xFF, "FillBucketSet": 0xFF,
            "StyleFlags": 0xFF, "AlignSet": 0xFF,
            "CxC": 255, "CxH": 255, "CyC": 255, "CyH": 255,
        },
    }
    sig = encode(role)
    assert len(sig) < 100, (len(sig), sig)
    assert decode(sig) == role
    print(f"  worst case:   {len(sig)} chars")


def test_role_only_no_format():
    role = {
        "mask": R_ZONE | R_TYPE,
        "nBad": 1, "nGood": 0, "thr": 65,
        "CxC": 10, "CxH": 6, "CyC": 240, "CyH": 6,
        "TypeSet": 1 << 16,
    }
    sig = encode(role)
    assert "format" not in decode(sig)
    assert len(sig) < 60
    print(f"  role only:    {len(sig)} chars -> {sig}")


def test_corrupted_signature_rejected():
    role = {"mask": R_ZONE, "CxC": 1, "CxH": 2, "CyC": 3, "CyH": 4}
    sig = encode(role)
    bad = sig[:-1] + ("A" if sig[-1] != "A" else "B")
    try:
        decode(bad)
    except AssertionError:
        return
    raise AssertionError("corrupted signature must be rejected")


def test_geometry_is_resolution_independent():
    """Same relative placement on different slide sizes -> same stored value."""
    for slide_w, left, width in ((720.0, 72.0, 360.0), (960.0, 96.0, 480.0)):
        cx = round(((left + width / 2) / slide_w) * 255)
        assert cx == 89, cx


if __name__ == "__main__":
    print("S3 signature format tests:")
    test_fnv_split_matches_reference()
    test_footer_signature_is_short()
    test_worst_case_all_features()
    test_role_only_no_format()
    test_corrupted_signature_rejected()
    test_geometry_is_resolution_independent()
    print("OK: all tests passed")
