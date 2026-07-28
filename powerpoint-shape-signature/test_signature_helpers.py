#!/usr/bin/env python3
"""Round-trip tests for ShapeSignature encode helpers (mirrors VBA logic)."""
from __future__ import annotations

DIGITS = "0123456789abcdefghijklmnopqrstuvwxyz"


def to_b36(n: int) -> str:
    if n == 0:
        return "0"
    neg = n < 0
    v = abs(n)
    out = []
    while v > 0:
        out.append(DIGITS[v % 36])
        v //= 36
    s = "".join(reversed(out))
    return "-" + s if neg else s


def from_b36(s: str) -> int:
    s = s.strip().lower()
    neg = s.startswith("-")
    if neg:
        s = s[1:]
    v = 0
    for c in s:
        p = DIGITS.index(c)
        v = v * 36 + p
    return -v if neg else v


def fnv1a_b36(s: str) -> str:
    h = 0x811C9DC5
    for ch in s:
        h ^= ord(ch) & 0xFF
        h = (h * 16777619) & 0xFFFFFFFF
    return to_b36(h)  # treat as unsigned via Python int; to_b36 on values > 2^31


def to_b36_unsigned(uh: int) -> str:
    if uh <= 0:
        return "0"
    out = []
    while uh >= 1:
        out.append(DIGITS[uh % 36])
        uh //= 36
    return "".join(reversed(out))


def fnv1a_b36_unsigned(s: str) -> str:
    h = 0x811C9DC5
    for ch in s:
        h ^= ord(ch) & 0xFF
        h = (h * 16777619) & 0xFFFFFFFF
    return to_b36_unsigned(h)


def build_signature(mask: int, fields: str) -> str:
    body = f"S1|{to_b36(mask)}|{fields}"
    return body + "|" + fnv1a_b36_unsigned(body)


def parse_signature(sig: str) -> tuple[int, str]:
    p = sig.split("|")
    assert len(p) == 4, p
    assert p[0] == "S1"
    body = "|".join(p[:3])
    assert fnv1a_b36_unsigned(body) == p[3]
    return from_b36(p[1]), p[2]


def test_b36_roundtrip():
    for n in [0, 1, 35, 36, 4095, 10000, -5, 2147483647]:
        assert from_b36(to_b36(n)) == n, n


def test_signature_roundtrip():
    fields = "d;1;2c8.1k4.5u.3c;rs;0"
    mask = 0x2F
    sig = build_signature(mask, fields)
    m2, f2 = parse_signature(sig)
    assert m2 == mask
    assert f2 == fields
    # tamper
    bad = sig[:-1] + ("0" if sig[-1] != "0" else "1")
    try:
        parse_signature(bad)
        raise AssertionError("tamper should fail")
    except AssertionError as e:
        if "tamper" in str(e):
            raise
        pass


def test_geo_quantization():
    # left=72pt on 720pt slide => 10% => 1000 units of 10000
    GEO_SCALE = 10000
    left, sw = 72.0, 720.0
    q = int(round((left / sw) * GEO_SCALE))
    assert q == 1000
    # independent of absolute slide size: same fraction
    left2, sw2 = 100.0, 1000.0
    q2 = int(round((left2 / sw2) * GEO_SCALE))
    assert q2 == 1000


if __name__ == "__main__":
    test_b36_roundtrip()
    test_signature_roundtrip()
    test_geo_quantization()
    print("OK: all ShapeSignature helper tests passed")
