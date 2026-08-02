#!/usr/bin/env python3
"""Зеркало VBA CollapseSpacesString и двухпроходной логики CollapseMultiSpaces."""

from __future__ import annotations


def collapse_spaces_string(s: str) -> str:
    while "  " in s:
        prev = len(s)
        s = s.replace("  ", " ")
        if len(s) >= prev:
            break
    return s


def collapse_multi_spaces(runs: list[str]) -> str:
    """1 Run → строка; несколько → collapse каждого + Replace на стыках."""
    if len(runs) <= 1:
        return collapse_spaces_string(runs[0] if runs else "")
    text = "".join(collapse_spaces_string(r) for r in runs)
    before = len(text)
    guard = 0
    while "  " in text:
        prev = len(text)
        text = text.replace("  ", " ", 1)
        if len(text) >= prev:
            break
        guard += 1
        if guard > before:
            break
    return text


CASES = [
    ("", ""),
    (" ", " "),
    ("a b", "a b"),
    ("a  b", "a b"),
    ("a   b", "a b"),
    ("a    b", "a b"),
    ("  a  b  ", " a b "),
    ("a" + (" " * 17) + "b", "a b"),
    ("a" + (" " * 100) + "b", "a b"),
    ("привет   мир", "привет мир"),
    ("a \n  b", "a \n b"),
    ("a\r  b", "a\r b"),
    ("   ", " "),
]

BOUNDARY_CASES = [
    (["hello ", " world"], "hello world"),
    (["a  ", "  b"], "a b"),
    (["a ", " ", " b"], "a b"),
    (["a   ", "   b   ", "  c"], "a b c"),
    (["keep", " format"], "keep format"),
]


def test_string() -> None:
    for src, expected in CASES:
        assert collapse_spaces_string(src) == expected, (src, expected)


def test_idempotent() -> None:
    for src, _ in CASES:
        once = collapse_spaces_string(src)
        assert once == collapse_spaces_string(once)


def test_boundaries() -> None:
    for runs, expected in BOUNDARY_CASES:
        got = collapse_multi_spaces(runs)
        assert got == expected, (runs, got, expected)
        assert "  " not in got


def test_single_run_path() -> None:
    assert collapse_multi_spaces(["a   b"]) == "a b"
    assert collapse_multi_spaces([]) == ""


def test_untouched() -> None:
    assert collapse_spaces_string("a\t\tb") == "a\t\tb"
    nbsp = "a\u00a0\u00a0b"
    assert collapse_spaces_string(nbsp) == nbsp


if __name__ == "__main__":
    test_string()
    test_idempotent()
    test_boundaries()
    test_single_run_path()
    test_untouched()
    print("ok: %d string, %d boundary" % (len(CASES), len(BOUNDARY_CASES)))
