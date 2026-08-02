#!/usr/bin/env python3
"""Зеркало VBA CollapseSpacesString / бинарного схлопывания для регрессии."""

from __future__ import annotations


def collapse_spaces_string(s: str) -> str:
    """Как ReplaceMultipleSpaces.CollapseSpacesString: серии пробелов → один."""
    if len(s) < 2 or "  " not in s:
        return s
    guard = 0
    while "  " in s:
        prev = len(s)
        s = s.replace("  ", " ")
        guard += 1
        if len(s) >= prev:
            break
        if guard > len(s) + 64:
            break
    return s


def collapse_spaces_binary_count(s: str) -> tuple[str, int]:
    """Эмуляция CollapseSpacesBinary: chunk 32,16,...,2 → один пробел."""
    total = 0
    chunk = 32
    while chunk >= 2:
        needle = " " * chunk
        guard = 0
        max_guard = len(s) + 32
        while needle in s:
            s = s.replace(needle, " ", 1)
            total += chunk - 1
            guard += 1
            if guard > max_guard:
                break
        chunk //= 2
    return s, total


def simulate_multi_run(runs: list[str]) -> str:
    """
    Pass1: collapse внутри каждого run.
    Pass2: Replace('  ',' ') по склейке, пока есть двойные пробелы.
    """
    collapsed = [collapse_spaces_string(r) for r in runs]
    text = "".join(collapsed)
    guard = 0
    while "  " in text and guard < 10000:
        text = text.replace("  ", " ", 1)
        guard += 1
    return text


CASES = [
    ("", ""),
    (" ", " "),
    ("a", "a"),
    ("a b", "a b"),
    ("a  b", "a b"),
    ("a   b", "a b"),
    ("a    b", "a b"),
    ("  a  b  ", " a b "),
    ("a" + (" " * 17) + "b", "a b"),
    ("a" + (" " * 32) + "b", "a b"),
    ("a" + (" " * 100) + "b", "a b"),
    ("привет   мир", "привет мир"),
    ("a \n  b", "a \n b"),
    ("a" + "\r" + "  b", "a\r b"),
    ("no-change", "no-change"),
    ("x y z", "x y z"),
    ("   ", " "),
]


BOUNDARY_CASES = [
    (["hello ", " world"], "hello world"),
    (["a  ", "  b"], "a b"),
    (["a ", " ", " b"], "a b"),
    (["keep", " format"], "keep format"),
    (["a   ", "   b   ", "  c"], "a b c"),
]


def test_collapse_spaces_string() -> None:
    for src, expected in CASES:
        got = collapse_spaces_string(src)
        assert got == expected, repr((src, got, expected))
        assert "  " not in got or got == expected


def test_idempotent() -> None:
    for src, _ in CASES:
        once = collapse_spaces_string(src)
        twice = collapse_spaces_string(once)
        assert once == twice


def test_binary_matches_string() -> None:
    for src, expected in CASES:
        if "  " not in src:
            continue
        got, _ = collapse_spaces_binary_count(src)
        # после бинарного прохода добить обычным (как страховка в VBA)
        got = collapse_spaces_string(got)
        assert got == expected, repr((src, got, expected))


def test_run_boundaries() -> None:
    for runs, expected in BOUNDARY_CASES:
        got = simulate_multi_run(runs)
        assert got == expected, repr((runs, got, expected))
        assert "  " not in got


def test_only_space_char() -> None:
    # таб и NBSP не трогаем — как в VBA-модуле
    s = "a\t\tb"
    assert collapse_spaces_string(s) == s
    nbsp = "a" + "\u00a0\u00a0" + "b"
    assert collapse_spaces_string(nbsp) == nbsp


if __name__ == "__main__":
    test_collapse_spaces_string()
    test_idempotent()
    test_binary_matches_string()
    test_run_boundaries()
    test_only_space_char()
    print("ok: %d string cases, %d boundary cases" % (len(CASES), len(BOUNDARY_CASES)))
