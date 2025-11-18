import pytest
import sys
import os
import re

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from excelxml import interval_regex, digit_set, char_set

@pytest.mark.parametrize("pos1, pos2, ref_set, zero_char, expected", [
    # Numeric tests
    ("5", "30", digit_set[0], digit_set[1], '(?:30|[1-2][0-9]|[5-9])'),
    ("1", "9", digit_set[0], digit_set[1], '[1-9]'),
    ("10", "10", digit_set[0], digit_set[1], '10'),
    ("95", "105", digit_set[0], digit_set[1], '(?:10[0-5]|9[5-9])'),
    ("95", "1000", digit_set[0], digit_set[1], '(?:1000|[1-9][0-9]{2}|9[5-9])'),
    ("1", "100", digit_set[0], digit_set[1], '(?:100|[1-9][0-9]|[1-9])'),

    # Character tests
    ("A", "C", char_set[0], char_set[1], '[A-C]'),
    ("A", "A", char_set[0], char_set[1], 'A'),
    ("Y", "AC", char_set[0], char_set[1], '(?:A[A-C]|[Y-Z])'),
    ("X", "AB", char_set[0], char_set[1], '(?:A[A-B]|[X-Z])'),
    ("A", "Z", char_set[0], char_set[1], '[A-Z]'),
    ("AA", "AZ", char_set[0], char_set[1], '(?:A[A-Z])'),
    ("A", "AZ", char_set[0], char_set[1], '(?:A[A-Z]|[A-Z])'),
])
def test_interval_regex(pos1, pos2, ref_set, zero_char, expected):
    """
    Tests the interval_regex function with various inputs.
    """
    regex_pattern = interval_regex(pos1, pos2, ref_set, zero_char)
    assert regex_pattern == expected

    # Compile the regex to check for validity
    try:
        crgx = re.compile(regex_pattern)
    except re.error as e:
        pytest.fail(f"Regex compilation failed for pattern '{regex_pattern}': {e}")

    # Basic validation of generated regex
    ## Check if numbers in the range boundaries match
    assert all(map(crgx.fullmatch, (pos1, pos2)))
    if pos1 != pos2:
        # test pos1 + 1, pos2 + 1
        cases = [(pos1, True), (pos2, False)]
        for pos, flag in cases:
            head, tail = pos, ''
            while head and head[-1] == ref_set[-1]:
                tail = ref_set[0] + tail
                head = head[:-1]
            head = head or ref_set[0]
            ndx = ref_set.index(head[-1]) + 1
            head = f'{head[:-1]}{ref_set[ndx]}'
            item = f'{head}{tail}'.lstrip(zero_char)
            assert bool(crgx.fullmatch(item)) == flag

        # test pos1 - 1 y pos2 - 1
        cases = [(pos2, True)]
        if pos1 != ref_set[0]:
            cases = [(pos1, False), (pos2, True)]
        for pos, flag in cases:
            head, tail = pos, ''
            while head and head[-1] == ref_set[0]:
                tail = ref_set[-1] + tail
                head = head[:-1]
            head = head or ref_set[1]
            ndx = ref_set.index(head[-1]) - 1
            head = f'{head[:-1]}{ref_set[ndx]}'
            item = f'{head}{tail}'.lstrip(zero_char)
            assert bool(crgx.fullmatch(item)) == flag
