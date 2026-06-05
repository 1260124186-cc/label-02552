import pytest

from bankcheck import is_numeric, to_float


class TestIsNumeric:
    def test_int(self):
        assert is_numeric(42) is True

    def test_float(self):
        assert is_numeric(3.14) is True

    def test_zero(self):
        assert is_numeric(0) is True

    def test_negative(self):
        assert is_numeric(-100) is True

    def test_string_int(self):
        assert is_numeric('42') is True

    def test_string_float(self):
        assert is_numeric('3.14') is True

    def test_string_negative(self):
        assert is_numeric('-7.5') is True

    def test_string_with_spaces(self):
        assert is_numeric('  100  ') is True

    def test_none(self):
        assert is_numeric(None) is False

    def test_empty_string(self):
        assert is_numeric('') is False

    def test_non_numeric_string(self):
        assert is_numeric('abc') is False

    def test_mixed_string(self):
        assert is_numeric('12abc') is False

    def test_bool(self):
        assert is_numeric(True) is True

    def test_list(self):
        assert is_numeric([1, 2]) is False


class TestToFloat:
    def test_int(self):
        assert to_float(42) == 42.0

    def test_float(self):
        assert to_float(3.14) == 3.14

    def test_zero(self):
        assert to_float(0) == 0.0

    def test_string_int(self):
        assert to_float('42') == 42.0

    def test_string_float(self):
        assert to_float('3.14') == 3.14

    def test_string_with_spaces(self):
        assert to_float('  100  ') == 100.0

    def test_none(self):
        assert to_float(None) is None

    def test_empty_string(self):
        assert to_float('') is None

    def test_non_numeric_string(self):
        assert to_float('abc') is None

    def test_negative_string(self):
        assert to_float('-7.5') == -7.5
