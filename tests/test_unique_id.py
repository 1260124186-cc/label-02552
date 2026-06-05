import pytest

from bankcheck import generate_unique_id


class TestGenerateUniqueId:
    def test_returns_string(self):
        uid = generate_unique_id()
        assert isinstance(uid, str)

    def test_non_empty(self):
        uid = generate_unique_id()
        assert len(uid) > 0

    def test_uniqueness(self):
        ids = {generate_unique_id() for _ in range(100)}
        assert len(ids) == 100

    def test_starts_with_timestamp(self):
        uid = generate_unique_id()
        timestamp_part = uid[:14]
        assert timestamp_part.isdigit()

    def test_uuid_part_is_hex(self):
        uid = generate_unique_id()
        uuid_part = uid[20:]
        assert len(uuid_part) == 32
        assert all(c in '0123456789abcdef' for c in uuid_part)

    def test_total_length(self):
        uid = generate_unique_id()
        assert len(uid) == 20 + 32
