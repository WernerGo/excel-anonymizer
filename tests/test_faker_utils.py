"""Tests for faker_utils.py: generator dispatch and type inference."""

import pytest
from faker import Faker

from faker_utils import infer_faker_type, make_generator


fake = Faker()


# ---------------------------------------------------------------------------
# make_generator
# ---------------------------------------------------------------------------

def test_make_generator_returns_callable():
    gen = make_generator("last_name", fake)
    assert callable(gen)


def test_make_generator_returns_string():
    for faker_type in ("last_name", "first_name", "full_name", "company", "city",
                       "department", "email", "word"):
        gen = make_generator(faker_type, fake)
        assert isinstance(gen(), str)


def test_make_generator_raises_on_unknown_type():
    with pytest.raises(ValueError, match="Unknown faker_type"):
        make_generator("banana", fake)


def test_generator_produces_unique_values():
    """Consecutive calls must not produce the same value (uniqueness proxy)."""
    gen = make_generator("last_name", Faker())  # fresh Faker instance per test
    results = {gen() for _ in range(20)}
    assert len(results) == 20


# ---------------------------------------------------------------------------
# infer_faker_type
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("group_name, expected", [
    ("last_names",       "last_name"),
    ("employee_surname", "last_name"),
    ("nachname",         "last_name"),
    ("first_names",      "first_name"),
    ("vorname",          "first_name"),
    ("person_names",     "last_name"),
    ("mitarbeiter",      "last_name"),
    ("departments",      "department"),
    ("abteilung",        "department"),
    ("kostenstellen",    "department"),
    ("locations",        "city"),
    ("standorte",        "city"),
    ("companies",        "company"),
    ("firma",            "company"),
    ("unknown_group",    "word"),
])
def test_infer_faker_type(group_name, expected):
    assert infer_faker_type(group_name) == expected
