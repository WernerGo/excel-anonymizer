"""Shared Faker dispatch used by faker_replace.py and generate_sample.py."""

from faker import Faker

_DEPARTMENTS = [
    "Engineering", "Finance", "Legal", "Marketing", "Operations",
    "Product", "Procurement", "Risk", "Sales", "Strategy", "Technology",
]

_DISPATCH = {
    "last_name":  lambda f: f.last_name(),
    "first_name": lambda f: f.first_name(),
    "full_name":  lambda f: f.name(),
    "company":    lambda f: f.company(),
    "city":       lambda f: f.city(),
    "email":      lambda f: f.email(),
    "word":       lambda f: f.word().capitalize(),
    "department": lambda f: f.random_element(_DEPARTMENTS),
}


def make_generator(faker_type: str, fake: Faker):
    """Return a callable that produces one fake value of the given type."""
    gen = _DISPATCH.get(faker_type)
    if gen is None:
        raise ValueError(
            f"Unknown faker_type '{faker_type}'. "
            f"Supported: {', '.join(_DISPATCH)}"
        )
    return lambda: gen(fake)


def infer_faker_type(group_name: str) -> str:
    """Guess a faker_type from the group name when none is configured."""
    name = group_name.lower()
    if any(w in name for w in ("last", "nachname", "surname")):
        return "last_name"
    if any(w in name for w in ("first", "vorname", "given")):
        return "first_name"
    if any(w in name for w in ("name", "person", "employee", "user", "mitarbeiter")):
        return "last_name"
    if any(w in name for w in ("department", "dept", "abteilung", "kostenstelle")):
        return "department"
    if any(w in name for w in ("city", "location", "place", "ort", "standort")):
        return "city"
    if any(w in name for w in ("company", "firma", "organisation")):
        return "company"
    return "word"
