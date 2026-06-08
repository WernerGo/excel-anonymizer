"""Shared Faker dispatch used by faker_replace.py and generate_sample.py."""

from faker import Faker
from faker.exceptions import UniquenessException

_DEPARTMENTS = [
    "Engineering", "Finance", "Legal", "Marketing", "Operations",
    "Product", "Procurement", "Risk", "Sales", "Strategy", "Technology",
]

# Maps faker_type → (unique_attr, base_callable)
# unique_attr: name of the attribute on fake.unique (None if not supported)
_TYPES: dict[str, tuple[str | None, callable]] = {
    "last_name":  ("last_name",  lambda f: f.last_name()),
    "first_name": ("first_name", lambda f: f.first_name()),
    "full_name":  ("name",       lambda f: f.name()),
    "company":    ("company",    lambda f: f.company()),
    "city":       ("city",       lambda f: f.city()),
    "email":      ("email",      lambda f: f.email()),
    "word":       ("word",       lambda f: f.word().capitalize()),
    "department": (None,         lambda f: f.random_element(_DEPARTMENTS)),
}


def make_generator(faker_type: str, fake: Faker):
    """
    Return a callable that produces one fake value of the given type.

    Uses fake.unique.<type>() to guarantee no two original values get the
    same replacement. If Faker's pool for that type is exhausted (more unique
    source values than Faker has names), falls back to base_value_N.
    """
    entry = _TYPES.get(faker_type)
    if entry is None:
        raise ValueError(
            f"Unknown faker_type '{faker_type}'. "
            f"Supported: {', '.join(_TYPES)}"
        )

    unique_attr, base_fn = entry
    overflow_counter = [0]

    def generate() -> str:
        if unique_attr is not None:
            try:
                return getattr(fake.unique, unique_attr)()
            except UniquenessException:
                # Faker pool exhausted — append counter to guarantee uniqueness
                overflow_counter[0] += 1
                return f"{base_fn(fake)}_{overflow_counter[0]}"
        # Types without unique support (e.g. department uses random_element)
        return base_fn(fake)

    return generate


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
    if any(w in name for w in ("compan", "firma", "organisation")):
        return "company"
    return "word"
