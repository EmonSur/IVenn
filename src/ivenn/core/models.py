from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable


def _normalise_elements(values: Iterable[object]) -> set[str]:
    """Convert user defined elements into a cleaned set of strings."""
    cleaned: set[str] = set()
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if not text:
            continue
        cleaned.add(text)
    return cleaned

@dataclass
class Set:
    label: str
    elements: set[str] = field(default_factory=set)
    desc: str = ""
    name: str | None = None

    def __init__(self, label: str, elements: Iterable[object], desc: str = ""):
        """Create a set with a set name, elements, and an optional description.
        
        :param label: Set name - please shorten to 20 characters.
        :param elements: Values to include in the set. Empty values are ignored and everything else is converted to strings.
        :param desc: Optional description of the set.
            
        ## Example:
        
            A = Set("A", [1, 2, 3])
        """
        self.label = str(label).strip()
        self.elements = _normalise_elements(elements)
        self.desc = str(desc).strip()
        self.name = None

    def set_description(self, new_desc: str) -> "Set":
        """Create or update the description of a set.
        
        :param new_description: Create/Update description of the ``Set``.
            
        ## Example:
        
            A = Set("A", [...])
            A.set_description("...")
        """
        self.desc = str(new_desc).strip()
        return self

    def get_description(self) -> str:
        """Return the set description."""
        return self.desc

    def union(self, other: "Set") -> set[str]:
        """Return the union of this set and another `Set`'s elements.
        
        :param other: The other ``Set`` to union with this one.
        
        ## Example
            A.union(B)
        """
        return self.elements | other.elements

    def intersection(self, other: "Set") -> set[str]:
        """Return the intersection of this set and another ``Set``'s elements.
        
        :param other: The other ``Set`` to intersect with this one.
            
        ## Example
            A.intersection(B)
        """
        return self.elements & other.elements

    def difference(self, other: "Set") -> set[str]:
        """Return the elements that are in this set but not in another ``Set``.
        
        :param other: The other ``Set`` to get the difference with this one.
            
        ## Example:
            A.difference(B)        
        """
        return self.elements - other.elements