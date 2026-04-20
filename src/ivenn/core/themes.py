from __future__ import annotations

from typing import Final

SET_ORDER: Final[tuple[str, ...]] = ("A", "B", "C", "D", "E", "F")

DEFAULT_OPACITY: Final[float] = 0.3

SET_COLOUR_THEMES = {
    "Default": {
        "_opacity": 0.3,
        "A": "#FAC32A",
        "B": "#65CE62",
        "C": "#6686F0",
        "D": "#FF6A00",
        "E": "#FFF068",
        "F": "#4ECDC4",
    },
    "Vibrant": {
        "_opacity": 0.3,
        "A": "#FF006E",
        "B": "#FB5607",
        "C": "#FFBE0B",
        "D": "#8338EC",
        "E": "#3A86FF",
        "F": "#06D6A0",
    },
    "Cool": {
        "_opacity": 0.3,
        "A": "#2D6A4F",
        "B": "#79B9F0",
        "C": "#1B4965",
        "D": "#85E382",
        "E": "#084880",
        "F": "#00B4D8",
    },
    "Bright": {
        "_opacity": 0.3,
        "A": "#FF595E",
        "B": "#FFCA3A",
        "C": "#8AC926",
        "D": "#1982C4",
        "E": "#6A4C93",
        "F": "#FF924C",
    },
    "Deep": {
        "_opacity": 0.3,
        "A": "#1B1F3B",
        "B": "#5A189A",
        "C": "#7B2CBF",
        "D": "#0F4C5C",
        "E": "#8D0801",
        "F": "#6F1D1B",
    },
    "Warm": {
        "_opacity": 0.4,
        "A": "#E76F51",
        "B": "#F4A261",
        "C": "#E9C46A",
        "D": "#D62828",
        "E": "#BC6C25",
        "F": "#FFB4A2",
    },
}


def theme_names() -> list[str]:
    """Return the available themes"""
    names = [name for name in SET_COLOUR_THEMES.keys() if not name.startswith("_")]
    names.sort(key=lambda name: (name != "Default", name))
    return names


def get_theme(name: str | None) -> dict[str, str | float]:
    """Return a theme, falling back to the Default theme if name provided does not exist."""
    if not name:
        return SET_COLOUR_THEMES["Default"]
    return SET_COLOUR_THEMES.get(name, SET_COLOUR_THEMES["Default"])


def validate_theme(theme: dict[str, str | float]) -> None:
    """Validate that a theme contains colours for A-F."""
    missing = [key for key in SET_ORDER if key not in theme]
    if missing:
        raise ValueError(f"Theme is missing set colours for: {', '.join(missing)}")