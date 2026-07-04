"""
fonts.py
Small helper to use IRANSansX where available, with a sensible
fallback on machines that don't have it installed.

Note: IRANSansX is a licensed commercial font, so it is not bundled
in this repository. Install it on the machine running the app
(Control Panel -> Fonts on Windows) to get the intended look; the
app will automatically fall back to Tahoma otherwise, which also
renders Persian text well.
"""

import tkinter.font as tkfont

PREFERRED_FONT = "IRANSansX"
FALLBACKS = ["IRANSansX", "IRANYekan", "Tahoma", "Segoe UI", "Arial"]

_resolved_family = None


def resolve_family(root=None):
    """Return the first available font family from the fallback list."""
    global _resolved_family
    if _resolved_family:
        return _resolved_family

    try:
        available = set(tkfont.families(root))
    except Exception:
        available = set()

    for name in FALLBACKS:
        if name in available:
            _resolved_family = name
            return name

    _resolved_family = "Tahoma"
    return _resolved_family


def get(size=10, weight="normal", root=None):
    family = resolve_family(root)
    return (family, size, weight)
