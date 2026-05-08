"""Output filename sanitization (ADR-0002).

Order:
1. Replace forbidden chars with `_`
2. Trim leading/trailing whitespace and trailing `.`
3. Reserved-name guard (CON, PRN, AUX, NUL, COM1-9, LPT1-9)
4. Empty basename → error
5. UTF-8 byte length > 255 → error
6. Caller MAY surface a warning when steps 1-3 changed the string.
"""

from __future__ import annotations

import re

from .errors import xtl_error

_FORBIDDEN = set('<>:"/\\|?*')
# ASCII control chars 0x00-0x1F
_FORBIDDEN |= {chr(i) for i in range(0x20)}

_RESERVED_BASENAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def sanitize_filename(rendered: str) -> tuple[str, str | None]:
    """Apply ADR-0002. Returns (sanitized, warning_or_None).

    Raises XtlError on empty/oversized.
    """
    original = rendered
    out = "".join("_" if ch in _FORBIDDEN else ch for ch in rendered)
    out = out.strip().rstrip(".")
    base, ext = _split_basename(out)
    if base.upper() in _RESERVED_BASENAMES:
        base = base + "_"
        out = base + ext
    if not out or not base:
        raise xtl_error(
            "xl3/filename/empty",
            f'Output filename "{original}" sanitized to an empty string and is invalid.',
        )
    if len(out.encode("utf-8")) > 255:
        raise xtl_error(
            "xl3/filename/too-long",
            f'Output filename "{out}" exceeds the 255-byte limit.',
        )
    if out != original:
        warning = f'Output filename "{original}" sanitized to "{out}"'
        return out, warning
    return out, None


_BASENAME_SPLIT = re.compile(r"^(.+?)(\.[^.]*)?$")


def _split_basename(name: str) -> tuple[str, str]:
    """Split into (basename, extension-with-dot). `.xlsx` stays in extension."""
    if "." not in name:
        return name, ""
    idx = name.rfind(".")
    return name[:idx], name[idx:]
