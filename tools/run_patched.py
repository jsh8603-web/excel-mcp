#!/usr/bin/env python3
"""Start one of the bundled MCP servers with Excel-compatible name validation.

Excel accepts any Unicode letter in a defined name or table name (Korean included), but both
bundled servers validate with ASCII-only patterns, so `세율` or `판관비율` is refused. This wrapper
loosens exactly those patterns in memory, then hands over to the server's own entry point.
Nothing in site-packages is modified, so a reinstall cannot undo or half-apply it.

  run_patched.py excel-mcp-server stdio --allow-dir DIR
  run_patched.py xlide-mcp --root DIR
"""
from __future__ import annotations

import re
import sys
from importlib import metadata

# first character: a letter, underscore or backslash; then letters, digits, underscore, period, backslash
UNICODE_NAME = re.compile(r"^(?:[^\W\d]|\\)[\w.\\]*$")
UNICODE_TABLE_NAME = re.compile(r"[^\W\d][\w.]{0,254}")


def patch() -> list[str]:
    done = []
    try:
        from pyofficeeditor.excel import _names

        _names._VALID_NAME = UNICODE_NAME
        done.append("pyofficeeditor._names._VALID_NAME")
    except Exception as exc:  # noqa: BLE001
        print(f"[xlkit] name patch skipped (pyofficeeditor): {exc}", file=sys.stderr)
    try:
        from excel_mcp.operations import tables

        tables._TABLE_NAME = UNICODE_TABLE_NAME
        done.append("excel_mcp.operations.tables._TABLE_NAME")
    except Exception as exc:  # noqa: BLE001
        print(f"[xlkit] name patch skipped (excel_mcp): {exc}", file=sys.stderr)
    return done


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    script = sys.argv[1]
    patch()
    entry = next((e for e in metadata.entry_points(group="console_scripts") if e.name == script), None)
    if entry is None:
        print(f"[xlkit] no console script named {script!r}", file=sys.stderr)
        return 3
    sys.argv = [script, *sys.argv[2:]]
    result = entry.load()()
    return result if isinstance(result, int) else 0


if __name__ == "__main__":
    sys.exit(main())
