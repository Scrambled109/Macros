from __future__ import annotations

import sys

from epl_converter.cli import run_cli


def main() -> int:
    if len(sys.argv) > 1:
        return run_cli(sys.argv[1:])
    from epl_converter.ui import run_ui

    run_ui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
