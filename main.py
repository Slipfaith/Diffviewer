from __future__ import annotations

import sys


def main() -> None:
    if len(sys.argv) > 1:
        from cli import main as cli_main

        raise SystemExit(cli_main())

    from ui.main_window import run_gui

    run_gui()


if __name__ == "__main__":
    main()
