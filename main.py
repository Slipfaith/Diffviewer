from __future__ import annotations

import sys

APP_VERSION = "1.5"


def main() -> None:
    if len(sys.argv) <= 1:
        if sys.platform.startswith("win"):
            try:
                import ctypes
                console = ctypes.windll.kernel32.GetConsoleWindow()
                if console:
                    ctypes.windll.user32.ShowWindow(console, 0)
            except Exception:
                pass
        from ui.main_window import run_gui

        run_gui()
        return

    from cli import main as cli_main

    raise SystemExit(cli_main())


if __name__ == "__main__":
    main()
