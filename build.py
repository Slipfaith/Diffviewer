from __future__ import annotations

import PyInstaller.__main__


def main() -> None:
    PyInstaller.__main__.run(["change_tracker.spec"])


if __name__ == "__main__":
    main()

