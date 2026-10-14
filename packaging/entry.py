"""PyInstaller entry point for BinderBuilder.exe.

    BinderBuilder.exe                      the GUI
    BinderBuilder.exe --cli sync <folder>  the command line (diagnostics, scripting); output goes to
                                           %APPDATA%\BinderBuilder\cli.log when there is no console
"""

import os
import sys
from pathlib import Path


def _cli(argv: list[str]) -> int:
    from binder_builder.cli import main

    if sys.stdout is None:  # windowed build: no console
        log = Path(os.environ.get("APPDATA") or Path.home()) / "BinderBuilder" / "cli.log"
        log.parent.mkdir(parents=True, exist_ok=True)
        sys.stdout = sys.stderr = open(log, "w", encoding="utf-8")
    return main(argv)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        raise SystemExit(_cli(sys.argv[2:]))
    from binder_builder.gui.app import main

    main()
