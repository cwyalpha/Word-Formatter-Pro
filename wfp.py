# -*- coding: utf-8 -*-
"""Compatibility entry point for Word Formatter Pro."""

import sys

from wfp_version import __version__


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "--test":
        from wfp_tests import main as test_main

        # Windowed Windows executables have no console streams. Keep --test
        # usable for release verification, with a readable result on disk.
        if sys.stdout is None or sys.stderr is None:
            import os
            import tempfile
            log_path = os.path.join(tempfile.gettempdir(), "Word-Formatter-Pro-tests.log")
            original_streams = sys.stdout, sys.stderr
            with open(log_path, "w", encoding="utf-8") as log:
                try:
                    sys.stdout = sys.stderr = log
                    return test_main(argv[1:])
                finally:
                    sys.stdout, sys.stderr = original_streams
        return test_main(argv[1:])
    if argv and argv[0] in ("--version", "-V"):
        print(__version__)
        return 0

    from wfp_gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
