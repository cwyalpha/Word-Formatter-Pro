# -*- coding: utf-8 -*-
"""Compatibility entry point for Word Formatter Pro v2.7.3."""

import sys


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == "--test":
        from wfp_tests import main as test_main

        return test_main(argv[1:])

    from wfp_gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
