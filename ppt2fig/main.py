import platform
import sys

from .i18n import DEFAULT_LANGUAGE


def main():
    if len(sys.argv) > 1:
        from .cli import main as cli_main

        cli_main()
        return

    if platform.system() in {"Windows", "Darwin"}:
        from .gui import main as gui_main

        gui_main()
        return

    from .cli import build_parser

    build_parser(DEFAULT_LANGUAGE).print_help()


if __name__ == "__main__":
    main()
