import argparse
from pathlib import Path

from . import __version__
from .core import detect_backends, export_selected_pages, parse_page_range
from .i18n import DEFAULT_LANGUAGE, SUPPORTED_LANGUAGES, get_translator, install_argparse_translations


def resolve_cli_language(argv=None):
    args = list(argv or [])
    for index, arg in enumerate(args):
        if arg == "--lang" and index + 1 < len(args):
            return args[index + 1]
        if arg.startswith("--lang="):
            return arg.split("=", 1)[1]
    return DEFAULT_LANGUAGE


def build_parser(lang=DEFAULT_LANGUAGE):
    install_argparse_translations(lang)
    translator = get_translator(lang)
    parser = argparse.ArgumentParser(
        prog="ppt2fig",
        description=translator("cli.description"),
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument(
        "--lang",
        choices=SUPPORTED_LANGUAGES,
        default=DEFAULT_LANGUAGE,
        help=translator("lang.help"),
    )
    parser.add_argument("pptx", nargs="?", help=translator("cli.arg.pptx"))
    parser.add_argument(
        "-p",
        "--pages",
        help=translator("cli.arg.pages"),
    )
    parser.add_argument(
        "-o",
        "--output",
        help=translator("cli.arg.output"),
    )
    parser.add_argument(
        "--office-bin",
        help=translator("cli.arg.office_bin"),
    )
    parser.add_argument(
        "--backend",
        choices=["auto", "libreoffice", "powerpoint", "wps"],
        default="auto",
        help=translator("cli.arg.backend"),
    )
    parser.add_argument(
        "--list-backends",
        action="store_true",
        help=translator("cli.arg.list_backends"),
    )
    parser.add_argument(
        "--powerpoint-intent",
        choices=["print", "screen"],
        default="print",
        help=translator("cli.arg.powerpoint_intent"),
    )
    parser.add_argument(
        "--bitmap-missing-fonts",
        action="store_true",
        help=translator("cli.arg.bitmap_missing_fonts"),
    )
    parser.add_argument("--no-crop", action="store_true", help=translator("cli.arg.no_crop"))
    parser.add_argument(
        "--percent-retain",
        type=float,
        default=0.0,
        help=translator("cli.arg.percent_retain"),
    )
    parser.add_argument(
        "--margin-size",
        type=float,
        default=0.0,
        help=translator("cli.arg.margin_size"),
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=191,
        help=translator("cli.arg.threshold"),
    )
    parser.add_argument(
        "--no-uniform",
        action="store_true",
        help=translator("cli.arg.no_uniform"),
    )
    parser.add_argument(
        "--no-same-size",
        action="store_true",
        help=translator("cli.arg.no_same_size"),
    )
    return parser


def default_output_path(pptx_path, pages):
    source = Path(pptx_path).resolve()
    page_label = "_".join(str(page) for page in pages)
    return source.with_name(f"{source.stem}.pages_{page_label}.pdf")


def main(argv=None):
    cli_lang = resolve_cli_language(argv)
    parser = build_parser(cli_lang)
    args = parser.parse_args(argv)
    translator = get_translator(args.lang)

    if args.list_backends:
        backends = detect_backends(explicit_path=args.office_bin, lang=args.lang)
        if not backends:
            print(translator("cli.list_backends.none"))
            return
        for backend in backends:
            status = (
                translator("cli.backend.supported")
                if backend.supported
                else translator("cli.backend.detected")
            )
            suffix = f" ({backend.detail})" if backend.detail else ""
            print(f"{backend.name}\t{status}\t{backend.path}{suffix}")
        return

    if not args.pptx:
        parser.error(translator("cli.error.missing_pptx"))
    if not args.pages:
        parser.error(translator("cli.error.missing_pages"))

    try:
        pages = parse_page_range(args.pages, lang=args.lang)
    except ValueError as exc:
        parser.error(str(exc))

    source = Path(args.pptx).resolve()
    if not source.exists():
        parser.error(translator("cli.error.input_not_found", path=source))
    if source.suffix.lower() not in {".pptx", ".ppt", ".odp"}:
        parser.error(translator("cli.error.invalid_extension"))

    output = Path(args.output).resolve() if args.output else default_output_path(source, pages)
    try:
        export_selected_pages(
            source,
            output,
            pages,
            lang=args.lang,
            backend=args.backend,
            office_bin=args.office_bin,
            powerpoint_intent=args.powerpoint_intent,
            bitmap_missing_fonts=args.bitmap_missing_fonts,
            no_crop=args.no_crop,
            percent_retain=args.percent_retain,
            margin_size=args.margin_size,
            use_uniform=not args.no_uniform,
            use_same_size=not args.no_same_size,
            threshold=args.threshold,
        )
    except Exception as exc:
        parser.exit(1, f"{parser.prog}: {translator('cli.error.prefix')}: {exc}\n")
    print(output)


if __name__ == "__main__":
    main()
