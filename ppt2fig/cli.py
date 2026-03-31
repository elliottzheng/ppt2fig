import argparse
from pathlib import Path

from . import __version__
from .core import detect_backends, export_selected_pages, parse_page_range


def build_parser():
    parser = argparse.ArgumentParser(
        prog="ppt2fig",
        description="将指定 PPTX 的指定页导出为 PDF，并可选自动裁剪白边。",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    parser.add_argument("pptx", nargs="?", help="输入的 PPTX 文件路径")
    parser.add_argument(
        "-p",
        "--pages",
        help="要导出的页码，支持 1,3,5-7 这种格式，页码从 1 开始",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="输出 PDF 路径，默认在输入文件同目录下生成",
    )
    parser.add_argument(
        "--office-bin",
        help="手动指定导出后端可执行文件路径，例如 soffice 或 libreoffice",
    )
    parser.add_argument(
        "--backend",
        choices=["auto", "libreoffice", "powerpoint", "wps"],
        default="auto",
        help="选择导出后端，默认 auto",
    )
    parser.add_argument(
        "--list-backends",
        action="store_true",
        help="列出当前系统检测到的候选后端并退出",
    )
    parser.add_argument(
        "--powerpoint-intent",
        choices=["print", "screen"],
        default="print",
        help="PowerPoint 后端的导出意图，默认 print",
    )
    parser.add_argument(
        "--bitmap-missing-fonts",
        action="store_true",
        help="PowerPoint 后端在字体无法嵌入时将文字位图化",
    )
    parser.add_argument("--no-crop", action="store_true", help="导出后不裁剪白边")
    parser.add_argument(
        "--percent-retain",
        type=float,
        default=0.0,
        help="保留原始边距的百分比，默认 0",
    )
    parser.add_argument(
        "--margin-size",
        type=float,
        default=0.0,
        help="裁剪后额外增加的白边，单位 bp，默认 0",
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=191,
        help="背景检测阈值，默认 191",
    )
    parser.add_argument(
        "--no-uniform",
        action="store_true",
        help="禁用统一裁剪",
    )
    parser.add_argument(
        "--no-same-size",
        action="store_true",
        help="禁用统一页面大小",
    )
    return parser


def default_output_path(pptx_path, pages):
    source = Path(pptx_path).resolve()
    page_label = "_".join(str(page) for page in pages)
    return source.with_name(f"{source.stem}.pages_{page_label}.pdf")


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.list_backends:
        backends = detect_backends(explicit_path=args.office_bin)
        if not backends:
            print("未检测到任何候选后端")
            return
        for backend in backends:
            status = "supported" if backend.supported else "detected"
            suffix = f" ({backend.detail})" if backend.detail else ""
            print(f"{backend.name}\t{status}\t{backend.path}{suffix}")
        return

    if not args.pptx:
        parser.error("缺少输入文件 pptx")
    if not args.pages:
        parser.error("missing required argument: -p/--pages")

    pages = parse_page_range(args.pages)
    source = Path(args.pptx).resolve()
    if not source.exists():
        parser.error(f"输入文件不存在: {source}")
    if source.suffix.lower() not in {".pptx", ".ppt", ".odp"}:
        parser.error("输入文件必须是 .pptx、.ppt 或 .odp")

    output = Path(args.output).resolve() if args.output else default_output_path(source, pages)
    export_selected_pages(
        source,
        output,
        pages,
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
    print(output)


if __name__ == "__main__":
    main()
