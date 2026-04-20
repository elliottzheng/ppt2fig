import argparse
from dataclasses import dataclass


DEFAULT_LANGUAGE = "zh"
SUPPORTED_LANGUAGES = ("zh", "en")


MESSAGES = {
    "lang.label": {"zh": "语言", "en": "Language"},
    "lang.help": {"zh": "界面/命令行语言，默认 zh", "en": "UI/CLI language, default: zh"},
    "lang.option.zh": {"zh": "中文", "en": "Chinese"},
    "lang.option.en": {"zh": "英文", "en": "English"},
    "common.success": {"zh": "成功", "en": "Success"},
    "common.error": {"zh": "错误", "en": "Error"},
    "common.browse": {"zh": "浏览", "en": "Browse"},
    "common.refresh": {"zh": "刷新", "en": "Refresh"},
    "common.load": {"zh": "载入", "en": "Load"},
    "common.delete": {"zh": "删除", "en": "Delete"},
    "common.filetype.presentation": {"zh": "演示文稿", "en": "Presentation"},
    "common.filetype.all_files": {"zh": "所有文件", "en": "All files"},
    "common.filetype.pdf": {"zh": "PDF 文件", "en": "PDF file"},
    "cli.description": {
        "zh": "将指定 PPTX 的指定页导出为 PDF，并可选自动裁剪白边。",
        "en": "Export selected pages from a presentation to PDF, with optional automatic whitespace cropping.",
    },
    "cli.arg.pptx": {"zh": "输入的 PPTX 文件路径", "en": "Input presentation path"},
    "cli.arg.pages": {
        "zh": "要导出的页码，支持 1,3,5-7 这种格式，页码从 1 开始",
        "en": "Pages to export, such as 1,3,5-7. Page numbers start at 1.",
    },
    "cli.arg.output": {
        "zh": "输出 PDF 路径，默认在输入文件同目录下生成",
        "en": "Output PDF path. Defaults to the input file directory.",
    },
    "cli.arg.office_bin": {
        "zh": "手动指定导出后端可执行文件路径，例如 soffice 或 libreoffice",
        "en": "Manually specify the export backend executable, such as soffice or libreoffice.",
    },
    "cli.arg.backend": {
        "zh": "选择导出后端，默认 auto",
        "en": "Select the export backend. Default: auto.",
    },
    "cli.arg.list_backends": {
        "zh": "列出当前系统检测到的候选后端并退出",
        "en": "List detected backend candidates on this system and exit.",
    },
    "cli.arg.powerpoint_intent": {
        "zh": "PowerPoint 后端的导出意图，默认 print",
        "en": "PowerPoint export intent. Default: print.",
    },
    "cli.arg.bitmap_missing_fonts": {
        "zh": "PowerPoint 后端在字体无法嵌入时将文字位图化",
        "en": "Rasterize text when PowerPoint cannot embed fonts.",
    },
    "cli.arg.no_crop": {"zh": "导出后不裁剪白边", "en": "Do not crop whitespace after export."},
    "cli.arg.percent_retain": {
        "zh": "保留原始边距的百分比，默认 0",
        "en": "Percentage of the original margins to retain. Default: 0.",
    },
    "cli.arg.margin_size": {
        "zh": "裁剪后额外增加的白边，单位 bp，默认 0",
        "en": "Extra margin to add after cropping, in bp. Default: 0.",
    },
    "cli.arg.threshold": {
        "zh": "背景检测阈值，默认 191",
        "en": "Background detection threshold. Default: 191.",
    },
    "cli.arg.no_uniform": {"zh": "禁用统一裁剪", "en": "Disable uniform cropping."},
    "cli.arg.no_same_size": {"zh": "禁用统一页面大小", "en": "Disable same-size pages."},
    "cli.error.missing_pptx": {"zh": "缺少输入文件 pptx", "en": "missing input file: pptx"},
    "cli.error.missing_pages": {"zh": "缺少必填参数: -p/--pages", "en": "missing required argument: -p/--pages"},
    "cli.error.prefix": {"zh": "错误", "en": "error"},
    "cli.error.input_not_found": {"zh": "输入文件不存在: {path}", "en": "input file does not exist: {path}"},
    "cli.error.invalid_extension": {
        "zh": "输入文件必须是 .pptx、.ppt 或 .odp",
        "en": "input file must be .pptx, .ppt, or .odp",
    },
    "cli.list_backends.none": {"zh": "未检测到任何候选后端", "en": "No backend candidates detected."},
    "cli.backend.supported": {"zh": "已支持", "en": "supported"},
    "cli.backend.detected": {"zh": "已检测", "en": "detected"},
    "core.error.missing_pdfcropmargins": {
        "zh": "当前环境缺少 pdfCropMargins，无法执行 PDF 裁剪",
        "en": "pdfCropMargins is not installed in the current environment, so PDF cropping is unavailable.",
    },
    "core.error.page_empty": {"zh": "页码不能为空", "en": "page specification cannot be empty"},
    "core.error.page_must_start_1": {"zh": "页码必须从 1 开始", "en": "page numbers must start at 1"},
    "core.error.invalid_page_range": {"zh": "无效页码范围: {part}", "en": "invalid page range: {part}"},
    "core.error.invalid_page_token": {"zh": "无效页码值: {value}", "en": "invalid page value: {value}"},
    "core.error.no_valid_pages": {"zh": "没有解析出有效页码", "en": "no valid pages were parsed"},
    "core.error.missing_pypdf": {
        "zh": "当前环境缺少 pypdf，无法抽取指定页",
        "en": "pypdf is not installed in the current environment, so selected pages cannot be extracted.",
    },
    "core.error.page_out_of_range": {
        "zh": "请求的页码 {page_number} 超出范围，导出的 PDF 只有 {total_pages} 页",
        "en": "requested page {page_number} is out of range; the exported PDF only has {total_pages} pages",
    },
    "core.detail.missing_comtypes": {"zh": "当前环境缺少 comtypes", "en": "comtypes is not installed in the current environment"},
    "core.detail.com_unavailable": {"zh": "COM 不可用: {error}", "en": "COM is unavailable: {error}"},
    "core.detail.com_available": {"zh": "COM 可用: {value}", "en": "COM is available: {value}"},
    "core.detail.powerpoint_not_detected_mac": {
        "zh": "未检测到 Microsoft PowerPoint.app",
        "en": "Microsoft PowerPoint.app was not detected",
    },
    "core.detail.powerpoint_applescript": {"zh": "可通过 AppleScript 驱动", "en": "driven through AppleScript"},
    "core.detail.powerpoint_cli_not_supported": {
        "zh": "当前系统不支持 PowerPoint CLI 导出",
        "en": "PowerPoint CLI export is not supported on this system",
    },
    "core.detail.wps_binary_not_detected": {
        "zh": "未检测到 WPS 可执行文件",
        "en": "no WPS executable was detected",
    },
    "core.detail.wps_com_progid_not_found": {
        "zh": "未找到 WPS Presentation COM ProgID",
        "en": "no WPS Presentation COM ProgID was found",
    },
    "core.detail.wps_mac_not_implemented": {
        "zh": "已检测到 WPS.app，但当前还未实现 macOS 自动导出",
        "en": "WPS.app was detected, but automatic export on macOS is not implemented yet",
    },
    "core.detail.wps_app_not_detected": {"zh": "未检测到 WPS.app", "en": "WPS.app was not detected"},
    "core.detail.wps_platform_not_implemented": {
        "zh": "已检测到 WPS 可执行文件，但当前还未实现该平台自动导出",
        "en": "a WPS executable was detected, but automatic export on this platform is not implemented yet",
    },
    "core.detail.wps_candidate_detected": {"zh": "已检测到 WPS 候选程序", "en": "WPS candidate detected"},
    "core.detail.explicit_office_bin": {"zh": "通过 --office-bin 指定", "en": "specified via --office-bin"},
    "core.detail.powerpoint_candidate_detected": {
        "zh": "已检测到 PowerPoint 候选程序",
        "en": "PowerPoint candidate detected",
    },
    "core.error.backend_not_found": {
        "zh": "未找到可用的导出后端。请求后端: {backend}。已检测到: {detected_names}。当前已实现的自动导出后端: LibreOffice、PowerPoint(Windows/macOS)、WPS(Windows)。",
        "en": "No usable export backend was found. Requested backend: {backend}. Detected: {detected_names}. Automatic export is currently implemented for LibreOffice, PowerPoint (Windows/macOS), and WPS (Windows).",
    },
    "core.error.libreoffice_export_failed_default": {"zh": "LibreOffice 导出失败", "en": "LibreOffice export failed"},
    "core.error.libreoffice_pdf_missing": {
        "zh": "LibreOffice 没有生成预期的 PDF 文件",
        "en": "LibreOffice did not generate the expected PDF file",
    },
    "core.error.powerpoint_windows_only": {"zh": "PowerPoint 后端仅支持 Windows", "en": "the PowerPoint backend only supports Windows"},
    "core.error.missing_comtypes_powerpoint": {
        "zh": "当前环境缺少 comtypes，无法调用 PowerPoint",
        "en": "comtypes is not installed in the current environment, so PowerPoint cannot be used.",
    },
    "core.error.powerpoint_export_failed": {"zh": "PowerPoint 导出失败: {error}", "en": "PowerPoint export failed: {error}"},
    "core.error.powerpoint_pdf_missing": {
        "zh": "PowerPoint 没有生成预期的 PDF 文件",
        "en": "PowerPoint did not generate the expected PDF file",
    },
    "core.error.wps_windows_only": {"zh": "WPS 后端仅支持 Windows", "en": "the WPS backend only supports Windows"},
    "core.error.missing_comtypes_wps": {
        "zh": "当前环境缺少 comtypes，无法调用 WPS",
        "en": "comtypes is not installed in the current environment, so WPS cannot be used.",
    },
    "core.error.wps_com_missing": {
        "zh": "未找到可用的 WPS Presentation COM 接口",
        "en": "no usable WPS Presentation COM interface was found",
    },
    "core.error.wps_export_failed_both": {
        "zh": "WPS 导出失败。ExportAsFixedFormat 错误: {fixed_error}; SaveAs(PDF) 错误: {save_error}",
        "en": "WPS export failed. ExportAsFixedFormat error: {fixed_error}; SaveAs(PDF) error: {save_error}",
    },
    "core.error.wps_export_failed": {"zh": "WPS 导出失败: {error}", "en": "WPS export failed: {error}"},
    "core.error.wps_pdf_missing": {"zh": "WPS 没有生成预期的 PDF 文件", "en": "WPS did not generate the expected PDF file"},
    "core.error.backend_unsupported": {
        "zh": "暂不支持使用后端 {backend} 导出",
        "en": "export with backend {backend} is not supported yet",
    },
    "core.error.active_powerpoint_not_running": {"zh": "PowerPoint未启动", "en": "PowerPoint is not running"},
    "core.error.active_powerpoint_no_file": {"zh": "没有打开的PPT文件", "en": "no presentation is open in PowerPoint"},
    "core.error.conversion_failed": {"zh": "转换过程出错", "en": "the conversion process failed"},
    "core.error.active_presentation_missing": {
        "zh": "PowerPoint未启动或没有打开的文件",
        "en": "PowerPoint is not running or no presentation is open",
    },
    "core.error.active_presentation_info_failed": {"zh": "无法获取PPT文件信息", "en": "unable to get presentation file information"},
    "gui.crop.title": {"zh": "PDF裁剪参数设置", "en": "PDF Crop Settings"},
    "gui.crop.quick": {"zh": "快速设置", "en": "Quick Presets"},
    "gui.crop.detail": {"zh": "详细参数", "en": "Detailed Options"},
    "gui.crop.preset.tight": {"zh": "紧密裁剪", "en": "Tight Crop"},
    "gui.crop.preset.small_margin": {"zh": "小白边", "en": "Small Margin"},
    "gui.crop.preset.medium_margin": {"zh": "中白边", "en": "Medium Margin"},
    "gui.crop.preset.keep_original": {"zh": "保留原边距", "en": "Keep Original Margin"},
    "gui.crop.percent": {"zh": "保留原始边距(%):", "en": "Retain Original Margin (%):"},
    "gui.crop.margin": {"zh": "额外白边(bp):", "en": "Extra Margin (bp):"},
    "gui.crop.threshold": {"zh": "检测阈值:", "en": "Threshold:"},
    "gui.crop.no_crop": {"zh": "不裁剪", "en": "No Crop"},
    "gui.crop.uniform": {"zh": "统一裁剪", "en": "Uniform Crop"},
    "gui.crop.same_size": {"zh": "统一页面大小", "en": "Same Page Size"},
    "gui.quick.title": {"zh": "PPT转PDF工具 v{version}", "en": "PPT to PDF Tool v{version}"},
    "gui.quick.export_button": {"zh": "转PDF", "en": "Export PDF"},
    "gui.quick.show_advanced": {"zh": "▼ 显示高级设置", "en": "▼ Show Advanced"},
    "gui.quick.hide_advanced": {"zh": "▲ 隐藏高级设置", "en": "▲ Hide Advanced"},
    "gui.quick.success_message": {"zh": "PDF已导出至：\n{path}", "en": "PDF exported to:\n{path}"},
    "gui.quick.convert_failed": {"zh": "转换失败", "en": "Conversion failed"},
    "gui.quick.save_title": {"zh": "选择输出 PDF", "en": "Choose Output PDF"},
    "gui.file.title": {"zh": "PPT2Fig 文件模式 v{version}", "en": "PPT2Fig File Mode v{version}"},
    "gui.file.header_title": {"zh": "PPT2Fig 文件模式", "en": "PPT2Fig File Mode"},
    "gui.file.header_subtitle": {
        "zh": "面向重复导出场景：左侧保留最近任务，右侧编辑当前配置。",
        "en": "Built for repeated exports: recent tasks stay on the left, and the active configuration is edited on the right.",
    },
    "gui.file.history_title": {"zh": "最近任务", "en": "Recent Tasks"},
    "gui.file.history_subtitle": {
        "zh": "双击可直接重导，单击后可载入到右侧编辑。",
        "en": "Double-click to re-export immediately, or single-click to load the task into the editor.",
    },
    "gui.file.history_untitled": {"zh": "未命名", "en": "Untitled"},
    "gui.file.history_label": {
        "zh": "{prefix}{name} | 页码 {pages} | {backend}",
        "en": "{prefix}{name} | Pages {pages} | {backend}",
    },
    "gui.file.history_rerun": {"zh": "重导", "en": "Re-export"},
    "gui.file.current_title": {"zh": "当前配置", "en": "Current Configuration"},
    "gui.file.choose_file": {"zh": "选择文件", "en": "Choose File"},
    "gui.file.output_same_dir": {"zh": "输出同目录", "en": "Output Beside Source"},
    "gui.file.refresh_backends": {"zh": "刷新后端", "en": "Refresh Backends"},
    "gui.file.auto_output_name": {"zh": "自动生成输出文件名", "en": "Auto-generate Output Name"},
    "gui.file.section_basic": {"zh": "基础信息", "en": "Basic Information"},
    "gui.file.field_source": {"zh": "输入文件", "en": "Input File"},
    "gui.file.field_pages": {"zh": "页码", "en": "Pages"},
    "gui.file.pages_hint": {"zh": "支持 1,3,5-7", "en": "Supports 1,3,5-7"},
    "gui.file.first_page": {"zh": "第一页", "en": "First Page"},
    "gui.file.field_output": {"zh": "输出 PDF", "en": "Output PDF"},
    "gui.file.section_backend": {"zh": "导出后端", "en": "Export Backend"},
    "gui.file.field_backend": {"zh": "后端", "en": "Backend"},
    "gui.file.field_program_path": {"zh": "程序路径", "en": "Program Path"},
    "gui.file.field_detection_status": {"zh": "检测状态", "en": "Detection Status"},
    "gui.file.section_powerpoint": {"zh": "PowerPoint 特定选项", "en": "PowerPoint Options"},
    "gui.file.field_export_intent": {"zh": "导出意图", "en": "Export Intent"},
    "gui.file.bitmap_missing_fonts": {"zh": "缺字库时将文字位图化", "en": "Rasterize text when fonts are missing"},
    "gui.file.section_crop": {"zh": "裁剪与页面设置", "en": "Cropping and Page Settings"},
    "gui.file.export_button": {"zh": "导出当前配置", "en": "Export Current Configuration"},
    "gui.file.status.ready": {
        "zh": "选择文件后即可导出，常用配置会自动进入左侧历史记录。",
        "en": "Choose a file to start exporting. Frequent configurations are added to the history list automatically.",
    },
    "gui.file.status.loaded_history": {
        "zh": "已载入历史任务，可直接修改后重新导出。",
        "en": "The history task has been loaded. You can edit it and export again.",
    },
    "gui.file.status.deleted_history": {"zh": "已删除一条历史任务。", "en": "One history task was deleted."},
    "gui.file.status.selected_file": {"zh": "已选择输入文件。", "en": "Input file selected."},
    "gui.file.status.refreshed_backends": {"zh": "已刷新后端检测结果。", "en": "Backend detection results refreshed."},
    "gui.file.status.export_failed": {"zh": "导出失败: {error}", "en": "Export failed: {error}"},
    "gui.file.status.export_done": {"zh": "导出完成: {path}", "en": "Export completed: {path}"},
    "gui.file.status.exporting": {"zh": "正在导出，请稍候...", "en": "Exporting. Please wait..."},
    "gui.file.backend_none_detected": {"zh": "未检测到可用后端", "en": "No usable backends detected"},
    "gui.file.dialog.select_input": {"zh": "选择输入文件", "en": "Choose Input File"},
    "gui.file.dialog.select_output": {"zh": "选择输出 PDF", "en": "Choose Output PDF"},
    "gui.file.dialog.select_backend_bin": {"zh": "选择后端可执行文件", "en": "Choose Backend Executable"},
    "gui.file.error.select_input": {"zh": "请选择输入文件", "en": "please choose an input file"},
    "gui.file.error.input_not_found": {"zh": "输入文件不存在", "en": "input file does not exist"},
}


ARGPARSE_TRANSLATIONS = {
    "zh": {
        "positional arguments": "位置参数",
        "options": "可选参数",
        "show this help message and exit": "显示此帮助信息并退出",
        "show program's version number and exit": "显示程序版本号并退出",
        "usage: ": "用法: ",
        "error: ": "错误: ",
        "argument %(argument_name)s: %(message)s": "参数 %(argument_name)s: %(message)s",
        "the following arguments are required: %s": "缺少以下必填参数: %s",
        "expected one argument": "需要一个参数值",
        "invalid %(type)s value: %(value)r": "无效%(type)s值: %(value)r",
        "invalid choice: %(value)r (choose from %(choices)s)": "无效选项: %(value)r（可选值: %(choices)s）",
        "%(prog)s: error: %(message)s\n": "%(prog)s: 错误: %(message)s\n",
    },
    "en": {},
}


@dataclass(frozen=True)
class Translator:
    lang: str = DEFAULT_LANGUAGE

    def __call__(self, key, **kwargs):
        message = MESSAGES[key][self.lang]
        return message.format(**kwargs)

    tr = __call__


def normalize_language(lang):
    return lang if lang in SUPPORTED_LANGUAGES else DEFAULT_LANGUAGE


def get_translator(lang=None):
    return Translator(normalize_language(lang))


def tr(key, *, lang=None, **kwargs):
    return get_translator(lang)(key, **kwargs)


def install_argparse_translations(lang):
    translations = ARGPARSE_TRANSLATIONS.get(normalize_language(lang), {})

    def gettext(message):
        return translations.get(message, message)

    def ngettext(singular, plural, count):
        message = singular if count == 1 else plural
        return gettext(message)

    argparse._ = gettext
    argparse.ngettext = ngettext
