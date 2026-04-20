# PPT2Fig

PPT2Fig exports selected slides or pages from `ppt`, `pptx`, and `odp` files to PDF, then optionally crops excess whitespace so the result is ready for papers, reports, and figure-heavy documents.

PPT2Fig 用来把 `ppt`、`pptx`、`odp` 中指定页面导出成 PDF，并可选自动裁掉多余白边，适合论文、汇报和文档插图场景。

## Overview / 项目用途

- Export a single page, multiple pages, or page ranges from presentation files.
- Support quick interactive export on desktop and scripted export in automation workflows.
- Keep backend identifiers stable: `auto`, `libreoffice`, `wps`, `powerpoint`.
- Provide both CLI and GUI entry points.
- 支持单页、多页、页码范围导出。
- 同时覆盖桌面交互导出和脚本/自动化导出。
- 保持后端标识稳定：`auto`、`libreoffice`、`wps`、`powerpoint`。
- 同时提供 CLI 和 GUI 入口。

## Backend Support / 后端支持

Auto-selection currently prefers:

```text
LibreOffice > WPS > PowerPoint
```

Supported platforms:

- `LibreOffice`: Windows / Linux / macOS
- `WPS`: Windows
- `PowerPoint`: Windows for file-mode export, Windows/macOS for quick active-presentation export

当前 `auto` 模式的优先级：

```text
LibreOffice > WPS > PowerPoint
```

支持情况：

- `LibreOffice`：Windows / Linux / macOS
- `WPS`：Windows
- `PowerPoint`：文件模式导出支持 Windows，快速活动演示文稿导出支持 Windows/macOS

## Modes / 使用方式

### Quick GUI / 快速 GUI

Best when you already have PowerPoint open and want to export the current slide quickly.

适合当前已经打开 PowerPoint，希望“改一页，导一页”。

```bash
ppt2fig
```

### File-Mode GUI / 文件模式 GUI

Best when you want to choose the source file, pages, backend, and output path explicitly.

适合手动指定文件路径、页码、后端和输出路径。

```bash
ppt2fig-file-gui
```

### CLI

Best for scripts, batch jobs, and AI/tool integrations.

适合脚本、批处理、自动化和 AI 调用。

```bash
ppt2fig ./demo.pptx --pages 3
```

## Installation / 安装

### Windows executables / Windows 可执行文件

Releases:

https://github.com/elliottzheng/ppt2fig/releases

- `ppt2fig.exe`: quick GUI
- `ppt2fig-file-gui.exe`: file-mode GUI
- `ppt2fig-cli.exe`: CLI
- `ppt2fig.exe`：快速 GUI
- `ppt2fig-file-gui.exe`：文件模式 GUI
- `ppt2fig-cli.exe`：CLI

### pip

```bash
pip install ppt2fig
```

After installation:

安装后可直接运行：

```bash
ppt2fig
```

### OpenClaw / ClawHub skill

```bash
clawhub install ppt2fig-export
```

After installation, AI tools can invoke PPT2Fig directly to export selected pages from a presentation to PDF.

安装后，AI 就可以直接调用 PPT2Fig，把指定演示文稿的指定页面导出成 PDF。

## CLI Quick Start / CLI 快速上手

```bash
ppt2fig ./demo.pptx --pages 3
ppt2fig ./demo.pptx --pages 1,3,5-7 -o ./figure.pdf
ppt2fig ./demo.pptx --pages 2 --no-crop
ppt2fig ./demo.pptx --pages 2 --backend libreoffice
ppt2fig ./demo.pptx --pages 2 --backend powerpoint --powerpoint-intent print
ppt2fig --list-backends
ppt2fig --help --lang en
```

Useful flags:

- `--pages`: required, supports `1,3,5-7`
- `--output`: output PDF path
- `--backend`: `auto` / `libreoffice` / `wps` / `powerpoint`
- `--office-bin`: manually specify a backend executable path
- `--no-crop`: skip whitespace cropping
- `--percent-retain`: retain part of the original margin
- `--margin-size`: add extra white margin after cropping
- `--threshold`: background detection threshold
- `--powerpoint-intent`: `print` or `screen`
- `--bitmap-missing-fonts`: rasterize text if fonts cannot be embedded
- `--lang`: `zh` or `en`, default `zh`

常用参数：

- `--pages`：必填，支持 `1,3,5-7`
- `--output`：输出 PDF 路径
- `--backend`：`auto` / `libreoffice` / `wps` / `powerpoint`
- `--office-bin`：手动指定后端程序路径
- `--no-crop`：不裁剪白边
- `--percent-retain`：保留部分原始边距
- `--margin-size`：裁剪后额外增加白边
- `--threshold`：背景检测阈值
- `--powerpoint-intent`：`print` 或 `screen`
- `--bitmap-missing-fonts`：字体无法嵌入时将文字转位图
- `--lang`：`zh` 或 `en`，默认 `zh`

## Notes / 注意事项

- The CLI exports the full PDF first and then extracts the selected pages.
- `detected` in `--list-backends` means a candidate program was found, not necessarily that full automatic export is supported on the current platform.
- PowerPoint export quality is limited by the official export interfaces it exposes.
- CLI 先导出整份 PDF，再抽取指定页。
- `--list-backends` 里的 `detected` 表示检测到候选程序，不一定代表当前平台已完整支持自动导出。
- PowerPoint 的 PDF 导出质量受其官方导出接口限制。

## Requirements / 系统要求

- Quick GUI: Windows/macOS with Microsoft PowerPoint available for the active-presentation workflow
- File-mode GUI / CLI: Windows / Linux / macOS, with LibreOffice recommended for cross-platform use
- Python: 3.6+
- 快速 GUI：Windows/macOS，且活动演示文稿导出流程需要可用的 Microsoft PowerPoint
- 文件模式 GUI / CLI：Windows / Linux / macOS，跨平台场景推荐安装 LibreOffice
- Python：3.6+

## Maintainer Docs / 维护者文档

- Build guide: `docs/BUILD.md`
- ClawHub publishing guide: `docs/CLAWHUB_PUBLISH.md`
- 编译指南：`docs/BUILD.md`
- ClawHub 发布指南：`docs/CLAWHUB_PUBLISH.md`

## License / 许可证

[MIT License](LICENSE)
