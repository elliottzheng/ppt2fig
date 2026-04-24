# PPT2Fig

PPT2Fig exports selected slides or pages from `ppt`, `pptx`, and `odp` files to PDF, then optionally crops excess whitespace so the result is ready for papers, reports, and figure-heavy documents.

## Overview

- Export a single page, multiple pages, or page ranges from presentation files.
- Support quick interactive export on desktop and scripted export in automation workflows.
- Keep backend identifiers stable: `auto`, `libreoffice`, `wps`, `powerpoint`.
- Provide both CLI and GUI entry points.

## Backend Support

Auto-selection currently prefers:

```text
LibreOffice > WPS > PowerPoint
```

Supported platforms:

- `LibreOffice`: Windows / Linux / macOS
- `WPS`: Windows
- `PowerPoint`: Windows for file-mode export, Windows/macOS for quick active-presentation export

## Modes

### Quick GUI

Best when you already have PowerPoint open and want to export the current slide quickly.

```bash
ppt2fig
```

### File-Mode GUI

Best when you want to choose the source file, pages, backend, and output path explicitly.

```bash
ppt2fig-file-gui
```

### CLI

Best for scripts, batch jobs, and AI/tool integrations.

```bash
ppt2fig ./demo.pptx --pages 3
```

## Installation

### Windows executables

Releases:

https://github.com/elliottzheng/ppt2fig/releases

- `ppt2fig.exe`: quick GUI
- `ppt2fig-file-gui.exe`: file-mode GUI
- `ppt2fig-cli.exe`: CLI

### pip

```bash
pip install ppt2fig
```

After installation:

```bash
ppt2fig
```

### OpenClaw / ClawHub skill

```bash
clawhub install ppt2fig-export
```

After installation, AI tools can invoke PPT2Fig directly to export selected pages from a presentation to PDF.

## CLI Quick Start

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

## Notes

- The CLI exports the full PDF first and then extracts the selected pages.
- `detected` in `--list-backends` means a candidate program was found, not necessarily that full automatic export is supported on the current platform.
- PowerPoint export quality is limited by the official export interfaces it exposes.

## Requirements

- Quick GUI: Windows/macOS with Microsoft PowerPoint available for the active-presentation workflow
- File-mode GUI / CLI: Windows / Linux / macOS, with LibreOffice recommended for cross-platform use
- Python: 3.6+

## Maintainer Docs

- Build guide: `docs/BUILD.md`
- ClawHub publishing guide: `docs/CLAWHUB_PUBLISH.md`

## License

[MIT License](LICENSE)
