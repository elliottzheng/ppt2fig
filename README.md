# PPT2Fig

PPT2Fig 用来把 `ppt`、`pptx`、`odp` 中指定页面导出成 PDF，并可选自动裁掉多余白边，适合论文、汇报和文档插图场景。

English README: [README.en.md](README.en.md)

## 项目用途

- 支持单页、多页、页码范围导出。
- 同时覆盖桌面交互导出和脚本/自动化导出。
- 保持后端标识稳定：`auto`、`libreoffice`、`wps`、`powerpoint`。
- 同时提供 CLI 和 GUI 入口。

## 后端支持

当前 `auto` 模式的优先级：

```text
LibreOffice > WPS > PowerPoint
```

支持情况：

- `LibreOffice`：Windows / Linux / macOS
- `WPS`：Windows
- `PowerPoint`：文件模式导出支持 Windows，快速活动演示文稿导出支持 Windows/macOS

## 使用方式

### 快速 GUI

适合当前已经打开 PowerPoint，希望“改一页，导一页”。

```bash
ppt2fig
```

### 文件模式 GUI

适合手动指定文件路径、页码、后端和输出路径。

```bash
ppt2fig-file-gui
```

### CLI

适合脚本、批处理、自动化和 AI 调用。

```bash
ppt2fig ./demo.pptx --pages 3
```

## 安装

### Windows 可执行文件

Releases:

https://github.com/elliottzheng/ppt2fig/releases

- `ppt2fig.exe`：快速 GUI
- `ppt2fig-file-gui.exe`：文件模式 GUI
- `ppt2fig-cli.exe`：CLI

### pip

```bash
pip install ppt2fig
```

安装后可直接运行：

```bash
ppt2fig
```

### OpenClaw / ClawHub 技能

```bash
clawhub install ppt2fig-export
```

安装后，AI 就可以直接调用 PPT2Fig，把指定演示文稿的指定页面导出成 PDF。

## CLI 快速上手

```bash
ppt2fig ./demo.pptx --pages 3
ppt2fig ./demo.pptx --pages 1,3,5-7 -o ./figure.pdf
ppt2fig ./demo.pptx --pages 2 --no-crop
ppt2fig ./demo.pptx --pages 2 --backend libreoffice
ppt2fig ./demo.pptx --pages 2 --backend powerpoint --powerpoint-intent print
ppt2fig --list-backends
ppt2fig --help --lang en
```

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

## 注意事项

- CLI 先导出整份 PDF，再抽取指定页。
- `--list-backends` 里的 `detected` 表示检测到候选程序，不一定代表当前平台已完整支持自动导出。
- PowerPoint 的 PDF 导出质量受其官方导出接口限制。

## 系统要求

- 快速 GUI：Windows/macOS，且活动演示文稿导出流程需要可用的 Microsoft PowerPoint
- 文件模式 GUI / CLI：Windows / Linux / macOS，跨平台场景推荐安装 LibreOffice
- Python：3.6+

## 维护者文档

- 编译指南：`docs/BUILD.md`
- ClawHub 发布指南：`docs/CLAWHUB_PUBLISH.md`

## 许可证

[MIT License](LICENSE)
