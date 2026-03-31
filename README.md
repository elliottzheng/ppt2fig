# PPT2Fig

PPT2Fig 用来把 PPT 页面导出成适合论文、汇报和文档插图使用的 PDF，并自动裁掉多余白边。

你可以把它当成：

- Windows 下的当前页一键导出工具
- 跨平台的 PPT 页面导出 CLI
- 可被 OpenClaw / ClawHub 调用的导出 skill

## 10 秒看懂

- 输入：`ppt` / `pptx` / `odp`
- 输出：裁好白边的 PDF
- 支持：单页、多页、页码范围
- 可用方式：GUI、CLI、ClawHub skill
- 后端：`LibreOffice`、`WPS`、`PowerPoint`

## 主要特性

- 适合论文作图、汇报插图、页面抽取
- 自动裁剪白边，减少后处理
- `auto` 模式自动选择可用后端
- Windows 下保留“当前 PowerPoint 当前页”快速导出
- 提供独立文件模式 GUI，适合重复导出
- 提供跨平台 CLI，适合自动化
- 可作为 OpenClaw / ClawHub skill 使用

## 我该用哪种方式

### 快速 GUI

适合 Windows + PowerPoint，想要“改一页，导一页”。

```bash
ppt2fig
```

### 文件模式 GUI

适合需要指定文件路径、页码、输出路径、后端的人。

```bash
ppt2fig-file-gui
```

### CLI / Skill

适合脚本、批处理、自动化、AI 调用。

```bash
ppt2fig ./demo.pptx --pages 3
```

## 安装

### Windows：直接下载 exe

最简单：

https://github.com/elliottzheng/ppt2fig/releases

- `ppt2fig.exe`：快速 GUI
- `ppt2fig-file-gui.exe`：文件模式 GUI
- `ppt2fig-cli.exe`：CLI

### Python：使用 pip

```bash
pip install ppt2fig
```

安装后可直接运行：

```bash
ppt2fig
```

### OpenClaw / ClawHub：安装 skill

```bash
clawhub install ppt2fig-export
```
安装后，AI 就可以直接调用 PPT2Fig，把指定演示文稿的指定页面导出成 PDF。需要搜索或更新时，可使用 `clawhub search ppt2fig` 和 `clawhub update ppt2fig-export`。

## CLI 快速上手

```bash
ppt2fig ./demo.pptx --pages 3
ppt2fig ./demo.pptx --pages 1,3,5-7 -o ./figure.pdf
ppt2fig ./demo.pptx --pages 2 --no-crop
ppt2fig ./demo.pptx --pages 2 --backend libreoffice
ppt2fig ./demo.pptx --pages 2 --backend powerpoint --powerpoint-intent print
ppt2fig --list-backends
```

查看版本可使用 `ppt2fig -v`。

## 后端说明

当前 `auto` 模式的优先级是：

```text
LibreOffice > WPS > PowerPoint
```

支持情况：

- `LibreOffice`：Windows / Linux / macOS
- `WPS`：Windows
- `PowerPoint`：Windows

如果你在 Linux/macOS 上使用 CLI，推荐安装 LibreOffice，并确保命令行里能找到 `soffice` 或 `libreoffice`。

## 常用参数

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

## 注意事项

- 快速 GUI 依赖当前打开的 PowerPoint
- 文件模式 GUI 和 CLI 更适合跨平台与重复导出
- CLI 通过先导出整份 PDF，再抽取指定页来工作
- `--list-backends` 中的 `detected` 表示检测到候选程序，不一定表示当前平台已完整支持自动导出
- PowerPoint 的 PDF 导出质量受其官方导出接口限制

## 系统要求

- 快速 GUI：Windows + Microsoft PowerPoint
- 文件模式 GUI / CLI：Windows / Linux / macOS，推荐 LibreOffice
- Python：3.8+

## 维护者文档

- 编译指南见 `docs/BUILD.md`
- ClawHub 发布指南见 `docs/CLAWHUB_PUBLISH.md`

## 许可证

[MIT License](LICENSE)
