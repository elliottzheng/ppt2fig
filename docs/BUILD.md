# Build Guide

This document is for repository maintainers who need to build the Windows executables.

## Requirements

- Windows
- A Python environment with project dependencies installed
- `PyInstaller`
- Optional: `upx.exe` in the repository root

Recommended environment:

```bash
conda activate ppt2fig
```

## Build Commands

Build the three Windows executables with the PyInstaller spec files in `packaging/`:

```cmd
python -m PyInstaller packaging\ppt2fig.spec --noconfirm --distpath dist --workpath build\ppt2fig --upx-dir .
python -m PyInstaller packaging\ppt2fig-file-gui.spec --noconfirm --distpath dist --workpath build\ppt2fig-file-gui --upx-dir .
python -m PyInstaller packaging\ppt2fig-cli.spec --noconfirm --distpath dist --workpath build\ppt2fig-cli --upx-dir .
```

If `PyInstaller` is only available in the `ppt2fig` conda environment, use:

```cmd
C:\Users\18350\.conda\envs\ppt2fig\python.exe -m PyInstaller packaging\ppt2fig.spec --noconfirm --distpath dist --workpath build\ppt2fig --upx-dir .
C:\Users\18350\.conda\envs\ppt2fig\python.exe -m PyInstaller packaging\ppt2fig-file-gui.spec --noconfirm --distpath dist --workpath build\ppt2fig-file-gui --upx-dir .
C:\Users\18350\.conda\envs\ppt2fig\python.exe -m PyInstaller packaging\ppt2fig-cli.spec --noconfirm --distpath dist --workpath build\ppt2fig-cli --upx-dir .
```

## Outputs

The executables are written to `dist/`:

- `ppt2fig.exe`
- `ppt2fig-file-gui.exe`
- `ppt2fig-cli.exe`

## Notes

- `ppt2fig-cli.exe` is the file used by the ClawHub skill download metadata.
- The GUI executables are Windows-only.
- UPX is optional. If `upx.exe` is present in the repository root, PyInstaller will use it automatically because the spec files enable `upx=True`.
