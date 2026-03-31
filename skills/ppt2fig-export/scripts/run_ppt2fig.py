import shutil
import subprocess
import sys
from pathlib import Path


def find_program():
    candidates = [
        str(Path.home() / ".openclaw" / "tools" / "ppt2fig-export" / "ppt2fig-cli.exe"),
        str(Path.home() / ".openclaw" / "tools" / "ppt2fig-export" / "ppt2fig-cli-upx.exe"),
        shutil.which("ppt2fig"),
        shutil.which("ppt2fig-cli"),
        str(Path.cwd() / "dist" / "ppt2fig-cli-upx.exe"),
        str(Path.cwd() / "dist" / "ppt2fig-cli.exe"),
        str(Path(__file__).resolve().parents[3] / "dist" / "ppt2fig-cli-upx.exe"),
        str(Path(__file__).resolve().parents[3] / "dist" / "ppt2fig-cli.exe"),
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return [candidate]
    return [sys.executable, "-m", "ppt2fig"]


def main():
    command = find_program() + sys.argv[1:]
    completed = subprocess.run(command)
    raise SystemExit(completed.returncode)


if __name__ == "__main__":
    main()
