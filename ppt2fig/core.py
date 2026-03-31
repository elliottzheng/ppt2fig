import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from dataclasses import dataclass


@dataclass(frozen=True)
class BackendInfo:
    name: str
    executable: str
    path: str
    supported: bool
    detail: str = ""


SYSTEM_NAME = platform.system()


def build_crop_args(
    input_pdf,
    output_pdf,
    *,
    percent_retain=0.0,
    margin_size=0.0,
    use_uniform=True,
    use_same_size=True,
    threshold=191,
):
    crop_args = ["-p", str(percent_retain)]
    if margin_size > 0:
        crop_args.extend(["-a", str(-margin_size)])
    if use_uniform:
        crop_args.append("-u")
    if use_same_size:
        crop_args.append("-s")
    if threshold != 191:
        crop_args.extend(["-t", str(threshold)])
    crop_args.extend([str(input_pdf), "-o", str(output_pdf)])
    return crop_args


def crop_pdf_file(
    pdf_file,
    *,
    no_crop=False,
    percent_retain=0.0,
    margin_size=0.0,
    use_uniform=True,
    use_same_size=True,
    threshold=191,
):
    if no_crop:
        return

    try:
        from pdfCropMargins import crop
    except ImportError as exc:
        raise RuntimeError("当前环境缺少 pdfCropMargins，无法执行 PDF 裁剪") from exc

    pdf_path = Path(pdf_file)
    tmp_output = pdf_path.with_suffix(pdf_path.suffix + ".crop")
    crop(
        build_crop_args(
            pdf_path,
            tmp_output,
            percent_retain=percent_retain,
            margin_size=margin_size,
            use_uniform=use_uniform,
            use_same_size=use_same_size,
            threshold=threshold,
        )
    )
    shutil.move(str(tmp_output), str(pdf_path))


def parse_page_range(page_spec):
    if not page_spec:
        raise ValueError("页码不能为空")

    pages = []
    for raw_part in page_spec.split(","):
        part = raw_part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start = int(start_text)
            end = int(end_text)
            if start <= 0 or end <= 0:
                raise ValueError("页码必须从 1 开始")
            if end < start:
                raise ValueError(f"无效页码范围: {part}")
            pages.extend(range(start, end + 1))
        else:
            page = int(part)
            if page <= 0:
                raise ValueError("页码必须从 1 开始")
            pages.append(page)

    if not pages:
        raise ValueError("没有解析出有效页码")

    deduplicated = []
    seen = set()
    for page in pages:
        if page not in seen:
            deduplicated.append(page)
            seen.add(page)
    return deduplicated


def extract_pdf_pages(input_pdf, output_pdf, pages):
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError as exc:
        raise RuntimeError("当前环境缺少 pypdf，无法抽取指定页") from exc

    reader = PdfReader(str(input_pdf))
    writer = PdfWriter()

    total_pages = len(reader.pages)
    for page_number in pages:
        if page_number > total_pages:
            raise ValueError(
                f"请求的页码 {page_number} 超出范围，导出的 PDF 只有 {total_pages} 页"
            )
        writer.add_page(reader.pages[page_number - 1])

    with open(output_pdf, "wb") as fh:
        writer.write(fh)


def _iter_backend_candidates():
    if SYSTEM_NAME == "Windows":
        program_files = [
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
            os.environ.get("LOCALAPPDATA"),
        ]
        for base in filter(None, program_files):
            yield ("libreoffice", os.path.join(base, "LibreOffice", "program", "soffice.exe"))
            yield ("wps", os.path.join(base, "Kingsoft", "WPS Office", "ksolaunch.exe"))
            yield ("wps", os.path.join(base, "Kingsoft", "WPS Office", "office6", "wps.exe"))
            yield ("wps", os.path.join(base, "Kingsoft", "WPS Office", "office6", "wpp.exe"))
            yield ("wps", os.path.join(base, "Kingsoft", "WPS Office", "12.1.0.20305", "office6", "wps.exe"))
            yield ("wps", os.path.join(base, "Kingsoft", "WPS Office", "12.1.0.20305", "office6", "wpp.exe"))
    elif SYSTEM_NAME == "Darwin":
        yield ("libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice")
        yield ("powerpoint", "/Applications/Microsoft PowerPoint.app")
        yield ("wps", "/Applications/WPS Office.app")
        yield ("wps", "/Applications/WPS Office/WPS Office.app")
    else:
        yield ("libreoffice", "/usr/bin/soffice")
        yield ("libreoffice", "/usr/local/bin/soffice")
        yield ("libreoffice", "/snap/bin/libreoffice")
        yield ("wps", "/usr/bin/wps")
        yield ("wps", "/usr/bin/wpp")
        yield ("wps", "/opt/kingsoft/wps-office/office6/wps")
        yield ("wps", "/opt/kingsoft/wps-office/office6/wpp")

    yield ("libreoffice", "soffice")
    yield ("libreoffice", "libreoffice")
    yield ("wps", "wps")
    yield ("wps", "wpp")
    yield ("wps", "ksolaunch")
    if SYSTEM_NAME == "Darwin":
        yield ("powerpoint", "Microsoft PowerPoint")


def _find_powerpoint_executable():
    if SYSTEM_NAME == "Windows":
        try:
            import winreg

            subkeys = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
            ]
            for subkey in subkeys:
                for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                    try:
                        with winreg.OpenKey(hive, subkey) as key:
                            value, _ = winreg.QueryValueEx(key, None)
                            if value and os.path.exists(value):
                                return value
                    except OSError:
                        continue
        except ImportError:
            pass

        common_roots = [
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
        ]
        patterns = [
            ("Microsoft Office", "root", "Office16", "POWERPNT.EXE"),
            ("Microsoft Office", "root", "Office15", "POWERPNT.EXE"),
            ("Microsoft Office", "Office16", "POWERPNT.EXE"),
            ("Microsoft Office", "Office15", "POWERPNT.EXE"),
        ]
        for root in filter(None, common_roots):
            for pattern in patterns:
                candidate = os.path.join(root, *pattern)
                if os.path.exists(candidate):
                    return candidate
        return shutil.which("POWERPNT.EXE")

    if SYSTEM_NAME == "Darwin":
        candidates = [
            "/Applications/Microsoft PowerPoint.app",
            str(Path.home() / "Applications" / "Microsoft PowerPoint.app"),
        ]
        for candidate in candidates:
            if os.path.exists(candidate):
                return candidate
    return None


def _find_windows_app_path(*exe_names):
    if SYSTEM_NAME != "Windows":
        return None

    try:
        import winreg
    except ImportError:
        return None

    for exe_name in exe_names:
        subkeys = [
            fr"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}",
            fr"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}",
        ]
        for subkey in subkeys:
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                try:
                    with winreg.OpenKey(hive, subkey) as key:
                        value, _ = winreg.QueryValueEx(key, None)
                        if value and os.path.exists(value):
                            return value
                except OSError:
                    continue
    return None


def _can_use_powerpoint_com():
    if SYSTEM_NAME == "Windows":
        try:
            import comtypes.client  # noqa: F401
        except ImportError:
            return False, "当前环境缺少 comtypes"

        try:
            app = comtypes.client.CreateObject("Powerpoint.Application")
            app.Quit()
            return True, ""
        except Exception as exc:
            return False, f"COM 不可用: {exc}"

    if SYSTEM_NAME == "Darwin":
        app_path = _find_powerpoint_executable()
        if not app_path:
            return False, "未检测到 Microsoft PowerPoint.app"
        return True, "可通过 AppleScript 驱动"

    return False, "当前系统不支持 PowerPoint CLI 导出"


def _find_wps_com_progid():
    if SYSTEM_NAME != "Windows":
        return None

    try:
        import comtypes.client
    except ImportError:
        return None

    for progid in ("KWPP.Application", "WPP.Application", "kwpp.Application", "wpp.Application"):
        try:
            app = comtypes.client.CreateObject(progid)
            try:
                app.Quit()
            except Exception:
                pass
            return progid
        except Exception:
            continue
    return None


def _can_use_wps_com(has_wps_binary=False):
    if SYSTEM_NAME == "Windows":
        try:
            import comtypes.client  # noqa: F401
        except ImportError:
            return False, "当前环境缺少 comtypes"

        if not has_wps_binary:
            return False, "未检测到 WPS 可执行文件"

        progid = _find_wps_com_progid()
        if progid:
            return True, f"COM 可用: {progid}"
        return False, "未找到 WPS Presentation COM ProgID"

    if SYSTEM_NAME == "Darwin":
        candidates = [
            "/Applications/WPS Office.app",
            "/Applications/WPS Office/WPS Office.app",
            str(Path.home() / "Applications" / "WPS Office.app"),
        ]
        for candidate in candidates:
            if os.path.exists(candidate):
                return False, "已检测到 WPS.app，但当前还未实现 macOS 自动导出"
        return False, "未检测到 WPS.app"

    if has_wps_binary:
        return False, "已检测到 WPS 可执行文件，但当前还未实现该平台自动导出"
    return False, "未检测到 WPS 可执行文件"


def detect_backends(explicit_path=None):
    backends = []
    seen_paths = set()
    has_wps_binary = False

    candidates = []
    if explicit_path:
        explicit_name = "custom"
        lower_path = explicit_path.lower()
        if "soffice" in lower_path or "libreoffice" in lower_path:
            explicit_name = "libreoffice"
        elif "powerpnt" in lower_path or "powerpoint" in lower_path or lower_path.endswith(".app"):
            explicit_name = "powerpoint"
        elif "\\wpp.exe" in lower_path or "\\wps.exe" in lower_path or "ksolaunch" in lower_path:
            explicit_name = "wps"
        candidates.append((explicit_name, explicit_path))
    if SYSTEM_NAME == "Windows":
        for app_path in filter(
            None,
            (
                _find_windows_app_path("wpp.exe"),
                _find_windows_app_path("wps.exe"),
                _find_windows_app_path("ksolaunch.exe"),
            ),
        ):
            candidates.append(("wps", app_path))
    candidates.extend(_iter_backend_candidates())

    for backend_name, candidate in candidates:
        resolved = shutil.which(candidate) if not os.path.isabs(candidate) else candidate
        if not resolved or not os.path.exists(resolved):
            continue
        normalized = os.path.normcase(os.path.abspath(resolved))
        if normalized in seen_paths:
            continue
        seen_paths.add(normalized)
        if backend_name == "wps":
            has_wps_binary = True

        supported = backend_name in {"libreoffice", "custom"}
        detail = ""
        if backend_name == "wps":
            supported = False
            detail = "已检测到 WPS 候选程序"
        elif backend_name == "custom":
            detail = "通过 --office-bin 指定"
        elif backend_name == "powerpoint":
            detail = "已检测到 PowerPoint 候选程序"
        elif explicit_path and os.path.abspath(resolved) == os.path.abspath(explicit_path):
            detail = "通过 --office-bin 指定"

        backends.append(
            BackendInfo(
                name=backend_name,
                executable=os.path.basename(resolved),
                path=os.path.abspath(resolved),
                supported=supported,
                detail=detail,
            )
        )

    if has_wps_binary:
        supported, detail = _can_use_wps_com(has_wps_binary=True)
        for index, item in enumerate(backends):
            if item.name == "wps":
                backends[index] = BackendInfo(
                    name=item.name,
                    executable=item.executable,
                    path=item.path,
                    supported=supported,
                    detail=detail,
                )

    powerpoint_executable = _find_powerpoint_executable()
    if powerpoint_executable:
        normalized = os.path.normcase(os.path.abspath(powerpoint_executable))
        if normalized not in seen_paths:
            supported, detail = _can_use_powerpoint_com()
            backends.append(
                BackendInfo(
                    name="powerpoint",
                    executable=os.path.basename(powerpoint_executable),
                    path=os.path.abspath(powerpoint_executable),
                    supported=supported,
                    detail=detail,
                )
            )
            seen_paths.add(normalized)

    return backends


def find_backend(explicit_path=None, backend="auto"):
    detected = detect_backends(explicit_path=explicit_path)
    if backend == "auto":
        preferred_order = ["libreoffice", "wps", "powerpoint", "custom"]
    else:
        preferred_order = [backend]

    for preferred_name in preferred_order:
        for item in detected:
            if item.name == preferred_name and item.supported:
                return item

    detected_names = ", ".join(f"{item.name}:{item.path}" for item in detected) or "无"
    raise FileNotFoundError(
        f"未找到可用的导出后端。请求后端: {backend}。"
        f" 已检测到: {detected_names}。"
        " 当前已实现的自动导出后端: LibreOffice、PowerPoint(Windows/macOS)、WPS(Windows)。"
    )


def export_pptx_to_pdf_with_libreoffice(pptx_path, output_pdf, office_bin=None):
    office_executable = find_backend(explicit_path=office_bin, backend="libreoffice").path
    source_path = Path(pptx_path).resolve()
    output_path = Path(output_pdf).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="ppt2fig-") as tmp_dir:
        command = [
            office_executable,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            tmp_dir,
            str(source_path),
        ]
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode != 0:
            stderr = result.stderr.strip()
            stdout = result.stdout.strip()
            details = stderr or stdout or "LibreOffice 导出失败"
            raise RuntimeError(details)

        converted_pdf = Path(tmp_dir) / f"{source_path.stem}.pdf"
        if not converted_pdf.exists():
            raise RuntimeError("LibreOffice 没有生成预期的 PDF 文件")

        shutil.copyfile(str(converted_pdf), str(output_path))
    return output_path


def export_pptx_to_pdf_with_powerpoint(
    pptx_path,
    output_pdf,
    *,
    intent="print",
    bitmap_missing_fonts=False,
):
    if platform.system() != "Windows":
        raise RuntimeError("PowerPoint 后端仅支持 Windows")

    try:
        import comtypes.client
    except ImportError as exc:
        raise RuntimeError("当前环境缺少 comtypes，无法调用 PowerPoint") from exc

    source_path = Path(pptx_path).resolve()
    output_path = Path(output_pdf).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    intent_value = 2 if intent == "print" else 1

    powerpoint = None
    presentation = None
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(
            str(source_path),
            WithWindow=True,
            ReadOnly=True,
        )
        if output_path.exists():
            output_path.unlink()
        presentation.ExportAsFixedFormat(
            str(output_path),
            2,
            Intent=intent_value,
            BitmapMissingFonts=bitmap_missing_fonts,
            UseISO19005_1=False,
        )
    except Exception as exc:
        raise RuntimeError(f"PowerPoint 导出失败: {exc}") from exc
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass

    if not output_path.exists():
        raise RuntimeError("PowerPoint 没有生成预期的 PDF 文件")
    return output_path


def export_pptx_to_pdf_with_wps(pptx_path, output_pdf):
    if platform.system() != "Windows":
        raise RuntimeError("WPS 后端仅支持 Windows")

    try:
        import comtypes.client
    except ImportError as exc:
        raise RuntimeError("当前环境缺少 comtypes，无法调用 WPS") from exc

    progid = _find_wps_com_progid()
    if not progid:
        raise RuntimeError("未找到可用的 WPS Presentation COM 接口")

    source_path = Path(pptx_path).resolve()
    output_path = Path(output_pdf).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    app = None
    presentation = None
    export_error = None
    try:
        app = comtypes.client.CreateObject(progid, dynamic=True)
        try:
            app.Visible = 1
        except Exception:
            pass

        presentation = app.Presentations.Open(str(source_path))

        if output_path.exists():
            output_path.unlink()

        try:
            presentation.ExportAsFixedFormat(str(output_path), 2)
        except Exception as exc:
            export_error = exc

        if not output_path.exists():
            try:
                # Inference from WPS community examples and PowerPoint compatibility:
                # 32 commonly maps to "save as PDF" in Presentation automation.
                presentation.SaveAs(str(output_path), 32)
            except Exception as exc:
                raise RuntimeError(
                    f"WPS 导出失败。ExportAsFixedFormat 错误: {export_error}; SaveAs(PDF) 错误: {exc}"
                ) from exc
    except Exception as exc:
        if isinstance(exc, RuntimeError):
            raise
        raise RuntimeError(f"WPS 导出失败: {exc}") from exc
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

    if not output_path.exists():
        raise RuntimeError("WPS 没有生成预期的 PDF 文件")
    return output_path


def export_selected_pages(
    pptx_path,
    output_pdf,
    pages,
    *,
    backend="auto",
    office_bin=None,
    powerpoint_intent="print",
    bitmap_missing_fonts=False,
    no_crop=False,
    percent_retain=0.0,
    margin_size=0.0,
    use_uniform=True,
    use_same_size=True,
    threshold=191,
):
    output_path = Path(output_pdf).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="ppt2fig-pages-") as tmp_dir:
        full_pdf = Path(tmp_dir) / "full.pdf"
        selected_backend = find_backend(explicit_path=office_bin, backend=backend)
        if selected_backend.name in {"libreoffice", "custom"}:
            export_pptx_to_pdf_with_libreoffice(pptx_path, full_pdf, office_bin=selected_backend.path)
        elif selected_backend.name == "powerpoint":
            export_pptx_to_pdf_with_powerpoint(
                pptx_path,
                full_pdf,
                intent=powerpoint_intent,
                bitmap_missing_fonts=bitmap_missing_fonts,
            )
        elif selected_backend.name == "wps":
            export_pptx_to_pdf_with_wps(pptx_path, full_pdf)
        else:
            raise RuntimeError(f"暂不支持使用后端 {selected_backend.name} 导出")
        extract_pdf_pages(full_pdf, output_path, pages)
        crop_pdf_file(
            output_path,
            no_crop=no_crop,
            percent_retain=percent_retain,
            margin_size=margin_size,
            use_uniform=use_uniform,
            use_same_size=use_same_size,
            threshold=threshold,
        )
    return output_path


def current_slide_2_pdf_windows(output_pdf_file):
    try:
        import comtypes.client
    except ImportError as exc:
        raise RuntimeError("当前环境缺少 comtypes，无法调用 PowerPoint") from exc

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    ppt_file = powerpoint.ActivePresentation
    output_pdf_file = os.path.abspath(output_pdf_file)
    if os.path.exists(output_pdf_file):
        os.remove(output_pdf_file)
    ppt_file.ExportAsFixedFormat(output_pdf_file, 2, RangeType=3)
    return True


def current_slide_2_pdf_mac(output_pdf_file):
    script = '''
    tell application "Microsoft PowerPoint"
        if not running then
            return "PowerPoint未启动"
        end if

        if (count of presentations) is 0 then
            return "没有打开的PPT文件"
        end if

        set pdfPath to "%s"
        set thePresentation to active presentation
        save active presentation in pdfPath as save as PDF
        return "success"
    end tell
    ''' % output_pdf_file

    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if "success" in result.stdout:
        return True
    raise RuntimeError(result.stdout.strip() or result.stderr.strip() or "转换过程出错")


def get_active_presentation_info():
    if platform.system() == "Windows":
        import comtypes.client

        try:
            powerpoint = comtypes.client.GetActiveObject("Powerpoint.Application")
            ppt_file = powerpoint.ActivePresentation
            return ppt_file.FullName, ppt_file.Name
        except Exception as exc:
            raise RuntimeError("PowerPoint未启动或没有打开的文件") from exc

    script = '''
    tell application "Microsoft PowerPoint"
        if not running then
            return "PowerPoint未启动"
        end if

        if (count of presentations) is 0 then
            return "没有打开的PPT文件"
        end if

        set thePresentation to active presentation
        return {full name of thePresentation, name of thePresentation}
    end tell
    '''
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if "PowerPoint未启动" in result.stdout or "没有打开的PPT文件" in result.stdout:
        raise RuntimeError(result.stdout.strip())

    output = result.stdout.strip().split(", ")
    if len(output) == 2:
        return output[0], output[1]
    raise RuntimeError("无法获取PPT文件信息")
