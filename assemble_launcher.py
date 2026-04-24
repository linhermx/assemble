from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

import requests


REPO = "linhermx/assemble"
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"

ASSET_NAME = "assemble_windows.exe"
APP_EXE_PREFIX = "assemble_v"
APP_EXE_SUFFIX = ".exe"


def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def ensure_dirs(root: Path) -> dict[str, Path]:
    app_dir = root / "app"
    downloads_dir = root / "downloads"
    logs_dir = root / "logs"
    app_dir.mkdir(parents=True, exist_ok=True)
    downloads_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    return {"root": root, "app": app_dir, "downloads": downloads_dir, "logs": logs_dir}


def parse_version(tag: str) -> tuple[int, int, int] | None:
    match = re.match(r"^v?(\d+)\.(\d+)\.(\d+)$", tag.strip())
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def parse_version_from_name(name: str) -> tuple[int, int, int] | None:
    match = re.match(
        rf"^{re.escape(APP_EXE_PREFIX)}(\d+)\.(\d+)\.(\d+){re.escape(APP_EXE_SUFFIX)}$",
        name,
    )
    if not match:
        return None
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def find_installed_app(app_dir: Path) -> tuple[tuple[int, int, int], Path] | None:
    candidates = []
    for item in app_dir.glob(f"{APP_EXE_PREFIX}*{APP_EXE_SUFFIX}"):
        version = parse_version_from_name(item.name)
        if version:
            candidates.append((version, item))
    if not candidates:
        return None
    candidates.sort(key=lambda value: value[0], reverse=True)
    return candidates[0]


def get_latest_release() -> tuple[tuple[int, int, int], str, str]:
    response = requests.get(LATEST_API, timeout=20)
    response.raise_for_status()
    data = response.json()

    tag = data.get("tag_name", "")
    version = parse_version(tag)
    if not version:
        raise RuntimeError(f"Tag invalido en latest release: {tag!r}")

    assets = data.get("assets", []) or []
    asset = next((item for item in assets if item.get("name") == ASSET_NAME), None)
    if not asset:
        raise RuntimeError(
            f"No se encontro el asset '{ASSET_NAME}' en el latest release.\n"
            "Sube ese ejecutable como asset en GitHub Releases."
        )

    url = asset.get("browser_download_url", "")
    if not url:
        raise RuntimeError("Asset sin browser_download_url")
    return version, tag, url


def download_file(url: str, dst: Path) -> None:
    with requests.get(url, stream=True, timeout=60) as response:
        response.raise_for_status()
        with dst.open("wb") as file:
            for chunk in response.iter_content(chunk_size=1024 * 256):
                if chunk:
                    file.write(chunk)


def run_app(exe_path: Path) -> None:
    subprocess.Popen([str(exe_path)], cwd=str(exe_path.parent))
    raise SystemExit(0)


def run_local_source() -> None:
    source_entry = base_dir() / "assemble_gui.py"
    subprocess.Popen([sys.executable, str(source_entry)], cwd=str(source_entry.parent))
    raise SystemExit(0)


def main() -> None:
    root = base_dir()
    dirs = ensure_dirs(root)

    tk_root = tk.Tk()
    tk_root.withdraw()

    installed = find_installed_app(dirs["app"])
    installed_version = installed[0] if installed else None
    installed_exe = installed[1] if installed else None

    try:
        latest_version, latest_tag, asset_url = get_latest_release()
    except Exception as exc:
        if installed_exe:
            messagebox.showwarning(
                "Actualizacion no disponible",
                f"No se pudo verificar update:\n{exc}\n\nSe abrira la version instalada.",
            )
            run_app(installed_exe)

        if not getattr(sys, "frozen", False):
            messagebox.showwarning(
                "Sin conexion a releases",
                f"No se pudo verificar update:\n{exc}\n\nSe abrira la GUI local.",
            )
            run_local_source()

        messagebox.showerror(
            "No hay app instalada",
            f"No se pudo verificar update y no existe app instalada.\n\nDetalle:\n{exc}",
        )
        return

    needs_install = installed_exe is None
    needs_update = (installed_version is None) or (latest_version > installed_version)

    if needs_install or needs_update:
        message = (
            f"Hay una version disponible: {latest_tag}\n"
            f"Instalada: {'ninguna' if installed_version is None else 'v' + '.'.join(map(str, installed_version))}\n\n"
            "Deseas actualizar ahora?"
        )
        if needs_install or messagebox.askyesno("Actualizacion disponible", message):
            tmp = dirs["downloads"] / ASSET_NAME
            try:
                download_file(asset_url, tmp)
                target = dirs["app"] / (
                    f"{APP_EXE_PREFIX}{latest_version[0]}.{latest_version[1]}.{latest_version[2]}{APP_EXE_SUFFIX}"
                )
                shutil.move(str(tmp), str(target))
                installed_exe = target
            except Exception as exc:
                if installed_exe:
                    messagebox.showwarning(
                        "Update fallo",
                        f"No se pudo actualizar:\n{exc}\n\nSe abrira la version instalada.",
                    )
                elif not getattr(sys, "frozen", False):
                    messagebox.showwarning(
                        "Update fallo",
                        f"No se pudo descargar la version publicada:\n{exc}\n\nSe abrira la GUI local.",
                    )
                    run_local_source()
                else:
                    messagebox.showerror("Update fallo", f"No se pudo descargar/instalar:\n{exc}")
                    return

    if installed_exe:
        run_app(installed_exe)

    if not getattr(sys, "frozen", False):
        run_local_source()

    messagebox.showerror("No se encontro la app", "No hay ejecutable para abrir.")


if __name__ == "__main__":
    main()
