"""Сборка релиза: каталог ``dist/kl`` через PyInstaller и установщик через Inno Setup.

Запуск:  uv run python tools/make_release.py [--no-installer]

Версия берётся из ``kinolist/__init__.py`` и передаётся и в ресурс версии exe,
и в сценарий установщика, поэтому менять её нужно только в одном месте.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

# Корень проекта: родитель каталога tools, в котором лежит этот скрипт.
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from kinolist import __author__, __version__  # noqa: E402

APP_NAME = "kl"
DISPLAY_NAME = "Kinolist Lib"
DIST = ROOT / "dist"
# PyInstaller собирает сюда, а затем содержимое переносится в dist/kl. Так сборка не падает,
# если dist/kl открыт в проводнике или является текущим каталогом консоли: удалить такой
# каталог нельзя, а вот заменить файлы внутри него можно.
STAGING = ROOT / "build" / "staging"
ENTRY_POINT = "kinolist_lib.py"
ICON = "images/icon.ico"
# Каталоги ресурсов копируются в сборку с теми же именами: так их находит kinolist.resources
# через sys._MEIPASS.
DATA_DIRS = ["templates", "images"]
INSTALLER_SCRIPT = "kinolist_lib.iss"
ISCC_CANDIDATES = [
    Path(R"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
    Path(R"C:\Program Files\Inno Setup 6\ISCC.exe"),
]

#: Ресурс версии Windows. Без него Диспетчер задач и свойства файла показывают только имя exe.
VERSION_INFO = """VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={nums}, prodvers={nums}, mask=0x3f, flags=0x0,
    OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0),
  ),
  kids=[
    StringFileInfo([
      StringTable(
        "040904B0",
        [
          StringStruct("CompanyName", "{author}"),
          StringStruct("FileDescription", "{name}"),
          StringStruct("FileVersion", "{version}"),
          StringStruct("InternalName", "{app}"),
          StringStruct("LegalCopyright", "(c) 2022-{year} {author}"),
          StringStruct("OriginalFilename", "{app}.exe"),
          StringStruct("ProductName", "{name}"),
          StringStruct("ProductVersion", "{version}"),
        ],
      )
    ]),
    VarFileInfo([VarStruct("Translation", [1033, 1200])]),
  ],
)
"""


def version_numbers(version: str) -> tuple[int, int, int, int]:
    """``"0.3.0"`` -> ``(0, 3, 0, 0)``; нечисловые части отбрасываются."""
    numbers = [int(part) for part in version.split(".") if part.isdigit()]
    return tuple([*numbers, 0, 0, 0, 0][:4])  # type: ignore[return-value]


def write_version_file(directory: Path) -> Path:
    """Пишет ресурс версии во временный каталог.

    Не в ``build/``: это рабочий каталог PyInstaller, и ``--clean`` очищает его до чтения файла.
    """
    path = directory / "version_info.txt"
    path.write_text(
        VERSION_INFO.format(
            nums=version_numbers(__version__),
            name=DISPLAY_NAME,
            app=APP_NAME,
            version=__version__,
            author=__author__,
            year=date.today().year,
        ),
        encoding="utf-8",
    )
    return path


def pyinstaller_command(version_file: Path) -> list[str]:
    cmd = [
        "uv", "run", "pyinstaller",
        "--clean",
        "--noconfirm",
        "--onedir",
        "--console",
        "--name", APP_NAME,
        "--icon", str(ROOT / ICON),
        "--version-file", str(version_file),
        "--distpath", str(STAGING),
        # Генерируемый kl.spec не нужен в корне проекта: build/ игнорируется git.
        # Относительные пути в spec считаются от его каталога, поэтому ниже пути абсолютные.
        "--specpath", str(ROOT / "build"),
    ]  # fmt: skip
    for data in DATA_DIRS:
        cmd += ["--add-data", f"{ROOT / data}{os.pathsep}{data}"]
    cmd.append(str(ROOT / ENTRY_POINT))
    return cmd


def find_iscc() -> Path | None:
    """Компилятор Inno Setup: переменная ISCC, PATH или стандартные каталоги установки."""
    env = os.environ.get("ISCC")
    if env and Path(env).is_file():
        return Path(env)
    found = shutil.which("ISCC")
    if found:
        return Path(found)
    return next((path for path in ISCC_CANDIDATES if path.is_file()), None)


def run(cmd: list[str]) -> int:
    print("Запуск:", " ".join(cmd))
    # Из корня проекта, чтобы относительные пути выше работали независимо от текущего каталога.
    return subprocess.run(cmd, cwd=ROOT, check=False).returncode


def replace_contents(target: Path, source: Path) -> None:
    """Заменяет содержимое ``target`` содержимым ``source``, сохраняя сам каталог ``target``."""
    target.mkdir(parents=True, exist_ok=True)
    for item in target.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()
    for item in source.iterdir():
        shutil.move(str(item), str(target / item.name))


def build_exe() -> int:
    shutil.rmtree(STAGING, ignore_errors=True)
    with tempfile.TemporaryDirectory() as tmp:
        code = run(pyinstaller_command(write_version_file(Path(tmp))))
    if code == 0:
        replace_contents(DIST / APP_NAME, STAGING / APP_NAME)
        shutil.rmtree(STAGING, ignore_errors=True)
    return code


def build_installer() -> int:
    iscc = find_iscc()
    if iscc is None:
        print("Inno Setup не найден (ISCC.exe), установщик не собран.", file=sys.stderr)
        return 1
    return run([str(iscc), f"/DMyAppVersion={__version__}", INSTALLER_SCRIPT])


def main() -> None:
    parser = argparse.ArgumentParser(description=f"Сборка релиза {DISPLAY_NAME} {__version__}.")
    parser.add_argument("--no-installer", action="store_true", help="собрать только каталог dist/kl без установщика")
    args = parser.parse_args()

    code = build_exe()
    if code != 0:
        sys.exit(code)
    print(f"\n=== Сборка {DISPLAY_NAME} {__version__} создана в dist/{APP_NAME} ===")

    if args.no_installer:
        return
    code = build_installer()
    if code != 0:
        sys.exit(code)
    print(f"\n=== Установщик создан в dist ({DISPLAY_NAME} {__version__} Setup.exe) ===")


if __name__ == "__main__":
    main()
