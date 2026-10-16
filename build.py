"""Build a portable SWordBook package for the current system.

    python build.py

Result (in the "release" folder):
    Windows: SWordBook-v<version>-windows-x64.zip
    Linux:   SWordBook-v<version>-linux-x64.tar.gz
    macOS:   SWordBook-v<version>-macos-arm64.zip

It installs what it needs (PySide6, PyInstaller, Pillow) with pip, draws the
icon with the program's own code, runs PyInstaller and packs the result.
"""
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "swordbook.py")
NAME = "SWordBook"


def run(*cmd):
    print(">", " ".join(cmd), flush=True)
    subprocess.check_call(cmd, cwd=HERE)


def version():
    with open(SOURCE, encoding="utf-8") as f:
        return re.search(r'^APP_VERSION = "([^"]+)"', f.read(), re.M).group(1)


def target():
    machine = platform.machine().lower()
    arch = "arm64" if machine in ("arm64", "aarch64") else "x64"
    if sys.platform == "win32":
        return "windows", arch
    if sys.platform == "darwin":
        return "macos", arch
    return "linux", arch


def make_icons(build_dir):
    """Draw the app icon with swordbook.app_icon() and save png/ico/icns."""
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    sys.path.insert(0, HERE)
    from PySide6.QtWidgets import QApplication
    from PIL import Image
    app = QApplication.instance() or QApplication([])
    import swordbook
    png = os.path.join(build_dir, "icon.png")
    swordbook.app_icon().pixmap(256, 256).save(png)
    image = Image.open(png)
    image.save(os.path.join(build_dir, "icon.ico"),
               sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])
    try:
        image.save(os.path.join(build_dir, "icon.icns"))
    except Exception:
        pass
    del app
    return png


def main():
    ver = version()
    system, arch = target()
    package = f"{NAME}-v{ver}-{system}-{arch}"
    print(f"Building {package}", flush=True)

    run(sys.executable, "-m", "pip", "install", "--upgrade",
        "PySide6-Essentials", "pyinstaller", "pillow")

    build_dir = os.path.join(HERE, "build")
    dist_dir = os.path.join(HERE, "dist")
    out_dir = os.path.join(HERE, "release")
    for d in (build_dir, dist_dir):
        shutil.rmtree(d, ignore_errors=True)
    os.makedirs(build_dir)
    os.makedirs(out_dir, exist_ok=True)

    make_icons(build_dir)
    icon = os.path.join(build_dir, {"windows": "icon.ico", "macos": "icon.icns"}
                        .get(system, "icon.png"))
    if not os.path.exists(icon):
        icon = os.path.join(build_dir, "icon.png")

    # Only the Qt modules the program uses; the rest of PySide6 is left out
    excludes = ["QtWebEngineCore", "QtWebEngineWidgets", "QtQml", "QtQuick",
                "Qt3DCore", "QtMultimedia", "QtCharts", "QtDataVisualization",
                "QtPdf", "QtSql", "QtTest", "QtBluetooth", "QtNetwork",
                "QtOpenGL", "QtSvg", "tkinter"]
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
           "--windowed", "--name", NAME, "--icon", icon,
           "--workpath", build_dir, "--distpath", dist_dir,
           "--specpath", build_dir]
    for module in excludes:
        cmd += ["--exclude-module",
                module if module == "tkinter" else f"PySide6.{module}"]
    if system == "macos":
        cmd += ["--osx-bundle-identifier", "io.github.swordbook"]
    run(*cmd, SOURCE)

    # Package: the folder (Windows/Linux) or the .app (macOS)
    base = os.path.join(out_dir, package)
    if system == "macos":
        content = os.path.join(dist_dir, f"{NAME}.app")
    else:
        content = os.path.join(dist_dir, NAME)
        shutil.copy(os.path.join(HERE, "LICENSE"), content)
        shutil.copy(os.path.join(HERE, "README.md"), content)
    if system == "linux":
        archive = base + ".tar.gz"
        with tarfile.open(archive, "w:gz") as tar:
            tar.add(content, arcname=package)
    else:
        archive = base + ".zip"
        if system == "macos":       # ditto keeps the .app permissions
            run("ditto", "-c", "-k", "--keepParent", content, archive)
        else:
            with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
                for folder, _, files in os.walk(content):
                    for file in files:
                        path = os.path.join(folder, file)
                        z.write(path, os.path.join(
                            package, os.path.relpath(path, content)))
    size = os.path.getsize(archive) / 1024 / 1024
    print(f"\nDone: {archive} ({size:.0f} MB)")


if __name__ == "__main__":
    main()
