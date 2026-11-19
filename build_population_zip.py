"""
Build dist/land_use_population.zip - the Land Use Population QGIS plugin on its own,
ready for Plugins > Manage and Install Plugins > Install from ZIP.
(The same tool is also inside the Urban Design Setup pack: py build_pack_zip.py)

    py build_population_zip.py

Bundles python-pptx and XlsxWriter in vendor/ (openpyxl, lxml, Pillow and typing_extensions
ship with QGIS). Needs internet once to download them (cached by pip afterwards).
"""

import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "land_use_population"
SRC = os.path.join(HERE, "land_use_population_plugin")
DIST = os.path.join(HERE, "dist")
BUILD = os.path.join(DIST, "build_population", NAME)
VENDOR_PKGS = ["python-pptx", "XlsxWriter"]


def main():
    shutil.rmtree(os.path.dirname(BUILD), ignore_errors=True)
    os.makedirs(BUILD)
    for f in os.listdir(SRC):
        full = os.path.join(SRC, f)
        if os.path.isfile(full) and not f.endswith(".pyc"):
            shutil.copy2(full, BUILD)
    print("Downloading", ", ".join(VENDOR_PKGS))
    env = dict(os.environ, PYTHONNOUSERSITE="1")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                           "--no-deps", "--upgrade", "--target", os.path.join(BUILD, "vendor")] + VENDOR_PKGS, env=env)
    for root, dirs, _ in os.walk(BUILD):
        for d in list(dirs):
            if d == "__pycache__" or (d == "bin" and os.path.basename(root) == "vendor"):
                shutil.rmtree(os.path.join(root, d))
                dirs.remove(d)
    os.makedirs(DIST, exist_ok=True)
    zpath = os.path.join(DIST, f"{NAME}.zip")
    if os.path.exists(zpath):
        os.remove(zpath)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(BUILD):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.join(NAME, os.path.relpath(full, BUILD)))
    shutil.rmtree(os.path.dirname(BUILD), ignore_errors=True)
    with open(os.path.join(SRC, "metadata.txt"), encoding="utf-8") as fh:
        ver = next((l.split("=", 1)[1].strip() for l in fh if l.startswith("version=")), "?")
    print(f"Built {zpath}  (version {ver}, {os.path.getsize(zpath) / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
