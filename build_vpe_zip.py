"""
Build dist/vector_pivot_export.zip - the Vector Pivot Export QGIS plugin, ready for
Plugins > Manage and Install Plugins > Install from ZIP.

    py build_vpe_zip.py

Bundles the pure-Python packages QGIS lacks (XlsxWriter, comtypes) in vendor/.
Needs internet once to download them (cached by pip afterwards).
"""

import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "vector_pivot_plugin")
NAME = "vector_pivot_export"
DIST = os.path.join(HERE, "dist")
BUILD = os.path.join(DIST, "build_vpe", NAME)
VENDOR_PKGS = ["XlsxWriter", "comtypes"]


def main():
    shutil.rmtree(os.path.dirname(BUILD), ignore_errors=True)  # OneDrive may keep a lock: tolerate leftovers
    os.makedirs(BUILD, exist_ok=True)
    for f in os.listdir(SRC):
        if os.path.isfile(os.path.join(SRC, f)) and not f.endswith(".pyc"):
            shutil.copy2(os.path.join(SRC, f), BUILD)

    print("Downloading", ", ".join(VENDOR_PKGS))
    env = dict(os.environ, PYTHONNOUSERSITE="1")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                           "--no-deps", "--target", os.path.join(BUILD, "vendor")] + VENDOR_PKGS, env=env)

    for root, dirs, _ in os.walk(BUILD):  # no caches / console scripts in the zip
        for d in list(dirs):
            if d == "__pycache__" or (d == "bin" and os.path.basename(root) == "vendor"):
                shutil.rmtree(os.path.join(root, d))
                dirs.remove(d)

    zpath = os.path.join(DIST, f"{NAME}.zip")
    if os.path.exists(zpath):
        os.remove(zpath)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(BUILD):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.join(NAME, os.path.relpath(full, BUILD)))
    shutil.rmtree(os.path.dirname(BUILD), ignore_errors=True)
    print(f"Built {zpath}  ({os.path.getsize(zpath) / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
