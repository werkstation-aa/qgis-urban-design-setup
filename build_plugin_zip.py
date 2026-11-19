"""
Build dist/lup_tables.zip - the Export Land Use QGIS plugin, ready for
Plugins > Manage and Install Plugins > Install from ZIP.

    py build_plugin_zip.py

The zip is self-contained: plugin shell, calculations, window, Processing tool,
colour library, slide template and the pure-Python packages QGIS lacks
(python-pptx, XlsxWriter, typing_extensions, comtypes). Needs internet once to
download those packages (cached by pip afterwards).
"""

import json
import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "lup_tables"
DIST = os.path.join(HERE, "dist")
BUILD = os.path.join(DIST, "build", NAME)
VENDOR_PKGS = ["python-pptx", "XlsxWriter", "typing_extensions", "comtypes"]


def main():
    shutil.rmtree(os.path.join(DIST, "build"), ignore_errors=True)
    os.makedirs(os.path.join(BUILD, "data"))

    for f in os.listdir(os.path.join(HERE, "lup_tables_plugin")):
        src = os.path.join(HERE, "lup_tables_plugin", f)
        if os.path.isfile(src) and not f.endswith(".pyc"):
            shutil.copy2(src, BUILD)
    for f in ("lup_pipeline.py", "lup_widget.py", "lup_colors.py", "lup_qgis_tool.py", "README.md"):
        shutil.copy2(os.path.join(HERE, f), BUILD)

    # default data, copied to <profile>/lup_tool on first run
    shutil.copy2(os.path.join(HERE, "LUP_template.pptx"), os.path.join(BUILD, "data", "LUP_template.pptx"))
    shutil.copy2(os.path.join(HERE, "RDTR_Colour Code.xlsx"), os.path.join(BUILD, "data", "RDTR_Colour Code.xlsx"))
    with open(os.path.join(HERE, "lup_colours.json"), encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg.update(library="", aliases=cfg.get("aliases") or {}, overrides={})  # portable: use bundled library
    with open(os.path.join(BUILD, "data", "lup_colours.json"), "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2, ensure_ascii=False)

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
    shutil.rmtree(os.path.join(DIST, "build"), ignore_errors=True)
    with open(os.path.join(HERE, "lup_tables_plugin", "metadata.txt"), encoding="utf-8") as fh:
        ver = next((l.split("=", 1)[1].strip() for l in fh if l.startswith("version=")), "?")
    print(f"Built {zpath}  (version {ver}, {os.path.getsize(zpath) / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
