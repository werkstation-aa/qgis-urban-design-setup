"""
Build dist/qgis_automation_pack.zip - every tool of this folder as ONE QGIS plugin
("Urban Design Setup"),
ready for Plugins > Manage and Install Plugins > Install from ZIP.

    py build_pack_zip.py            release zip (stops if metadata.txt still has REPLACE_ME)
    py build_pack_zip.py --draft    local test zip, placeholders allowed

The pack is assembled from the files in this folder, so it always carries the
current versions:
    automation_pack/        plugin shell (menu, toolbar, Processing provider), LICENSE
    scripts/                qgis_project_setup_r2.py, qgis_project_arrange.py,
                            qgis_layout_automation.py (imported by the tool windows)
    scripts/styles/         styles/*.qml, shared by Project Setup and Project Arrange
    lup/                    lup_*.py + README.md            (Export Land Use)
    vector_pivot/           vector_pivot_plugin/*.py        (Vector Pivot Export)
    population/             land_use_population_plugin/pop_*  (Land Use Population)
    data/                   LUP template, RDTR colour library, colour settings
    vendor/                 python-pptx, XlsxWriter, typing_extensions, comtypes
                            (tests, COM server code, OLE icons and pip metadata removed;
                            each package's licence kept in vendor/LICENSES/)
Needs internet once to download the vendor packages (cached by pip afterwards).
"""

import json
import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "qgis_automation_pack"
DIST = os.path.join(HERE, "dist")
BUILD = os.path.join(DIST, "build_pack", NAME)
VENDOR_PKGS = ["python-pptx", "XlsxWriter", "typing_extensions", "comtypes"]

SCRIPTS = ["qgis_project_setup_r2.py", "qgis_project_arrange.py", "qgis_layout_automation.py"]
SHELL_EXCLUDE = {"icons_preview.png", "ICON_BRIEF.md"}  # design notes stay out of the plugin
# vendor parts the plugin never imports: tests (with binary type libraries), COM server
# registration, OLE-object icons of python-pptx
VENDOR_PRUNE = [os.path.join("comtypes", "test"), os.path.join("comtypes", "server"),
                os.path.join("pptx", "templates", "docx-icon.emf"),
                os.path.join("pptx", "templates", "generic-icon.emf"),
                os.path.join("pptx", "templates", "pptx-icon.emf"),
                os.path.join("pptx", "templates", "xlsx-icon.emf")]
PLACEHOLDER = "REPLACE_ME"
LUP = ["lup_pipeline.py", "lup_widget.py", "lup_colors.py", "lup_qgis_tool.py"]
LUP_README = "LUP_README.md"  # shipped as lup/README.md
# automation_pack/icons/ is the master copy; the standalone plugins get the same files
STANDALONE_ICONS = {"lup.svg": os.path.join("lup_tables_plugin", "icon.svg"),
                    "pivot.svg": os.path.join("vector_pivot_plugin", "icon.svg"),
                    "population.svg": os.path.join("land_use_population_plugin", "icon.svg")}
POPULATION = ["pop_core.py", "pop_layer.py", "pop_widget.py", "pop_pptx.py", "pop_reference.json",
              "population_template.pptx"]


def copy_files(src_dir, dst_dir, names=None, exclude=()):
    os.makedirs(dst_dir, exist_ok=True)
    for f in names or os.listdir(src_dir):
        src = os.path.join(src_dir, f)
        if f in exclude:
            continue
        if os.path.isfile(src) and not f.endswith(".pyc"):
            shutil.copy2(src, dst_dir)
        elif os.path.isdir(src) and f != "__pycache__":
            copy_files(src, os.path.join(dst_dir, f))


def check_metadata(draft):
    """A release zip must not carry placeholder email / links."""
    with open(os.path.join(HERE, "automation_pack", "metadata.txt"), encoding="utf-8") as fh:
        todo = [l.split("=", 1)[0] for l in fh if PLACEHOLDER in l]
    if todo and not draft:
        sys.exit("metadata.txt still has %s in: %s. Fill them in, or build with --draft for a local test."
                 % (PLACEHOLDER, ", ".join(todo)))
    if todo:
        print("DRAFT build - not for upload, placeholders in:", ", ".join(todo))


def prune_vendor(vendor):
    """Drop unused vendor parts; keep each package's licence in vendor/LICENSES/."""
    for rel in VENDOR_PRUNE:
        p = os.path.join(vendor, rel)
        if os.path.isdir(p):
            shutil.rmtree(p)
        elif os.path.isfile(p):
            os.remove(p)
    licences = os.path.join(vendor, "LICENSES")
    os.makedirs(licences, exist_ok=True)
    for d in os.listdir(vendor):
        if not d.endswith(".dist-info"):
            continue
        pkg = d.split("-", 1)[0]
        info = os.path.join(vendor, d)
        for root, _, files in os.walk(info):
            for f in files:
                if f.upper().startswith(("LICENSE", "LICENCE", "COPYING")):
                    shutil.copy2(os.path.join(root, f), os.path.join(licences, pkg + ".txt"))
        shutil.rmtree(info)


def main():
    draft = "--draft" in sys.argv[1:]
    check_metadata(draft)
    shutil.rmtree(os.path.dirname(BUILD), ignore_errors=True)  # OneDrive may keep a lock: tolerate leftovers
    copy_files(os.path.join(HERE, "automation_pack"), BUILD, exclude=SHELL_EXCLUDE)
    copy_files(HERE, os.path.join(BUILD, "scripts"), SCRIPTS)
    copy_files(os.path.join(HERE, "styles"), os.path.join(BUILD, "scripts", "styles"))
    copy_files(HERE, os.path.join(BUILD, "lup"), LUP)
    shutil.copy2(os.path.join(HERE, LUP_README), os.path.join(BUILD, "lup", "README.md"))
    copy_files(os.path.join(HERE, "vector_pivot_plugin"), os.path.join(BUILD, "vector_pivot"),
               ["__init__.py", "vpe_core.py", "vpe_layer.py", "vpe_algorithm.py", "vpe_widget.py"])
    with open(os.path.join(BUILD, "vector_pivot", "__init__.py"), "w") as fh:
        fh.write("")  # subpackage of the pack: no classFactory of its own
    copy_files(os.path.join(HERE, "land_use_population_plugin"), os.path.join(BUILD, "population"), POPULATION)
    with open(os.path.join(BUILD, "population", "__init__.py"), "w") as fh:
        fh.write("")
    for name, dst in STANDALONE_ICONS.items():
        shutil.copy2(os.path.join(HERE, "automation_pack", "icons", name), os.path.join(HERE, dst))
    brief = os.path.join(BUILD, "icons", "ICON_BRIEF.md")  # design notes stay out of the plugin
    if os.path.exists(brief):
        os.remove(brief)
    if not os.path.isfile(os.path.join(BUILD, "LICENSE")):
        sys.exit("automation_pack/LICENSE is missing: the QGIS plugin repository needs one.")

    # LUP defaults, copied to <profile>/lup_tool on first run
    data = os.path.join(BUILD, "data")
    os.makedirs(data, exist_ok=True)
    shutil.copy2(os.path.join(HERE, "LUP_template.pptx"), os.path.join(data, "LUP_template.pptx"))
    shutil.copy2(os.path.join(HERE, "RDTR_Colour Code.xlsx"), os.path.join(data, "RDTR_Colour Code.xlsx"))
    with open(os.path.join(HERE, "lup_colours.json"), encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg.update(library="", aliases=cfg.get("aliases") or {}, overrides={})  # portable: use bundled library
    with open(os.path.join(data, "lup_colours.json"), "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, indent=2, ensure_ascii=False)

    print("Downloading", ", ".join(VENDOR_PKGS))
    env = dict(os.environ, PYTHONNOUSERSITE="1")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
                           "--no-deps", "--upgrade", "--target", os.path.join(BUILD, "vendor")] + VENDOR_PKGS,
                          env=env)
    prune_vendor(os.path.join(BUILD, "vendor"))

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
    with open(os.path.join(HERE, "automation_pack", "metadata.txt"), encoding="utf-8") as fh:
        ver = next((l.split("=", 1)[1].strip() for l in fh if l.startswith("version=")), "?")
    print(f"Built {zpath}  (version {ver}, {os.path.getsize(zpath) / 1e6:.1f} MB)"
          + ("  DRAFT - not for upload" if draft else ""))


if __name__ == "__main__":
    main()
