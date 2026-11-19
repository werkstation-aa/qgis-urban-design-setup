# Urban Design Setup (QGIS plugin)

Urban design project tools in one QGIS plugin, under the **Urban Design Setup** menu and toolbar.

| Tool | What it does |
|---|---|
| **Project Setup** | Builds the standard layer tree, adds basemaps, live Indonesian datasets (ArcGIS REST), OSM local roads and QML styles |
| **Project Arrange** | Sorts the loaded layers into the setup tree and styles them |
| **Layout Automation** | Spatial bookmark × map theme → print layout pages and 300 dpi images |
| **Export Land Use** | Land use plan → Excel database with pivots + PowerPoint tables and charts ([details](LUP_README.md)) |
| **Export Pivot to Excel** | Any vector layer → Excel pivot coloured by its symbology |
| **Land Use Population** | Resident / visitor / worker population per parcel from SNI 03-1733-2004 and Permen PU 26/2008 standards, with Excel and PowerPoint output |

Export Land Use and Export Pivot are also available in the Processing Toolbox. Excel pivot tables need Microsoft Excel on Windows; without it, formula sheets are written instead. Everything else works on any platform.

## Install

- From QGIS: **Plugins → Manage and Install Plugins**, search for *Urban Design Setup*.
- Or download the zip from the releases and use **Install from ZIP**.

Requires QGIS 3.22 or newer.

## Repository layout

| Path | Contents |
|---|---|
| `automation_pack/` | Plugin shell: `metadata.txt`, `plugin.py`, tool windows, icons, `LICENSE` |
| `qgis_project_setup_r2.py`, `qgis_project_arrange.py`, `qgis_layout_automation.py` | Project Setup / Arrange / Layout scripts (also runnable from the QGIS Python Console) |
| `styles/` | QML layer styles used by Project Setup and Project Arrange |
| `lup_*.py`, `LUP_template.pptx`, `RDTR_Colour Code.xlsx`, `lup_colours.json` | Export Land Use |
| `vector_pivot_plugin/` | Export Pivot to Excel |
| `land_use_population_plugin/` | Land Use Population |
| `lup_tables_plugin/` | Standalone Export Land Use plugin (older packaging) |
| `build_pack_zip.py` | Builds the plugin zip |

## Build the plugin zip

```
py build_pack_zip.py
```

The zip is written to `dist/qgis_automation_pack.zip`. The build downloads python-pptx, XlsxWriter, typing_extensions and comtypes once with pip and bundles them in `vendor/` (their licences are kept in `vendor/LICENSES/`).

## License

GPL-2.0-or-later. See [automation_pack/LICENSE](automation_pack/LICENSE).

## Issues

Report bugs and requests in the [issue tracker](../../issues).
