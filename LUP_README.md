# Land Use Plan → Excel database → PowerPoint tables

Turns the land-use polygon layer into:

- **`*_LUP_Database.xlsx`**: the Excel database (sheets below).
- **`*_LUP_Tables.pptx`**: land-use summary slide and program/industrial-breakdown slide, on the theme of `LUP_template.pptx`.

## Excel database

| Sheet | Contents |
|---|---|
| **Overall MP** | Pivot: Type › SUBZONA, with area (sqm, Ha) and % of total. Next to it is a site-total check showing the Road adjustment used on the slides. |
| **Industrial** | Pivot filtered to SUBZONA = Industrial: area by IND_TYPE |
| **Phasing** | Pivot: Phase › SUBZONA |
| **Intensitas** | Pivot: Type › SUBZONA › Parcel_ID, with BCR, FAR, height, GOSR, area, GFA, population |
| **Population** | Pivot: Pop_Type › population |
| **Source** | The layer's attribute table exactly as it is, as an Excel table (`tblSource`). All pivots use it. |
| **Colours** | Colour legend: value, swatch, HEX, RGB, where the colour came from, library match, feature count, area |
| **Checks** | Data-quality warnings and the attribute vs polygon area check |

- **Pivot tables** are created by Microsoft Excel in the background, so Excel must be installed. After editing Source (adding rows is fine, since it's a table), use **Data → Refresh All**. Without Excel, or with pivots switched off, the summary sheets are written as SUMIFS formulas instead.
- **Colours** are applied with conditional formatting (one "cell value = …" rule per land use) on the Source sheet and on the pivot label columns. They therefore follow the values when you edit or refresh.

## Colour coding (layer symbology ⇄ Excel cells ⇄ slides)

One colour scheme is used for the layer style, the Excel cells and the slides. It's keyed on the **Colour by field** (e.g. `SUBZONA` or `ZONA`). Each value takes its colour from the first source that has one:

1. **Manual colour** set in the window (always wins)
2. Then, in the chosen **Colour source** order:
   - **Layer symbology**: the layer's categorised style on that field
   - **Colour library**: `RDTR_Colour Code.xlsx` (Zona | Kode Z | Sub-Zona | Kode SZ | RGB), the RDTR zone colour codes of Peraturan Menteri ATR/Kepala BPN No. 14 Tahun 2021 (Pedoman Penyusunan Basis Data dan Penyajian Peta RTRW dan RDTR). Values are matched by Sub-Zona or Zona name, code, an alias (e.g. *Industrial → KPI*, *Road → BJ*, *Green → RTH-8*), or a close spelling.
   - **Project palette**: the default colours in `lup_pipeline.py` (`COLORS`)

In the window's **Colours** tab:
- **Check the chain:** each value shows its layer symbol, library match, the colour used and how it will look in Excel and on the slides.
- **Fix a match:** choose a different library entry from the list, or click the colour to set it by hand.
- **Write to the layer:** **Apply these colours to the layer style** writes the scheme to the layer as categorised symbology. In standalone mode it's also saved as the shapefile's default `.qml`.

When colouring by a field other than SUBZONA (e.g. ZONA), the SUBZONA rows on the slides and pivots take the colour of their main ZONA value. Settings are saved in `lup_colours.json`.

There are three ways to run it. All use the same calculations in `lup_pipeline.py`.

## A. Export Land Use window (preview + export)

A resizable window with the settings on the left and a live preview on the right:
- **Slide 1 / Slide 2 tabs:** the slides drawn from the same layout data as the `.pptx`, so what you see is what gets exported. Zoom with **Fit**, **Fit width**, any % (type e.g. `130%`), the − / + buttons, or **Ctrl + mouse wheel**. The view scrolls when zoomed in.
- **Overall MP, Industrial, Population, Phasing, Intensitas, Checks tabs:** the Excel sheets' values. The Checks tab shows ⚠ and a count when there are data problems.
- **Panels:** drag the dividers to resize the settings, preview and log panes. The settings panel scrolls when the window is small. Window size, divider positions and settings are remembered.
- **Live preview:** follows the layer's selection and edits. "Selected features only" reports on the selection with its own area as the total.
- **Shortcuts:** **F5** refreshes the preview, **Ctrl+S** generates the files.

**In QGIS:** run `Install_QGIS_Tool.bat`, restart QGIS, then tick **Export Land Use** once under Plugins → Manage and Install Plugins → Installed. Open it from the toolbar button or Plugins → Land Use Plan. It starts on the active layer and can use any polygon layer in the project.

**Without opening QGIS:** double-click `Run_LUP_Widget.bat`. It uses QGIS's Python, opens the last file used or the shapefile in `Shapefile\`, and has an **Open file…** button for others.

## B. Processing tool (batch / models)

**Install (once per PC, and again after editing the scripts):** double-click `Install_QGIS_Tool.bat`.
- It copies the tool into your QGIS profile and downloads `python-pptx` once into a private folder. No admin rights are needed, and QGIS's own Python is not modified.
- For a QGIS profile other than `default`: `Install_QGIS_Tool.bat ProfileName`

**Run:** Processing Toolbox → **Scripts → Land Use Plan → Export Land Use (Excel + PowerPoint)**. If QGIS was already open, click the toolbox's refresh button first.

| Option | Notes |
|---|---|
| Land use layer | Any polygon layer (shapefile, GeoPackage, PostGIS…). Tick **Selected features only** to report on a selection. |
| Site total area | GRAND TOTAL in sqm (default 5,000,000). Road absorbs the difference from the data total. `0` = plain sum of the layer. |
| Take parcel area from | The `Area_sqm` field, or recalculate from the polygon geometry (use this after editing shapes). |
| Number format | Indonesian (1.234.567 / 12,34) or English. |
| PowerPoint template | Defaults to `LUP_template.pptx`. |
| Advanced → field names | Change these if another project's layer uses different field names. |

Installed files (in `%APPDATA%\QGIS\QGIS3\profiles\default\`):
- `processing\scripts\lup_qgis_tool.py`: the Processing tool
- `python\plugins\lup_tables\`: the plugin button for the window
- `lup_tool\`: `lup_pipeline.py`, `lup_widget.py`, `lup_colors.py`, `lup_colours.json`, the template, a copy of the colour library, and python-pptx + comtypes

## Project files

| File | What it is |
|---|---|
| `lup_pipeline.py` | All calculations, the Excel writer, and the slide layout used by both the `.pptx` and the preview |
| `lup_widget.py` | The Export Land Use window |
| `lup_colors.py` | Colour scheme: layer symbology, RDTR colour library matching, project palette |
| `lup_colours.json` | Colour settings (library path, field, priority, chosen matches, manual colours) |
| `RDTR_Colour Code.xlsx` | RDTR zone colour codes, Permen ATR/KBPN No. 14 Tahun 2021 (used when no other library path is set) |
| `LUP_template.pptx` | Slide template (theme and layout only, no content) |
| `lup_tables_plugin\` | QGIS plugin that opens the window |
| `lup_qgis_tool.py` | QGIS Processing tool |
| `Install_QGIS_Tool.bat` | Installs or updates the window, plugin and Processing tool in QGIS |
| `Run_LUP_Widget.bat` | Opens the window without starting QGIS |
| `Run_LUP_Automation.bat` | Command-line run, no QGIS needed |

## C. Command line (no QGIS needed)

Double-click `Run_LUP_Automation.bat` (outputs go to `Output\`), or:

```
py lup_pipeline.py --shp "Shapefile\Other.shp" --site-total 5000000
py lup_pipeline.py --color-field ZONA --priority library
py lup_pipeline.py --no-pivots
```

- `--color-field ZONA` colours by the ZONA field.
- `--priority library` puts the colour library first.
- `--no-pivots` writes formula sheets instead of pivot tables.

## Required fields

- **Required:** `Type` (Saleable / Non-Saleable), `SUBZONA`, `Area_sqm`. `Area_sqm` is not needed if areas are recalculated from geometry.
- **Optional:** `IND_TYPE`, `Phase`, `Parcel_ID`, `Pop_Total`, `Pop_Type`, `BCR (%)`, `FAR`, `Max Height`, `GOSR (%)`

Row order, chart labels and the project palette are set at the top of `lup_pipeline.py` (`SALEABLE_ORDER`, `NON_SALEABLE_ORDER`, `SHORT`, `COLORS`). Default English → RDTR aliases are in `lup_colors.py` (`DEFAULT_ALIASES`).

## Troubleshooting

- **"Cannot write … close it in Excel/PowerPoint"**: the output file is open; close it and run again.
- **"Missing Python package (pptx)"**: run `Install_QGIS_Tool.bat` again; it needs internet once.
- **"Pivot tables need Microsoft Excel … Writing formula sheets instead"**: Excel isn't installed or couldn't start. The workbook is still complete, with formula summary sheets.
- **NumPy errors in the QGIS log** ("compiled using NumPy 1.x"): packages installed with `pip install --user` for Python 3.12 are also loaded by QGIS. A newer NumPy there breaks GDAL's NumPy bridge, which some raster tools use. This LUP tool is not affected. To fix it for QGIS only, add an environment variable in QGIS → Settings → Options → System → Environment: `PYTHONNOUSERSITE` = `1`, then restart QGIS. This also hides those per-user packages from QGIS.
