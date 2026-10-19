# Kostika cleaners

Two Python tk-GUI tools that turn Priority ERP exports into clean procurement Excels for Kostika operators. Each is shipped as a single Windows `.exe` built by GitHub Actions.

## Tools

| Script | Purpose | Latest version |
|---|---|---|
| `trisim_purchase_cleaner.py` | Clean shutter (תריסים) orders. Rounds height up to nearest multiple of חלוקה (slat profile). | see `APP_VERSION` |
| `aluminum_cleaner.py` | Clean material/profile orders. Dialog edits bar count and length per profile and recalculates total kg. | see `APP_VERSION` |

## Material order (aluminum 2.5)

Double-click the exe, choose the original Priority export and an output location,
then confirm/edit each profile's bar count and bar length (metres) in the dialog.
The old second quantities-spreadsheet prompt is replaced by this dialog.
Only the SKU is shown; repeated SKUs also show their color so their source rows
can be edited independently. Set count to zero to omit a row. Cancel stops export.

Weight uses the original export's total kg / (bar count × length in metres) as
the kg/metre rate. Example: 10 × 6 m weighing 15 kg gives 0.25 kg/m;
changing the order to 14 × 6 m produces 21 kg. The Excel `מידה` column remains
in millimetres, as in the Priority input. Invalid/missing source rates are reported
instead of silently exporting an incorrect weight.

Existing command-line batch usage and `--qty-update` remain available.

## Master mapping

`kostika_mapping.xlsx` (formerly `מקטים תריסים כולל.xlsx`, Hebrew name still accepted as fallback) — SKU → supplier / חלוקה / description. Bundled into both `.exe` files. Operators can drop an updated copy next to the `.exe` to override the bundled default without rebuilding.

## Get the latest .exe

1. Open the [Actions tab](../../actions/workflows/build.yml).
2. Click the most recent successful run.
3. Scroll to **Artifacts** at the bottom and download `kostika-cleaners-trisim-<ver>-aluminum-<ver>.zip`.
4. Inside: both `.exe` files, `kostika_mapping.xlsx`, and `BUILD.txt` (with manual rebuild instructions).

For a tagged release (`git tag v1.x && git push --tags`), the same artifacts get attached to the GitHub Release page.

## Build locally on Windows

See `BUILD.txt` for the manual PyInstaller commands (one-liner each, `--onefile --windowed`).

## Project state and recent changes

- `git log --oneline` for the chronology.
- The `/Users/idant/.claude/plans/load-this-as-a-synchronous-haven.md` plan file (local) holds the original feedback round + audit punch list.
