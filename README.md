# Kostika cleaners

Two Python tk-GUI tools that turn Priority ERP exports into clean procurement Excels for Kostika operators. Each is shipped as a single Windows `.exe` built by GitHub Actions.

## Tools

| Script | Purpose | Latest version |
|---|---|---|
| `trisim_purchase_cleaner.py` | Clean shutter (תריסים) orders. Rounds height up to nearest multiple of חלוקה (slat profile). | see `APP_VERSION` |
| `aluminum_cleaner.py` | Clean material/profile orders. Optional dual-input mode merges a "מחיר לסדרה" XLS to update quantities and recalculate weights. | see `APP_VERSION` |

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
