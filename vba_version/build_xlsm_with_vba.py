"""
Auto-build SupplierReliability.xlsm with VBA modules pre-imported.

One-shot script — opens Excel via COM (xlwings), starts from generated .xlsx,
imports all .bas modules from vba_modules/, saves as .xlsm. No manual Alt+F11.

PREREQUISITES (one-time setup):
1. Excel installed
2. Trust Center > Macro Settings > tick "Trust access to the VBA project object model"
3. pip install xlwings
4. Run generate_starter_xlsx.py first to create the .xlsx shell

USAGE:
    python build_xlsm_with_vba.py
"""
from __future__ import annotations

import sys
from pathlib import Path


def main() -> None:
    try:
        import xlwings as xw
    except ImportError:
        print("ERROR: xlwings not installed. Run: pip install xlwings")
        sys.exit(1)

    here = Path(__file__).resolve().parent
    xlsx_path = here / "SupplierReliability.xlsx"
    xlsm_path = here / "SupplierReliability.xlsm"
    bas_dir = here / "vba_modules"

    if not xlsx_path.exists():
        print(f"ERROR: {xlsx_path.name} not found.")
        print("Run generate_starter_xlsx.py first.")
        sys.exit(1)

    bas_files = sorted(bas_dir.glob("*.bas"))
    if not bas_files:
        print(f"ERROR: No .bas files in {bas_dir}")
        sys.exit(1)

    print(f"Opening Excel...")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False

    try:
        print(f"Loading {xlsx_path.name}...")
        wb = app.books.open(str(xlsx_path))

        # Delete xlsm if exists
        if xlsm_path.exists():
            xlsm_path.unlink()

        # SaveAs xlsm (FileFormat 52 = xlOpenXMLWorkbookMacroEnabled)
        print(f"Saving as .xlsm...")
        wb.api.SaveAs(str(xlsm_path), FileFormat=52)

        # Import each .bas module
        print(f"Importing {len(bas_files)} VBA modules...")
        for bas_file in bas_files:
            try:
                wb.api.VBProject.VBComponents.Import(str(bas_file))
                print(f"  OK  {bas_file.name}")
            except Exception as e:
                print(f"  FAIL {bas_file.name}: {e}")
                print()
                print("If error is about 'Programmatic access to VBA project not trusted':")
                print("  Excel > File > Options > Trust Center > Trust Center Settings")
                print("  > Macro Settings > tick 'Trust access to the VBA project object model'")
                print("  Then re-run this script.")
                raise

        wb.save()
        wb.close()
        print()
        print(f"[OK] Built: {xlsm_path}")
        print()
        print("Open SupplierReliability.xlsm in Excel.")
        print("Macros are ready — just add form buttons to Dashboard (per SETUP-VBA.md Buoc 4).")
    finally:
        app.quit()


if __name__ == "__main__":
    main()
