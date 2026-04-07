import os
import glob
import xlrd
import pandas as pd

FOLDER     = "acil servis forecasting"
OUTPUT_CSV = "patient_counts.csv"
DATE_COL   = "GELIS_ZAMANI"


def excel_date_to_datetime(excel_val):
    return xlrd.xldate_as_datetime(excel_val, 0)


def read_sheet(sh):
    """Read a sheet into a DataFrame. Returns None if it lacks DATE_COL."""
    if sh.nrows < 2:
        return None
    headers = [sh.cell_value(0, c) for c in range(sh.ncols)]
    if DATE_COL not in headers:
        return None
    rows = [
        [sh.cell_value(r, c) for c in range(sh.ncols)]
        for r in range(1, sh.nrows)
    ]
    return pd.DataFrame(rows, columns=headers)


def read_file(path):
    wb = xlrd.open_workbook(path)
    frames = []
    for sheet_name in wb.sheet_names():
        sh = wb.sheet_by_name(sheet_name)
        df = read_sheet(sh)
        if df is not None:
            df["_sheet"] = sheet_name
            frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)

    # Deduplicate across sheets: two rows are the same patient record
    # when every data column (excluding the helper _sheet column) is identical.
    data_cols = [c for c in combined.columns if c != "_sheet"]
    before = len(combined)
    combined = combined.drop_duplicates(subset=data_cols)
    after = len(combined)
    if before != after:
        print(f"    Removed {before - after} duplicate rows (cross-sheet)")

    # Convert Excel float dates → Python date
    dates = []
    for val in combined[DATE_COL]:
        if val:
            try:
                dates.append(excel_date_to_datetime(float(val)).date())
            except Exception:
                pass

    return pd.DataFrame({"date": dates})


def main():
    pattern = os.path.join(FOLDER, "*.xls")
    files = sorted(glob.glob(pattern))
    print(f"Found {len(files)} files.")

    all_records = []
    for path in files:
        print(f"  Reading: {os.path.basename(path)}")
        df = read_file(path)
        all_records.append(df)

    combined = pd.concat(all_records, ignore_index=True)
    combined["date"] = pd.to_datetime(combined["date"])

    daily = (
        combined.groupby("date")
        .size()
        .reset_index(name="num_patients")
        .sort_values("date")
    )

    daily.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(daily)} rows to '{OUTPUT_CSV}'")
    print(daily.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
