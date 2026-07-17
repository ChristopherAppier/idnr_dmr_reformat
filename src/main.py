import pandas as pd
from pathlib import Path
import calendar
from tkinter import Tk, filedialog

def is_missing(value):
    if pd.isna(value):
        return True
    text = str(value).strip().lower()
    return text in {'', 'nan', 'none'}

def main():
    # Prompt the user to choose the folder containing DMR files
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    selected_folder = filedialog.askdirectory(
        title='Select folder containing DMR files',
        initialdir=str(Path(__file__).parent)
    )
    root.destroy()

    if not selected_folder:
        print('No folder selected. Exiting.')
        return

    folder_path = Path(selected_folder)

    # Creating a csv to hold the data scraped from the DMRs (overwrites any existing file)
    output_file = folder_path / 'IDNR_DMR_Data.csv'
    all_data = []

    month_lookup = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}
    month_lookup.update({name.lower(): i for i, name in enumerate(calendar.month_abbr) if name})

    # Import each DMR workbook and process all tabs except Cover
    for file in list(folder_path.glob('*.xlsx')) + list(folder_path.glob('*.xlsm')):
        workbook = pd.ExcelFile(file)
        cover_sheet_name = next(
            (sheet for sheet in workbook.sheet_names if sheet.strip().lower() == 'cover'),
            None
        )
        if cover_sheet_name is None:
            print(f"Skipping file with no Cover sheet: {file.name}")
            continue

        # Read month/year from Cover if available; otherwise leave DATE blank for this file
        date_data = None
        try:
            cover_df = workbook.parse(sheet_name=cover_sheet_name, header=None)
            month_raw = cover_df.iloc[12, 2]
            year_raw = cover_df.iloc[12, 3]

            if is_missing(month_raw) or is_missing(year_raw):
                print(f"Could not identify date in file: {file}")
            else:
                month_text = str(month_raw).strip().lower()
                if month_text in month_lookup:
                    month_num = month_lookup[month_text]
                else:
                    month_num = int(float(month_raw))

                year_num = int(float(year_raw))
                days_in_month = calendar.monthrange(year_num, month_num)[1]
                date_data = pd.Series(
                    [f"{year_num}/{month_num:02d}/{day:02d}" for day in range(1, days_in_month + 1)],
                    name='DATE | (YYYY/MM/DD)'
                )
        except Exception:
            print(f"Could not identify date in file: {file}")

        for sheet_name in workbook.sheet_names:
            if sheet_name.strip().lower() == 'cover':
                continue

            try:
                # Read each non-cover tab using the same structure as the current effluent parsing
                df = pd.read_excel(file, sheet_name=sheet_name, header=[9, 10, 11]).iloc[6:37, 1:]
                df = df.dropna(axis=1, how='all').reset_index(drop=True)
                if date_data is None:
                    row_count = len(df)
                    date_slice = pd.Series([''] * row_count, name='DATE | (YYYY/MM/DD)')
                else:
                    row_count = min(len(date_data), len(df))
                    date_slice = date_data.iloc[:row_count].reset_index(drop=True)

                df = df.iloc[:row_count].reset_index(drop=True)

                # Flatten Excel multi-row headers into a single readable line per column
                df.columns = [
                    ' | '.join(str(part).strip() for part in col if pd.notna(part) and str(part).strip())
                    for col in df.columns.to_flat_index()
                ]

                # Add tab name next to DATE so each row keeps its source sheet
                tab_data = pd.Series([sheet_name] * row_count, name='SOURCE')
                combined_data = pd.concat([date_slice, tab_data, df], axis=1)
                all_data.append(combined_data)
            except Exception:
                print(f"Skipping sheet in file: {file.name} [{sheet_name}]")

    # Write once so headers are always at the top and all discovered columns are preserved
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True, sort=False)
        first_columns = ['DATE | (YYYY/MM/DD)', 'SOURCE']
        other_columns = [col for col in final_df.columns if col not in first_columns]
        if other_columns:
            other_df = final_df[other_columns].replace(r'^\s*$', pd.NA, regex=True)
            keep_other_columns = [col for col in other_columns if not other_df[col].isna().all()]
        else:
            keep_other_columns = []

        final_df = final_df[first_columns + keep_other_columns]
        final_df.to_csv(output_file, index=False, header=True)
    else:
        pd.DataFrame(columns=['DATE | (YYYY/MM/DD)', 'SOURCE']).to_csv(output_file, index=False, header=True)

    return

if __name__ == "__main__":
    main()