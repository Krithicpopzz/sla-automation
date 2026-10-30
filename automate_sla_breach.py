from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

# ---------------------------------------------------------
# CONFIGURATION: update these file paths if needed
# ---------------------------------------------------------
IMEI_FILE = Path("IMEI.xlsx")
RESPONSE_FILE = Path("Response.xlsx")
RESOLUTION_FILE = Path("Resolution.xlsx")
OUTPUT_FILE = Path("Completed_IMEI.xlsx")

# Exact column headers expected in the workbooks
NUMBER_COL = "Number"
SOURCE_BREACHED_COL = "Has breached"
RESOLUTION_TARGET_COL = "Has Breached Resolution SLA"
RESPONSE_TARGET_COL = "Has Breached Response SLA"


def normalize_number(value):
    """Normalize common Excel number representations for matching."""
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def read_source_mapping(file_path, sheet_name=0):
    """Read Number -> Has Breached mapping from a source workbook."""
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    required = {NUMBER_COL, SOURCE_BREACHED_COL}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"{file_path.name} is missing required column(s): {', '.join(sorted(missing))}"
        )

    df = df[[NUMBER_COL, SOURCE_BREACHED_COL]].copy()

    # Normalize identifiers and ignore rows without a Number.
    df["_match_number"] = df[NUMBER_COL].map(normalize_number)
    df = df.dropna(subset=["_match_number"]).copy()

    duplicates = df.loc[
        df["_match_number"].duplicated(keep=False), "_match_number"
    ].unique()
    if len(duplicates):
        sample = ", ".join(duplicates[:10])
        raise ValueError(
            f"{file_path.name} has duplicate Number values. Examples: {sample}"
        )

    # Accept actual booleans or text TRUE/FALSE; reject unexpected values.
    def parse_boolean(value, row_number):
        if pd.isna(value):
            raise ValueError(
                f"{file_path.name}: blank '{SOURCE_BREACHED_COL}' value near Excel row {row_number}."
            )
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)) and value in (0, 1):
            return bool(value)
        text = str(value).strip().casefold()
        if text == "true":
            return True
        if text == "false":
            return False
        raise ValueError(
            f"{file_path.name}: unexpected value {value!r} in '{SOURCE_BREACHED_COL}' "
            f"near Excel row {row_number}; expected TRUE or FALSE."
        )

    parsed_values = []
    for i, value in enumerate(df[SOURCE_BREACHED_COL].tolist(), start=2):
        parsed_values.append(parse_boolean(value, i))

    return dict(zip(df["_match_number"], parsed_values))


def main():
    for file_path in (IMEI_FILE, RESPONSE_FILE, RESOLUTION_FILE):
        if not file_path.exists():
            raise FileNotFoundError(
                f"Could not find {file_path}. Put all three input files in the same folder as this script "
                "or update the file paths in the configuration."
            )

    # Read the main file to validate headers and prepare matching keys.
    main_df = pd.read_excel(IMEI_FILE, engine="openpyxl")
    if NUMBER_COL not in main_df.columns:
        raise ValueError(f"{IMEI_FILE.name} is missing the '{NUMBER_COL}' column.")

    for target_col in (RESOLUTION_TARGET_COL, RESPONSE_TARGET_COL):
        if target_col not in main_df.columns:
            raise ValueError(
                f"{IMEI_FILE.name} is missing the target column '{target_col}'. "
                "This script expects both target columns to already exist."
            )

    main_keys = main_df[NUMBER_COL].map(normalize_number)
    if main_keys.isna().any():
        raise ValueError(f"{IMEI_FILE.name} contains blank values in '{NUMBER_COL}'.")

    if main_keys.duplicated().any():
        duplicate_examples = main_keys[main_keys.duplicated(keep=False)].unique()[:10]
        raise ValueError(
            f"{IMEI_FILE.name} has duplicate Number values, although they are expected to be unique. "
            f"Examples: {', '.join(duplicate_examples)}"
        )

    response_map = read_source_mapping(RESPONSE_FILE)
    resolution_map = read_source_mapping(RESOLUTION_FILE)



    # Load the original workbook and update only the two target columns,
    # preserving the other cells and the workbook's existing layout.
    workbook = load_workbook(IMEI_FILE)
    worksheet = workbook.worksheets[0]  # first worksheet

    # Find headers in row 1.
    headers = {
        str(cell.value).strip(): cell.column
        for cell in worksheet[1]
        if cell.value is not None
    }

    for required_col in (
        NUMBER_COL,
        RESOLUTION_TARGET_COL,
        RESPONSE_TARGET_COL,
    ):
        if required_col not in headers:
            raise ValueError(
                f"Could not find '{required_col}' in the first row of the active worksheet "
                f"of {IMEI_FILE.name}."
            )

    number_column = headers[NUMBER_COL]
    resolution_column = headers[RESOLUTION_TARGET_COL]
    response_column = headers[RESPONSE_TARGET_COL]

    updated_rows = 0
    for row in range(2, worksheet.max_row + 1):
        number_value = normalize_number(worksheet.cell(row, number_column).value)
        if number_value is None:
            continue

        # Write actual Boolean values, not the text strings "TRUE"/"FALSE".
        worksheet.cell(row, resolution_column).value = resolution_map.get(
            number_value, "NOT FOUND"
        )
        worksheet.cell(row, response_column).value = response_map.get(
            number_value, "NOT FOUND"
        )
        updated_rows += 1

    workbook.save(OUTPUT_FILE)
    print(f"Success! Updated {updated_rows} rows.")
    print(f"Output file: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
