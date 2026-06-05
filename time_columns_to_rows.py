import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Missing dependencies: pandas / openpyxl. Please run: pip install pandas openpyxl", file=sys.stderr)
    raise

DEFAULT_TIME_COLUMN_NAME = "\u65f6\u95f4"
DEFAULT_VALUE_COLUMN_NAME = "\u503c"
DEFAULT_OUTPUT_SUFFIX = "_\u65f6\u95f4\u5217\u8f6c\u884c.xlsx"

TIME_HEADER_PATTERNS = [
    re.compile(r"^\s*\d{4}\D{1,4}\s*$"),
    re.compile(r"^\s*\d{4}\D{1,6}\d{1,2}\D{0,6}\s*$"),
    re.compile(r"^\s*\d{4}\s*[-/.]\s*\d{1,2}(\s*[-/.]\s*\d{1,2})?\s*$"),
    re.compile(r"^\s*\d{1,2}\D{1,4}\s*$"),
    re.compile(r"^\s*\d{4}\s*(q|Q)\s*[1-4]\s*$"),
    re.compile(r"^\s*(q|Q)[1-4]\s*\d{4}\s*$"),
    re.compile(r"^\s*\d{4}\s*$"),
    re.compile(r"^\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s*\d{2,4}?\s*$", re.I),
]
NON_TIME_HEADER_WORDS = {
    "\u65f6\u95f4", "\u65e5\u671f", "\u6708\u4efd", "\u5e74\u5ea6", "\u5e74\u4efd", "\u5b63\u5ea6", "\u5468", "\u661f\u671f",
    "period", "date", "month", "year", "quarter", "week", "time",
}


def normalize_header(value):
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if text.endswith(" 00:00:00"):
        text = text[:-9]
    if re.search(r"\d{4}\D{1,6}\d{1,2}", text):
        text = re.sub(r"\s+", "", text)
    return text


def looks_like_excel_serial(value):
    if isinstance(value, bool):
        return False
    try:
        number = float(value)
    except (TypeError, ValueError):
        return False
    return 20000 <= number <= 80000 and number.is_integer()


def looks_like_time_header(header):
    text = normalize_header(header)
    if not text:
        return False
    lowered = text.lower()
    if lowered in NON_TIME_HEADER_WORDS:
        return False
    if isinstance(header, (datetime, pd.Timestamp)):
        return True
    if looks_like_excel_serial(header):
        return True
    return any(pattern.match(text) for pattern in TIME_HEADER_PATTERNS)


def make_unique_columns(columns):
    seen = {}
    unique_columns = []
    display_names = {}
    for index, column in enumerate(columns):
        base_name = normalize_header(column) or f"col_{index + 1}"
        count = seen.get(base_name, 0) + 1
        seen[base_name] = count
        unique_name = base_name if count == 1 else f"{base_name}__dup{count}"
        unique_columns.append(unique_name)
        display_names[unique_name] = base_name
    return unique_columns, display_names


def prepare_columns(df):
    unique_columns, display_names = make_unique_columns(df.columns)
    df.columns = unique_columns
    return display_names


def strip_duplicate_suffix(column):
    return re.sub(r"__dup\d+$", "", normalize_header(column))


def read_raw_table(input_path, sheet_name=None):
    suffix = input_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(input_path, sheet_name=sheet_name if sheet_name else 0, header=None)
    if suffix == ".csv":
        return pd.read_csv(input_path, header=None)
    if suffix == ".tsv":
        return pd.read_csv(input_path, sep="\t", header=None)
    raise ValueError("Only .xlsx, .xls, .xlsm, .csv, .tsv files are supported.")


def row_time_score(row):
    return sum(looks_like_time_header(value) for value in row.tolist())


def build_table_from_detected_header(raw_df, min_count=2, scan_rows=30):
    raw_df = raw_df.dropna(how="all").dropna(axis=1, how="all")
    if raw_df.empty:
        raise ValueError("The table is empty.")

    best_index = None
    best_score = 0
    for row_index in raw_df.index[:scan_rows]:
        score = row_time_score(raw_df.loc[row_index])
        if score > best_score:
            best_index = row_index
            best_score = score

    if best_score < min_count:
        raise ValueError("No time columns were detected. Please check headers like 2024-01, 2024 year/month, Q1 2024.")

    header_values = raw_df.loc[best_index].tolist()
    data = raw_df.loc[raw_df.index > best_index].copy()
    data.columns = header_values
    data = data.dropna(how="all").dropna(axis=1, how="all")
    return data


def detect_time_columns(df, display_names=None, min_count=2):
    display_names = display_names or {column: column for column in df.columns}
    time_columns = [column for column in df.columns if looks_like_time_header(display_names.get(column, column))]
    if len(time_columns) >= min_count:
        return time_columns
    if time_columns:
        return time_columns
    raise ValueError("No time columns were detected.")


def write_table(df, output_path):
    suffix = output_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        df.to_excel(output_path, index=False)
    elif suffix == ".csv":
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
    elif suffix == ".tsv":
        df.to_csv(output_path, index=False, sep="\t", encoding="utf-8-sig")
    else:
        raise ValueError("Output file must be .xlsx, .csv, or .tsv.")


def transpose_time_columns(input_file, output_file=None, sheet_name=None, time_column_name=DEFAULT_TIME_COLUMN_NAME, value_column_name=DEFAULT_VALUE_COLUMN_NAME):
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    raw_df = read_raw_table(input_path, sheet_name)
    df = build_table_from_detected_header(raw_df)
    display_names = prepare_columns(df)
    time_columns = detect_time_columns(df, display_names)
    id_columns = [column for column in df.columns if column not in time_columns]

    result = df.melt(
        id_vars=id_columns,
        value_vars=time_columns,
        var_name=time_column_name,
        value_name=value_column_name,
    )
    result[time_column_name] = result[time_column_name].map(
        lambda column: normalize_header(display_names.get(column, strip_duplicate_suffix(column)))
    )

    if output_file is None:
        output_path = input_path.with_name(f"{input_path.stem}{DEFAULT_OUTPUT_SUFFIX}")
    else:
        output_path = Path(output_file)

    write_table(result, output_path)
    return output_path, [display_names.get(column, strip_duplicate_suffix(column)) for column in time_columns], len(result)


def clean_path(value):
    return value.strip().strip('"').strip("'")


def interactive_args():
    print("Time columns to rows")
    input_file = clean_path(input("Input table path: "))
    output_file = clean_path(input("Output path, press Enter for default: "))
    sheet_name = clean_path(input("Sheet name, press Enter for first sheet: "))
    return input_file, output_file or None, sheet_name or None


def main():
    if len(sys.argv) == 1:
        input_file, output_file, sheet_name = interactive_args()
        output_path, time_columns, row_count = transpose_time_columns(input_file, output_file, sheet_name)
        print(f"Detected {len(time_columns)} time columns: {', '.join(map(str, time_columns))}")
        print(f"Generated {row_count} rows: {output_path}")
        input("Done. Press Enter to exit.")
        return

    parser = argparse.ArgumentParser(description="Detect time columns in a table and turn them into rows.")
    parser.add_argument("input", help="Input table: xlsx/xls/xlsm/csv/tsv")
    parser.add_argument("-o", "--output", help="Output path")
    parser.add_argument("-s", "--sheet", help="Excel sheet name")
    parser.add_argument("--time-name", default=DEFAULT_TIME_COLUMN_NAME, help="Output time column name")
    parser.add_argument("--value-name", default=DEFAULT_VALUE_COLUMN_NAME, help="Output value column name")
    args = parser.parse_args()

    output_path, time_columns, row_count = transpose_time_columns(
        args.input,
        args.output,
        args.sheet,
        args.time_name,
        args.value_name,
    )
    print(f"Detected {len(time_columns)} time columns: {', '.join(map(str, time_columns))}")
    print(f"Generated {row_count} rows: {output_path}")


if __name__ == "__main__":
    main()
