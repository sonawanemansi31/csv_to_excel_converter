import pandas as pd
import argparse
import logging
import os
import sys


def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )


def parse_rename_columns(rename_str):
    """
    Convert string like:
    "old1:new1,old2:new2"
    into dict:
    {"old1": "new1", "old2": "new2"}
    """
    rename_dict = {}
    if rename_str:
        pairs = rename_str.split(",")
        for pair in pairs:
            if ":" in pair:
                old, new = pair.split(":", 1)
                rename_dict[old.strip()] = new.strip()
    return rename_dict


def clean_dataframe(df):
    """
    Basic cleaning:
    - Strip whitespace from column names
    - Strip whitespace from string cells
    - Fill missing values
    """
    # Clean column names
    df.columns = [col.strip() for col in df.columns]

    # Strip whitespace from string values
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()

    # Replace 'nan' strings back to actual NaN
    df.replace("nan", pd.NA, inplace=True)

    # Fill missing values
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("Unknown")

    return df


def parse_dates(df):
    """
    Try to automatically parse date-like columns
    """
    for col in df.columns:
        if "date" in col.lower() or "time" in col.lower():
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce")
                logging.info(f"Parsed date column: {col}")
            except Exception as e:
                logging.warning(f"Could not parse date column '{col}': {e}")
    return df


def convert_csv_to_excel(input_file, output_file, rename_columns=None):
    try:
        # Check file exists
        if not os.path.exists(input_file):
            logging.error(f"Input file not found: {input_file}")
            sys.exit(1)

        # Check extension
        if not input_file.lower().endswith(".csv"):
            logging.error("Input file must be a CSV file.")
            sys.exit(1)

        logging.info(f"Reading CSV file: {input_file}")
        df = pd.read_csv(input_file)

        logging.info("Cleaning data...")
        df = clean_dataframe(df)

        logging.info("Parsing date columns...")
        df = parse_dates(df)

        # Rename columns if provided
        if rename_columns:
            logging.info(f"Renaming columns: {rename_columns}")
            df.rename(columns=rename_columns, inplace=True)

        # Ensure output ends with .xlsx
        if not output_file.lower().endswith(".xlsx"):
            output_file += ".xlsx"

        logging.info(f"Saving Excel file: {output_file}")
        df.to_excel(output_file, index=False, engine="openpyxl")

        logging.info("Conversion completed successfully!")

    except pd.errors.EmptyDataError:
        logging.error("The CSV file is empty.")
        sys.exit(1)

    except pd.errors.ParserError:
        logging.error("Error parsing the CSV file. Please check file format.")
        sys.exit(1)

    except PermissionError:
        logging.error("Permission denied. Close the Excel file if it is open.")
        sys.exit(1)

    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        sys.exit(1)


def main():
    setup_logging()

    parser = argparse.ArgumentParser(description="CSV to Excel Converter")
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to input CSV file"
    )
    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Path to output Excel file"
    )
    parser.add_argument(
        "-r", "--rename",
        required=False,
        help='Column renames in format "old1:new1,old2:new2"'
    )

    args = parser.parse_args()

    rename_dict = parse_rename_columns(args.rename)

    convert_csv_to_excel(
        input_file=args.input,
        output_file=args.output,
        rename_columns=rename_dict
    )


if __name__ == "__main__":
    main()