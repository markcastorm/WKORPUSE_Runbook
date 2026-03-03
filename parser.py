# parser.py
# Parse and clean downloaded Excel data from SEIBro

import os
import logging
import pandas as pd
import xlrd
import config

logger = logging.getLogger(__name__)


class SEIBroParser:
    """Parses and cleans Excel data from SEIBro website"""

    def __init__(self):
        self.debug = config.DEBUG_MODE
        self.logger = logger

    def is_html_file(self, file_path):
        """
        Check if file is actually HTML (common with Korean govt sites)

        Args:
            file_path: Path to the file

        Returns:
            True if HTML, False otherwise
        """
        try:
            with open(file_path, 'rb') as f:
                # Read first 1000 bytes to check for HTML markers
                header = f.read(1000)

                # Check for common HTML markers
                html_markers = [
                    b'<html', b'<HTML',
                    b'<!DOCTYPE', b'<!doctype',
                    b'<meta', b'<META',
                    b'<table', b'<TABLE'
                ]

                for marker in html_markers:
                    if marker in header:
                        self.logger.info("File detected as HTML format")
                        return True

                return False

        except Exception as e:
            self.logger.error(f"Error checking file format: {e}")
            return False

    def read_html_table(self, file_path):
        """
        Read HTML table from file (common format for Korean govt websites)

        Args:
            file_path: Path to the HTML file

        Returns:
            pandas DataFrame or None if failed
        """
        self.logger.info("Reading HTML table from file...")

        try:
            # Read HTML file with proper encoding
            with open(file_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # Parse all tables in the HTML
            tables = pd.read_html(html_content, encoding='utf-8')

            if not tables:
                self.logger.error("No tables found in HTML")
                return None

            # Usually the data table is the largest one
            largest_table = max(tables, key=lambda x: len(x))

            self.logger.info(f"Found {len(tables)} table(s), using largest one")
            self.logger.info(f"Shape: {largest_table.shape[0]} rows x {largest_table.shape[1]} columns")

            return largest_table

        except UnicodeDecodeError:
            # Try with different encodings common in Korean files
            try:
                self.logger.info("Trying EUC-KR encoding...")
                with open(file_path, 'r', encoding='euc-kr') as f:
                    html_content = f.read()

                tables = pd.read_html(html_content)
                if tables:
                    largest_table = max(tables, key=lambda x: len(x))
                    return largest_table

            except Exception as e:
                self.logger.error(f"Failed with EUC-KR encoding: {e}")

        except Exception as e:
            self.logger.error(f"Error reading HTML table: {e}")

        return None

    def read_excel_file(self, file_path):
        """
        Read the downloaded Excel file.
        Handles Korean encoding, various Excel formats, and HTML tables.

        Args:
            file_path: Path to the Excel file

        Returns:
            pandas DataFrame or None if failed
        """

        self.logger.info(f"Reading file: {file_path}")

        if not os.path.exists(file_path):
            self.logger.error(f"File not found: {file_path}")
            return None

        # First check if file is actually HTML (common with Korean govt sites)
        if self.is_html_file(file_path):
            self.logger.info("File is HTML format, using HTML parser")
            df = self.read_html_table(file_path)

            if df is not None:
                self.logger.info(f"Successfully read HTML file")
                self.logger.info(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
                self.logger.debug(f"Columns: {list(df.columns)}")
                return df
            else:
                self.logger.error("Failed to read HTML table")
                return None

        # If not HTML, try reading as Excel
        try:
            # Try reading with pandas (handles most Excel formats)
            # Use xlrd for .xls files (older Excel format)
            df = pd.read_excel(file_path, engine='xlrd')

            self.logger.info(f"Successfully read Excel file")
            self.logger.info(f"Shape: {df.shape[0]} rows x {df.shape[1]} columns")
            self.logger.debug(f"Columns: {list(df.columns)}")

            return df

        except Exception as e:
            self.logger.error(f"Error reading Excel file: {e}")

            # Try alternative reading methods
            try:
                self.logger.info("Trying alternative read method...")
                df = pd.read_excel(file_path)
                return df
            except Exception as e2:
                self.logger.error(f"Alternative read also failed: {e2}")
                return None

    def validate_columns(self, df):
        """
        Validate that the DataFrame has expected columns.

        Args:
            df: pandas DataFrame

        Returns:
            True if valid, False otherwise
        """

        if df is None or df.empty:
            self.logger.error("DataFrame is empty or None")
            return False

        # Expected Korean column names
        expected_columns = [col['korean'] for col in config.OUTPUT_COLUMNS]

        # Check if columns match (allowing for slight variations)
        actual_columns = list(df.columns)

        self.logger.debug(f"Expected columns: {expected_columns}")
        self.logger.debug(f"Actual columns: {actual_columns}")

        # Check column count
        if len(actual_columns) != len(expected_columns):
            self.logger.warning(
                f"Column count mismatch: expected {len(expected_columns)}, got {len(actual_columns)}"
            )

        # Check for expected column names
        missing_columns = []
        for expected in expected_columns:
            if expected not in actual_columns:
                missing_columns.append(expected)

        if missing_columns:
            self.logger.warning(f"Missing columns: {missing_columns}")
            if config.REQUIRE_ALL_COLUMNS:
                return False

        return True

    def clean_numeric_value(self, value):
        """
        Clean and convert a value to numeric format.
        Removes commas, spaces, and handles Korean number formatting.

        Args:
            value: Raw value from Excel

        Returns:
            Cleaned numeric value or original if not numeric
        """

        if pd.isna(value):
            return None

        if isinstance(value, (int, float)):
            return value

        # Convert to string and clean
        value_str = str(value).strip()

        # Remove commas and spaces (common in Korean number formatting)
        value_str = value_str.replace(',', '').replace(' ', '').replace('\u3000', '')

        try:
            # Try to convert to int first (most values are integers)
            if '.' in value_str:
                return float(value_str)
            else:
                return int(value_str)
        except ValueError:
            # Return original string if not numeric
            return value

    def clean_string_value(self, value):
        """
        Clean string values.
        Removes extra whitespace and normalizes encoding.

        Args:
            value: Raw string value

        Returns:
            Cleaned string
        """

        if pd.isna(value):
            return ""

        # Convert to string and clean
        value_str = str(value).strip()

        # Normalize whitespace
        value_str = ' '.join(value_str.split())

        return value_str

    def clean_dataframe(self, df):
        """
        Clean the entire DataFrame.
        Removes formatting issues, normalizes data types.

        Args:
            df: Raw pandas DataFrame

        Returns:
            Cleaned pandas DataFrame
        """

        self.logger.info("Cleaning DataFrame...")

        if df is None or df.empty:
            return None

        # Create a copy to avoid modifying original
        cleaned_df = df.copy()

        # Process each column based on expected type
        for col_info in config.OUTPUT_COLUMNS:
            korean_name = col_info['korean']
            col_type = col_info['type']

            if korean_name not in cleaned_df.columns:
                continue

            if col_type == 'integer':
                cleaned_df[korean_name] = cleaned_df[korean_name].apply(
                    lambda x: self.clean_numeric_value(x)
                )
            elif col_type == 'numeric':
                cleaned_df[korean_name] = cleaned_df[korean_name].apply(
                    lambda x: self.clean_numeric_value(x)
                )
            elif col_type == 'string':
                cleaned_df[korean_name] = cleaned_df[korean_name].apply(
                    lambda x: self.clean_string_value(x)
                )

        # Remove any completely empty rows
        cleaned_df = cleaned_df.dropna(how='all')

        # Reset index
        cleaned_df = cleaned_df.reset_index(drop=True)

        self.logger.info(f"Cleaned DataFrame: {cleaned_df.shape[0]} rows")

        return cleaned_df

    def validate_data(self, df):
        """
        Validate the cleaned data.

        Args:
            df: Cleaned pandas DataFrame

        Returns:
            True if valid, False otherwise
        """

        if df is None or df.empty:
            self.logger.error("DataFrame is empty after cleaning")
            return False

        # Check minimum rows
        if len(df) < config.MIN_EXPECTED_ROWS:
            self.logger.warning(
                f"Low row count: {len(df)} (expected at least {config.MIN_EXPECTED_ROWS})"
            )

        # Check for any NaN values in critical columns
        critical_columns = ['순위', '종목코드', '종목명']
        for col in critical_columns:
            if col in df.columns:
                null_count = df[col].isna().sum()
                if null_count > 0:
                    self.logger.warning(f"Column '{col}' has {null_count} null values")

        # Validate rank column is sequential
        if '순위' in df.columns:
            ranks = df['순위'].dropna().astype(int).tolist()
            expected_ranks = list(range(1, len(ranks) + 1))
            if ranks != expected_ranks:
                self.logger.warning("Rank column is not sequential")

        self.logger.info("Data validation completed")
        return True

    def parse_file(self, file_path):
        """
        Main method to parse and clean an Excel file.

        Args:
            file_path: Path to the downloaded Excel file

        Returns:
            Dict with cleaned data and metadata, or None if failed
        """

        self.logger.info(f"\nParsing file: {file_path}")

        # Read the Excel file
        df = self.read_excel_file(file_path)

        if df is None:
            self.logger.error("Failed to read Excel file")
            return None

        # Validate columns
        if not self.validate_columns(df):
            self.logger.error("Column validation failed")
            return None

        # Clean the data
        cleaned_df = self.clean_dataframe(df)

        if cleaned_df is None:
            self.logger.error("Failed to clean data")
            return None

        # Validate the cleaned data
        self.validate_data(cleaned_df)

        # Build result dictionary
        result = {
            'data': cleaned_df,
            'row_count': len(cleaned_df),
            'column_count': len(cleaned_df.columns),
            'columns': list(cleaned_df.columns),
            'source_file': file_path
        }

        self.logger.info(f"Successfully parsed file: {result['row_count']} rows")

        return result


def main():
    """Test the parser with a sample file"""
    import sys
    from logger_setup import setup_logging

    setup_logging()

    # Check for command line argument
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Use sample file from project_information
        file_path = os.path.join(
            os.path.dirname(__file__),
            'project_information',
            'ÁÖ¿ä±¹ ¿ÜÈ_ÁÖ½Ä ¿¹Å¹°áÁ¦ÇöÈ².xls'
        )

    print("\n" + "=" * 60)
    print(" WKORPUSE Parser Test")
    print("=" * 60 + "\n")

    if not os.path.exists(file_path):
        print(f"[ERROR] File not found: {file_path}")
        sys.exit(1)

    print(f"Input file: {file_path}")
    print()

    parser = SEIBroParser()
    result = parser.parse_file(file_path)

    if result:
        print("\n" + "=" * 60)
        print(" Parse Result")
        print("=" * 60)
        print(f"  Rows: {result['row_count']}")
        print(f"  Columns: {result['column_count']}")
        print(f"  Column names: {result['columns']}")
        print()
        print("First 5 rows:")
        print(result['data'].head().to_string())
        print("=" * 60 + "\n")
    else:
        print("\n[ERROR] Parse failed\n")


if __name__ == '__main__':
    main()
