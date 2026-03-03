# file_generator.py
# Generate clean Excel DATA file for WKORPUSE dataset

import os
import logging
import shutil
import xlwt
import pandas as pd
import config

logger = logging.getLogger(__name__)


class WKORPUSEFileGenerator:
    """Generates clean Excel DATA files in the required format"""

    def __init__(self):
        self.debug = config.DEBUG_MODE
        self.logger = logger

    def create_data_file(self, df, data_date_str, output_path):
        """
        Create the DATA Excel file without borders or formatting issues.

        Args:
            df: Cleaned pandas DataFrame
            data_date_str: Date string for the data (YYYYMMDD format)
            output_path: Path to save the Excel file

        Returns:
            Path to created file
        """

        self.logger.info(f"Creating DATA file: {output_path}")

        try:
            # Create workbook and sheet
            workbook = xlwt.Workbook(encoding='utf-8')
            sheet = workbook.add_sheet('DATA')

            # Create styles (no borders)
            # Header style
            header_style = xlwt.XFStyle()
            header_font = xlwt.Font()
            header_font.bold = False  # No bold to match sample
            header_style.font = header_font

            # Number style (no borders, plain)
            number_style = xlwt.XFStyle()
            number_style.num_format_str = '#,##0'  # Number with comma separator

            # Text style (no borders, plain)
            text_style = xlwt.XFStyle()

            # Write header row
            for col_idx, column_name in enumerate(df.columns):
                sheet.write(0, col_idx, column_name, header_style)

            # Write data rows
            for row_idx, row in df.iterrows():
                for col_idx, column_name in enumerate(df.columns):
                    value = row[column_name]

                    # Determine style based on column type
                    col_info = next(
                        (c for c in config.OUTPUT_COLUMNS if c['korean'] == column_name),
                        None
                    )

                    if col_info and col_info['type'] in ['integer', 'numeric']:
                        # Write as number
                        if pd.notna(value):
                            sheet.write(row_idx + 1, col_idx, value, number_style)
                    else:
                        # Write as text
                        if pd.notna(value):
                            sheet.write(row_idx + 1, col_idx, str(value), text_style)

            # Set column widths for better readability
            col_widths = {
                0: 2000,   # 순위 (Rank)
                1: 3000,   # 국가 (Country)
                2: 5000,   # 종목코드 (Stock Code)
                3: 15000,  # 종목명 (Stock Name)
                4: 5000,   # 매수결제 (Buy Settlement)
                5: 5000,   # 매도결제 (Sell Settlement)
                6: 5000,   # 순매수결제 (Net Purchase)
            }

            for col_idx, width in col_widths.items():
                if col_idx < len(df.columns):
                    sheet.col(col_idx).width = width

            # Save the file
            workbook.save(output_path)
            self.logger.info(f"DATA file saved: {output_path}")

            return output_path

        except Exception as e:
            self.logger.error(f"Error creating DATA file: {e}")
            raise

    def create_csv_file(self, df, output_path):
        """
        Create a CSV version of the data file (UTF-8 encoding).

        Args:
            df: Cleaned pandas DataFrame
            output_path: Path to save the CSV file

        Returns:
            Path to created file
        """

        self.logger.info(f"Creating CSV file: {output_path}")

        try:
            # Save as CSV with UTF-8 BOM for Excel compatibility
            df.to_csv(output_path, index=False, encoding='utf-8-sig')

            self.logger.info(f"CSV file saved: {output_path}")
            return output_path

        except Exception as e:
            self.logger.error(f"Error creating CSV file: {e}")
            raise

    def generate_files(self, parsed_data, data_date_str, output_dir):
        """
        Generate output files from parsed data.

        Args:
            parsed_data: Dict with 'data' key containing pandas DataFrame
            data_date_str: Date string for the data (YYYYMMDD format)
            output_dir: Directory to save output files

        Returns:
            Dict with paths to created files
        """

        # Create output directory
        os.makedirs(output_dir, exist_ok=True)

        df = parsed_data['data']

        # Generate filename with data date
        data_filename = config.DATA_FILE_PATTERN.format(date=data_date_str)
        data_path = os.path.join(output_dir, data_filename)

        # Also create CSV version
        csv_filename = data_filename.replace('.xls', '.csv')
        csv_path = os.path.join(output_dir, csv_filename)

        # Create the files
        self.create_data_file(df, data_date_str, data_path)
        self.create_csv_file(df, csv_path)

        # Copy to 'latest' folder
        latest_dir = config.LATEST_OUTPUT_DIR
        os.makedirs(latest_dir, exist_ok=True)

        latest_data_path = os.path.join(latest_dir, f"WKORPUSE_DATA_latest.xls")
        latest_csv_path = os.path.join(latest_dir, f"WKORPUSE_DATA_latest.csv")

        shutil.copy2(data_path, latest_data_path)
        shutil.copy2(csv_path, latest_csv_path)

        self.logger.info("Files also copied to 'latest' folder")

        return {
            'data_file': data_path,
            'csv_file': csv_path,
            'latest_data': latest_data_path,
            'latest_csv': latest_csv_path,
            'data_date': data_date_str,
            'row_count': len(df)
        }


def main():
    """Test the file generator with sample data"""
    import sys
    from logger_setup import setup_logging

    setup_logging()

    print("\n" + "=" * 60)
    print(" WKORPUSE File Generator Test")
    print("=" * 60 + "\n")

    # Create sample data matching the expected structure
    sample_data = {
        '순위': [1, 2, 3, 4, 5],
        '국가': ['미국', '미국', '미국', '미국', '미국'],
        '종목코드': ['US5951121038', 'US25460G2865', 'US88160R1014', 'US02079K3059', 'US78463V1070'],
        '종목명': ['MICRON TECHNOLOGY INC', 'DIREXION DAILY TSLA BULL 2X SHARES', 'TESLA INC', 'ALPHABET INC CL A', 'SPDR GOLD SHARES ETF'],
        '매수결제': [44060511, 45323354, 71492119, 42327673, 24291902],
        '매도결제': [13903519, 15369582, 41980340, 13464514, 2611786],
        '순매수결제': [30156992, 29953772, 29511778, 28863159, 21680117]
    }

    df = pd.DataFrame(sample_data)

    parsed_data = {
        'data': df,
        'row_count': len(df)
    }

    generator = WKORPUSEFileGenerator()
    result = generator.generate_files(parsed_data, '20260130', config.OUTPUT_DIR)

    print("\n" + "=" * 60)
    print(" Generated Files")
    print("=" * 60)
    for key, value in result.items():
        print(f"  {key}: {value}")
    print("=" * 60 + "\n")


if __name__ == '__main__':
    main()
