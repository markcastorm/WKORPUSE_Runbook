#!/usr/bin/env python3
# orchestrator.py
# Main orchestrator for WKORPUSE data collection

import os
import sys
import io
from datetime import datetime

# Fix Windows console encoding for Korean characters
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import config
from logger_setup import setup_logging
from scraper import SEIBroDownloader
from parser import SEIBroParser
from file_generator import WKORPUSEFileGenerator
import logging

logger = logging.getLogger(__name__)


def print_banner():
    """Print a welcome banner"""
    print("\n" + "=" * 70)
    print(" WKORPUSE - KSD Korea Securities Depository Data Collection")
    print(" Weekdaily KOR Purchase of US Equities")
    print("=" * 70 + "\n")


def print_configuration():
    """Print current configuration"""
    print("Configuration:")
    print("-" * 70)
    print(f"  Data Date: {config.DATA_DATE_STR} ({config.DATA_DATE.strftime('%A, %B %d, %Y')})")
    print(f"  Source: {config.BASE_URL[:60]}...")
    print(f"  Output: {config.OUTPUT_DIR}")
    print(f"  Downloads: {config.DOWNLOAD_DIR}")
    print(f"  Run Timestamp: {config.RUN_TIMESTAMP}")
    print("-" * 70 + "\n")


def main():
    """Main execution flow"""

    try:
        # Setup logging
        setup_logging()

        print_banner()
        print_configuration()

        # Step 1: Download Excel from SEIBro
        print("STEP 1: Downloading data from SEIBro")
        print("=" * 70)

        downloader = SEIBroDownloader()
        download_result = downloader.download_data()

        if not download_result:
            logger.error("Download failed")
            print("\n[ERROR] Download failed. Exiting.")
            sys.exit(1)

        downloaded_file = download_result['file_path']
        data_date_str = download_result['data_date_str']

        print(f"\n[SUCCESS] Downloaded file: {os.path.basename(downloaded_file)}\n")
        logger.info(f"Downloaded: {downloaded_file}")

        # Step 2: Parse and clean the Excel file
        print("\nSTEP 2: Parsing and cleaning data")
        print("=" * 70 + "\n")

        parser = SEIBroParser()
        parsed_data = parser.parse_file(downloaded_file)

        if not parsed_data:
            logger.error("Parsing failed")
            print("\n[ERROR] Failed to parse downloaded file. Exiting.")
            sys.exit(1)

        print(f"  Rows extracted: {parsed_data['row_count']}")
        print(f"  Columns: {', '.join(parsed_data['columns'])}")
        print(f"\n[SUCCESS] Data parsed successfully\n")
        logger.info(f"Parsed {parsed_data['row_count']} rows")

        # Step 3: Generate output files
        print("\nSTEP 3: Generating output files")
        print("=" * 70 + "\n")

        generator = WKORPUSEFileGenerator()
        output_files = generator.generate_files(
            parsed_data,
            data_date_str,
            config.OUTPUT_DIR
        )

        # Step 4: Summary
        print("\n" + "=" * 70)
        print(" EXECUTION COMPLETE")
        print("=" * 70 + "\n")

        print("Summary:")
        print(f"  Data Date: {data_date_str}")
        print(f"  Rows: {output_files['row_count']}")
        print(f"  Frequency: Weekdaily")
        print()

        print("Output files:")
        print(f"  DATA (XLS): {os.path.basename(output_files['data_file'])}")
        print(f"  DATA (CSV): {os.path.basename(output_files['csv_file'])}")
        print()

        print(f"Output directory: {os.path.dirname(output_files['data_file'])}")
        print(f"Latest files: {config.LATEST_OUTPUT_DIR}")
        print()

        print("=" * 70 + "\n")

        logger.info("Orchestrator completed successfully")

        return 0

    except KeyboardInterrupt:
        print("\n\n[INTERRUPTED] Process interrupted by user")
        logger.warning("Process interrupted by user")
        sys.exit(130)

    except Exception as e:
        print(f"\n[ERROR] An unexpected error occurred: {e}")
        logger.exception("Unexpected error in orchestrator")
        sys.exit(1)


if __name__ == '__main__':
    sys.exit(main())
