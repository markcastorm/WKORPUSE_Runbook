# config.py
# WKORPUSE - KSD Korea Securities Depository Data Collection Configuration

import os
from datetime import datetime, timedelta

# =============================================================================
# DATA SOURCE CONFIGURATION
# =============================================================================

BASE_URL = 'https://seibro.or.kr/websquare/control.jsp?w2xPath=/IPORTAL/user/ovsSec/BIP_CNTS10013V.xml&menuNo=921'
PROVIDER_NAME = 'KSD - Korea Securities Depository'
DATASET_NAME = 'WKORPUSE'
DATASET_DESCRIPTION = 'Weekdaily KOR Purchase of US Equities'
COUNTRY = 'South Korea'
CURRENCY = 'USD'

# =============================================================================
# TIMESTAMPED FOLDERS CONFIGURATION
# =============================================================================

# Generate timestamp for this run (format: YYYYMMDD_HHMMSS)
RUN_TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')

# Use timestamped folders to avoid conflicts between runs
USE_TIMESTAMPED_FOLDERS = True

# =============================================================================
# DATE CONFIGURATION
# =============================================================================

def get_data_date(reference_date=None):
    """
    Calculate the data date based on weekday logic.
    - No data on weekends
    - If today is Monday, get Friday's data
    - Otherwise, get previous day's data

    Args:
        reference_date: Date to calculate from (defaults to today)

    Returns:
        datetime object for the data date
    """
    if reference_date is None:
        reference_date = datetime.now()

    weekday = reference_date.weekday()  # Monday=0, Sunday=6

    if weekday == 0:  # Monday - get Friday's data
        data_date = reference_date - timedelta(days=3)
    elif weekday == 6:  # Sunday - get Friday's data
        data_date = reference_date - timedelta(days=2)
    elif weekday == 5:  # Saturday - get Friday's data
        data_date = reference_date - timedelta(days=1)
    else:  # Tuesday-Friday - get previous day's data
        data_date = reference_date - timedelta(days=1)

    return data_date

# Target date for data extraction (None = auto-calculate based on today)
TARGET_DATE = None  # Options: None (auto), datetime(2026, 1, 30), etc.

# Get the actual data date
if TARGET_DATE is None:
    DATA_DATE = get_data_date()
else:
    DATA_DATE = TARGET_DATE

# Format data date for various uses
DATA_DATE_STR = DATA_DATE.strftime('%Y%m%d')  # For filename: 20260130
DATA_DATE_DISPLAY = DATA_DATE.strftime('%Y/%m/%d')  # For web form: 2026/01/30
DATA_DATE_KOREAN = DATA_DATE.strftime('%Y/%m/%d')  # Korean date format

# =============================================================================
# WEB SCRAPING SELECTORS
# =============================================================================

SELECTORS = {
    # Filter radio buttons - 현황구분 (Status Classification)
    'filter_settlement_amount': 'input#a1_radio1_input_0',  # 결제금액 radio
    'filter_custody_amount': 'input#a1_radio1_input_1',  # 보관금액 radio

    # Detail filter - 세부구분 (only visible when 결제금액 is selected)
    'filter_buy_settlement': 'input#area_radio_2_input_0',  # 매수결제
    'filter_sell_settlement': 'input#area_radio_2_input_1',  # 매도결제
    'filter_buy_sell_settlement': 'input#area_radio_2_input_2',  # 매수+매도결제
    'filter_net_purchase': 'input#area_radio_2_input_3',  # 순매수결제 radio

    # Date input fields
    'time_period_dropdown': 'select#sd1_selectbox1_input_0',  # Time period dropdown (1주/1개월/3개월/6개월/1년)
    'date_start': 'input#sd1_inputCalendar1_input',  # Start date input
    'date_end': 'input#sd1_inputCalendar2_input',  # End date input

    # Country filter radio buttons - 국가구분
    'country_usa': 'input#area_radio_input_1',  # 미국 (USA)

    # Search button - the image inside an anchor with getPList
    'search_button': 'img#image2',  # 조회 button image
    'search_button_parent': 'a[href*="getPList"]',  # Parent anchor with JS function

    # Excel download button
    'excel_download': 'a#ExcelDownload_a',  # 엑셀다운로드 button

    # Data table (grid)
    'data_table': 'div#grid1',
    'table_body': 'tbody',
    'table_rows': 'tr',
}

# =============================================================================
# OUTPUT COLUMN STRUCTURE
# =============================================================================

# Column mapping from Korean to internal codes
OUTPUT_COLUMNS = [
    {
        'korean': '순위',
        'english': 'Rank',
        'code': 'RANK',
        'type': 'integer'
    },
    {
        'korean': '국가',
        'english': 'Country',
        'code': 'COUNTRY',
        'type': 'string'
    },
    {
        'korean': '종목코드',
        'english': 'StockCode',
        'code': 'STOCK_CODE',
        'type': 'string'
    },
    {
        'korean': '종목명',
        'english': 'StockName',
        'code': 'STOCK_NAME',
        'type': 'string'
    },
    {
        'korean': '매수결제',
        'english': 'BuySettlement',
        'code': 'BUY_SETTLEMENT',
        'type': 'numeric'
    },
    {
        'korean': '매도결제',
        'english': 'SellSettlement',
        'code': 'SELL_SETTLEMENT',
        'type': 'numeric'
    },
    {
        'korean': '순매수결제',
        'english': 'NetPurchase',
        'code': 'NET_PURCHASE',
        'type': 'numeric'
    }
]

# =============================================================================
# BROWSER CONFIGURATION
# =============================================================================

HEADLESS_MODE = False  # Set to True for production
DEBUG_MODE = True
WAIT_TIMEOUT = 30  # Increased for slower Korean website
PAGE_LOAD_DELAY = 5  # Wait for dynamic content
DOWNLOAD_WAIT_TIME = 30  # Wait time for Excel download (increased for slower sites)

# =============================================================================
# OUTPUT CONFIGURATION
# =============================================================================

# Base directories
BASE_DOWNLOAD_DIR = './downloads'
BASE_OUTPUT_DIR = './output'
BASE_LOG_DIR = './logs'

# Apply timestamping if enabled
if USE_TIMESTAMPED_FOLDERS:
    DOWNLOAD_DIR = os.path.join(BASE_DOWNLOAD_DIR, RUN_TIMESTAMP)
    OUTPUT_DIR = os.path.join(BASE_OUTPUT_DIR, RUN_TIMESTAMP)
    LOG_DIR = os.path.join(BASE_LOG_DIR, RUN_TIMESTAMP)
else:
    DOWNLOAD_DIR = BASE_DOWNLOAD_DIR
    OUTPUT_DIR = BASE_OUTPUT_DIR
    LOG_DIR = BASE_LOG_DIR

# Latest folder (always contains most recent extraction)
LATEST_OUTPUT_DIR = os.path.join(BASE_OUTPUT_DIR, 'latest')

# File naming patterns
# Output: WKORPUSE_DATA_YYYYMMDD.xls (date is data date, not run date)
DATA_FILE_PATTERN = 'WKORPUSE_DATA_{date}.xls'

# Log file naming
LOG_FILE_PATTERN = 'wkorpuse_{timestamp}.log'

# =============================================================================
# LOGGING CONFIGURATION
# =============================================================================

# Log levels: DEBUG, INFO, WARNING, ERROR, CRITICAL
LOG_LEVEL = 'DEBUG' if DEBUG_MODE else 'INFO'

# Log format
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Console output
LOG_TO_CONSOLE = True
LOG_TO_FILE = True

# =============================================================================
# DATA CLEANING CONFIGURATION
# =============================================================================

# Expected number of columns in source data
EXPECTED_COLUMN_COUNT = 7

# Encoding for reading Korean Excel files
SOURCE_ENCODING = 'utf-8'
OUTPUT_ENCODING = 'utf-8'

# =============================================================================
# VALIDATION SETTINGS
# =============================================================================

# Minimum expected rows in data (sanity check)
MIN_EXPECTED_ROWS = 10

# Validate that all required columns are present
REQUIRE_ALL_COLUMNS = True

# =============================================================================
# ERROR HANDLING
# =============================================================================

# Maximum retries for download failures
MAX_DOWNLOAD_RETRIES = 3
RETRY_DELAY = 2.0  # Seconds between retries

# Continue processing even if some steps fail
CONTINUE_ON_ERROR = False
