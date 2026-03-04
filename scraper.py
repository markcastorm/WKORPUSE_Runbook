# scraper.py
# Downloads Excel data from KSD SEIBro website

import os
import time
import logging
import glob
import random
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import config

# Setup logging
logger = logging.getLogger(__name__)


def human_delay(min_seconds=0.5, max_seconds=1.5):
    """Add a random delay to simulate human behavior"""
    delay = random.uniform(min_seconds, max_seconds)
    time.sleep(delay)


class SEIBroDownloader:
    """Downloads Excel data from KSD SEIBro website"""

    def __init__(self):
        self.driver = None
        self.download_dir = None
        self.logger = logger
        self.data_date = config.DATA_DATE

    def setup_driver(self):
        """Initialize Chrome driver with download preferences"""

        # Create download directory
        self.download_dir = os.path.abspath(config.DOWNLOAD_DIR)
        os.makedirs(self.download_dir, exist_ok=True)

        chrome_options = Options()

        # Set download directory
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        if config.HEADLESS_MODE:
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--disable-gpu')

        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--window-size=1920,1080')
        # Support Korean encoding
        chrome_options.add_argument('--lang=ko-KR')

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.set_page_load_timeout(config.WAIT_TIMEOUT)

        self.logger.info("Chrome driver initialized")
        self.logger.info(f"Download directory: {self.download_dir}")

    def navigate_to_page(self):
        """Navigate to the SEIBro data page"""

        self.logger.info(f"Navigating to {config.BASE_URL}")

        self.driver.get(config.BASE_URL)
        time.sleep(config.PAGE_LOAD_DELAY)

        self.logger.info("Page loaded successfully")

    def wait_for_element(self, selector, by=By.CSS_SELECTOR, timeout=None):
        """Wait for an element to be present and return it"""

        if timeout is None:
            timeout = config.WAIT_TIMEOUT

        try:
            wait = WebDriverWait(self.driver, timeout)
            element = wait.until(EC.presence_of_element_located((by, selector)))
            return element
        except TimeoutException:
            self.logger.error(f"Timeout waiting for element: {selector}")
            return None

    def wait_for_clickable(self, selector, by=By.CSS_SELECTOR, timeout=None):
        """Wait for an element to be clickable and return it"""

        if timeout is None:
            timeout = config.WAIT_TIMEOUT

        try:
            wait = WebDriverWait(self.driver, timeout)
            element = wait.until(EC.element_to_be_clickable((by, selector)))
            return element
        except TimeoutException:
            self.logger.error(f"Timeout waiting for clickable element: {selector}")
            return None

    def click_element_safely(self, element, description="element"):
        """Click an element with error handling"""

        try:
            # Scroll element into view
            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
            time.sleep(0.5)

            # Try regular click first
            element.click()
            self.logger.debug(f"Clicked {description}")
            return True

        except ElementClickInterceptedException:
            # Try JavaScript click as fallback
            try:
                self.driver.execute_script("arguments[0].click();", element)
                self.logger.debug(f"JS clicked {description}")
                return True
            except Exception as e:
                self.logger.error(f"Failed to click {description}: {e}")
                return False

        except Exception as e:
            self.logger.error(f"Error clicking {description}: {e}")
            return False

    def set_filter_settlement_amount(self):
        """Select '결제금액' (Settlement Amount) radio button"""

        self.logger.info("Setting filter: 결제금액 (Settlement Amount)")

        try:
            # Find and click the radio button
            radio = self.wait_for_clickable(config.SELECTORS['filter_settlement_amount'])
            if radio:
                # Add human delay before clicking
                human_delay(0.5, 1.0)
                self.click_element_safely(radio, "결제금액 radio")
                # Wait longer for sub-filter (세부구분) to appear with human-like delay
                human_delay(1.5, 2.5)
                return True
            return False
        except Exception as e:
            self.logger.error(f"Error setting settlement amount filter: {e}")
            return False

    def set_filter_net_purchase(self):
        """Select '순매수결제' (Net Purchase Settlement) radio button"""

        self.logger.info("Setting filter: 순매수결제 (Net Purchase Settlement)")

        try:
            radio = self.wait_for_clickable(config.SELECTORS['filter_net_purchase'])
            if radio:
                # Add human delay before clicking
                human_delay(0.4, 0.8)
                self.click_element_safely(radio, "순매수결제 radio")
                # Human-like delay after clicking
                human_delay(0.6, 1.2)
                return True
            return False
        except Exception as e:
            self.logger.error(f"Error setting net purchase filter: {e}")
            return False

    def set_time_period(self, period='1주'):
        """
        Select time period from dropdown

        Args:
            period: Time period option ('1주', '1개월', '3개월', '6개월', '1년')
        """

        self.logger.info(f"Setting time period: {period}")

        try:
            # Find the select dropdown
            dropdown = self.wait_for_element(config.SELECTORS['time_period_dropdown'])
            if dropdown:
                # Add human delay before interacting
                human_delay(0.5, 1.0)

                # Select the desired option using JavaScript for reliability
                self.driver.execute_script(
                    """
                    var select = arguments[0];
                    var options = select.options;
                    for(var i = 0; i < options.length; i++) {
                        if(options[i].text === arguments[1]) {
                            select.selectedIndex = i;
                            select.dispatchEvent(new Event('change', { bubbles: true }));
                            return true;
                        }
                    }
                    return false;
                    """,
                    dropdown, period
                )

                # Human-like delay after selection
                human_delay(0.7, 1.3)
                self.logger.debug(f"Time period set to: {period}")
                return True
            return False
        except Exception as e:
            self.logger.error(f"Error setting time period: {e}")
            return False

    def type_slowly(self, element, text, typing_speed_range=(0.05, 0.15)):
        """
        Type text character by character with random delays to simulate human typing

        Args:
            element: The input element to type into
            text: The text to type
            typing_speed_range: Tuple of (min, max) seconds between keystrokes
        """
        for char in text:
            element.send_keys(char)
            time.sleep(random.uniform(*typing_speed_range))

    def set_date_range(self):
        """Set the date range (same date for start and end) - simulates human typing"""

        date_str = self.data_date.strftime('%Y/%m/%d')
        self.logger.info(f"Setting date range: {date_str}")

        try:
            # Set start date - simulate human typing
            start_input = self.wait_for_clickable(config.SELECTORS['date_start'])
            if start_input:
                # Click to focus
                self.click_element_safely(start_input, "start date input")
                human_delay(0.3, 0.6)

                # Clear existing value
                start_input.send_keys(Keys.CONTROL + "a")
                human_delay(0.1, 0.3)
                start_input.send_keys(Keys.BACKSPACE)
                human_delay(0.2, 0.4)

                # Type the date slowly like a human
                self.type_slowly(start_input, date_str, typing_speed_range=(0.08, 0.18))
                human_delay(0.3, 0.6)

                # Press Tab or Enter to confirm
                start_input.send_keys(Keys.TAB)
                human_delay(0.5, 0.9)

                self.logger.debug(f"Start date set to: {date_str}")

            # Set end date - simulate human typing
            end_input = self.wait_for_clickable(config.SELECTORS['date_end'])
            if end_input:
                # Click to focus
                self.click_element_safely(end_input, "end date input")
                human_delay(0.3, 0.6)

                # Clear existing value
                end_input.send_keys(Keys.CONTROL + "a")
                human_delay(0.1, 0.3)
                end_input.send_keys(Keys.BACKSPACE)
                human_delay(0.2, 0.4)

                # Type the date slowly like a human
                self.type_slowly(end_input, date_str, typing_speed_range=(0.08, 0.18))
                human_delay(0.3, 0.6)

                # Press Tab or Enter to confirm
                end_input.send_keys(Keys.TAB)
                human_delay(0.5, 0.9)

                self.logger.debug(f"End date set to: {date_str}")

            self.logger.info(f"Date range set to: {date_str}")
            return True

        except Exception as e:
            self.logger.error(f"Error setting date range: {e}")
            return False

    def set_country_usa(self):
        """Select '미국' (USA) country filter"""

        self.logger.info("Setting filter: 미국 (USA)")

        try:
            radio = self.wait_for_clickable(config.SELECTORS['country_usa'])
            if radio:
                # Add human delay before clicking
                human_delay(0.4, 0.9)
                self.click_element_safely(radio, "미국 radio")
                # Human-like delay after clicking
                human_delay(0.6, 1.1)
                return True
            return False
        except Exception as e:
            self.logger.error(f"Error setting USA filter: {e}")
            return False

    def click_search(self):
        """Click the search button (조회) - with human-like delay"""

        self.logger.info("Clicking search button (조회)")

        try:
            # Add significant human delay before clicking search
            # Simulate user reviewing the form before clicking search
            self.logger.debug("Waiting before search (simulating human review)...")
            human_delay(1.5, 2.5)

            # Try clicking the search image first
            search_btn = self.wait_for_clickable(config.SELECTORS['search_button'])
            if search_btn:
                self.click_element_safely(search_btn, "조회 button (image)")
                # Wait for table to update with human-like delay
                human_delay(config.PAGE_LOAD_DELAY, config.PAGE_LOAD_DELAY + 1.0)
                self.logger.info("Search executed, waiting for results...")
                return True

            # Fallback: try clicking the parent anchor
            self.logger.info("Trying fallback: parent anchor")
            search_anchor = self.wait_for_clickable(config.SELECTORS['search_button_parent'])
            if search_anchor:
                self.click_element_safely(search_anchor, "조회 button (anchor)")
                human_delay(config.PAGE_LOAD_DELAY, config.PAGE_LOAD_DELAY + 1.0)
                self.logger.info("Search executed via anchor, waiting for results...")
                return True

            # Fallback: execute JavaScript directly
            self.logger.info("Trying fallback: JavaScript getPList")
            self.driver.execute_script("getPList(1);")
            human_delay(config.PAGE_LOAD_DELAY, config.PAGE_LOAD_DELAY + 1.0)
            self.logger.info("Search executed via JS, waiting for results...")
            return True

        except Exception as e:
            self.logger.error(f"Error clicking search: {e}")
            return False

    def wait_for_data_table(self, timeout=15):
        """Wait for data table (grid2) to load with data"""

        self.logger.info("Waiting for data table to load...")

        try:
            # Wait for grid2 (the visible table) to have rows
            start_time = time.time()
            while time.time() - start_time < timeout:
                try:
                    # grid2 is the visible data table, grid1 is hidden
                    rows = self.driver.find_elements(By.CSS_SELECTOR, "div#grid2 tbody tr")
                    if rows and len(rows) > 0:
                        self.logger.info(f"Data table loaded with {len(rows)} rows")
                        return True
                except:
                    pass
                time.sleep(1)

            self.logger.warning("Timeout waiting for data table")
            return False

        except Exception as e:
            self.logger.error(f"Error waiting for data table: {e}")
            return False

    def dismiss_alerts(self):
        """Dismiss any alert dialogs that might be blocking"""

        try:
            from selenium.webdriver.common.alert import Alert
            alert = Alert(self.driver)
            alert_text = alert.text
            self.logger.info(f"Found alert: {alert_text}")
            alert.accept()
            time.sleep(0.5)
            return True
        except:
            return False

    def extract_table_data(self):
        """
        Extract data directly from the visible data table (grid2) on the page.
        Saves the table as an HTML file for the parser to process.

        Returns:
            Path to saved HTML file, or None if failed
        """

        self.logger.info("Extracting data from page table (grid2)...")

        try:
            # Get the table HTML from grid2 (the visible data table)
            # grid1 is hidden (display:none), grid2 is the actual visible table
            table_html = self.driver.execute_script("""
                var table = document.querySelector('table#grid2_body_table');
                if (!table) {
                    table = document.querySelector('div#grid2 table');
                }
                return table ? table.outerHTML : null;
            """)

            if not table_html:
                self.logger.error("No data table found in grid2")
                return None

            # Count rows for logging
            row_count = table_html.count('<tr id="row13"')
            self.logger.info(f"Extracted table with {row_count} data rows")

            # Wrap in basic HTML structure
            html_content = f"""<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
{table_html}"""

            # Save to download directory as .xls (HTML format)
            output_filename = "주요국 외화주식 예탁결제현황.xls"
            output_path = os.path.join(self.download_dir, output_filename)

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            self.logger.info(f"Table data extracted and saved: {output_filename}")
            return output_path

        except Exception as e:
            self.logger.error(f"Error extracting table data: {e}")
            return None

    def download_data(self):
        """
        Main method to download data from SEIBro.
        Returns path to downloaded file.
        """

        try:
            self.setup_driver()
            self.navigate_to_page()

            # Set all filters
            print("\n" + "=" * 60)
            print(f"Configuring filters for date: {config.DATA_DATE_STR}")
            print("=" * 60 + "\n")

            # Step 1: Set 결제금액 filter
            if not self.set_filter_settlement_amount():
                self.logger.warning("Could not set settlement amount filter")

            # Step 2: Set 순매수결제 filter
            if not self.set_filter_net_purchase():
                self.logger.warning("Could not set net purchase filter")

            # Step 3: Set time period dropdown (1주 = 1 week, suitable for single-day queries)
            if not self.set_time_period('1주'):
                self.logger.warning("Could not set time period - continuing anyway")

            # Step 4: Set date range
            if not self.set_date_range():
                self.logger.error("Failed to set date range")
                return None

            # Step 5: Set USA country filter
            if not self.set_country_usa():
                self.logger.warning("Could not set USA filter")

            # Step 6: Click search
            if not self.click_search():
                self.logger.error("Failed to execute search")
                return None

            # Step 7: Wait for data table to load
            print("\nWaiting for data to load...")
            self.wait_for_data_table(timeout=15)

            # Additional wait for page to stabilize with human-like behavior
            human_delay(2.0, 3.5)

            # Dismiss any alerts that might block download
            self.dismiss_alerts()

            # Step 8: Extract data from table
            print("\n" + "=" * 60)
            print("Extracting data from table...")
            print("=" * 60 + "\n")

            downloaded_file = self.extract_table_data()

            if downloaded_file:
                # Use ascii-safe filename for console output (Korean chars can't print on Windows cp1252)
                try:
                    print(f"[SUCCESS] Downloaded: {os.path.basename(downloaded_file)}")
                except UnicodeEncodeError:
                    print(f"[SUCCESS] Downloaded: {os.path.basename(downloaded_file).encode('ascii', 'replace').decode()}")
                self.logger.info(f"Successfully downloaded: {downloaded_file}")

                return {
                    'file_path': downloaded_file,
                    'data_date': self.data_date,
                    'data_date_str': config.DATA_DATE_STR
                }
            else:
                print("[FAILED] Download failed")
                self.logger.error("Download failed")
                return None

        except Exception as e:
            self.logger.exception(f"Error in download_data: {e}")
            return None

        finally:
            if self.driver:
                self.driver.quit()
                self.logger.info("Browser closed")


def main():
    """Test the downloader"""
    from logger_setup import setup_logging
    setup_logging()

    print("\n" + "=" * 60)
    print(" WKORPUSE Data Downloader Test")
    print("=" * 60 + "\n")

    print(f"Data date: {config.DATA_DATE_STR}")
    print(f"Download directory: {config.DOWNLOAD_DIR}")
    print()

    downloader = SEIBroDownloader()
    result = downloader.download_data()

    if result:
        print("\n" + "=" * 60)
        print(" Download Result")
        print("=" * 60)
        print(f"  File: {result['file_path']}")
        print(f"  Data Date: {result['data_date_str']}")
        print("=" * 60 + "\n")
    else:
        print("\n[ERROR] Download failed\n")


if __name__ == '__main__':
    main()
