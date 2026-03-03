# logger_setup.py
# Logging configuration for WKORPUSE runbook

import os
import logging
from datetime import datetime
import config


def setup_logging():
    """
    Set up logging with both console and file handlers.
    Creates timestamped log files in logs folder.
    """

    # Create log directory
    os.makedirs(config.LOG_DIR, exist_ok=True)

    # Generate log filename
    log_filename = config.LOG_FILE_PATTERN.format(timestamp=config.RUN_TIMESTAMP)
    log_file_path = os.path.join(config.LOG_DIR, log_filename)

    # Create logger
    logger = logging.getLogger()
    logger.setLevel(getattr(logging, config.LOG_LEVEL))

    # Remove existing handlers
    logger.handlers = []

    # Create formatters
    formatter = logging.Formatter(
        config.LOG_FORMAT,
        datefmt=config.LOG_DATE_FORMAT
    )

    # Console handler
    if config.LOG_TO_CONSOLE:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(getattr(logging, config.LOG_LEVEL))
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    # File handler
    if config.LOG_TO_FILE:
        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        file_handler.setLevel(getattr(logging, config.LOG_LEVEL))
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    # Silence noisy third-party loggers
    logging.getLogger('selenium').setLevel(logging.WARNING)
    logging.getLogger('selenium.webdriver').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    logging.getLogger('urllib3.connectionpool').setLevel(logging.WARNING)
    logging.getLogger('xlrd').setLevel(logging.WARNING)
    logging.getLogger('xlwt').setLevel(logging.WARNING)

    # Log the setup
    logger.info("=" * 60)
    logger.info("WKORPUSE Runbook - Logging initialized")
    logger.info(f"Log file: {log_file_path}")
    logger.info(f"Log level: {config.LOG_LEVEL}")
    logger.info(f"Run timestamp: {config.RUN_TIMESTAMP}")
    logger.info(f"Data date: {config.DATA_DATE_STR}")
    logger.info("=" * 60)

    return logger


if __name__ == '__main__':
    # Test logging
    logger = setup_logging()
    logger.debug("This is a debug message")
    logger.info("This is an info message")
    logger.warning("This is a warning message")
    logger.error("This is an error message")
    logger.critical("This is a critical message")
