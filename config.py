"""Application configuration for Oracle ERP Data Quality Monitor."""
import logging
import os

from dotenv import load_dotenv


load_dotenv()


def get_required_env(variable_name):
    """Return a required environment variable or raise an error."""
    value = os.getenv(variable_name)

    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {variable_name}"
        )

    return value

# Oracle Client configuration
ORACLE_CLIENT_PATH = get_required_env(
    "ORACLE_CLIENT_PATH"
)
# Oracle Database configuration
DB_CONFIG = {
    "account": get_required_env("ORACLE_USER"),
    "password": get_required_env("ORACLE_PASSWORD"),
    "dsn": get_required_env("ORACLE_DSN"),
    "encoding": "UTF-8",
}

# Report configuration
OUTPUT_DIRECTORY = "./output"
OUTPUT_FILENAME_TEMPLATE = (
    "DQ_PO_DISTRIBUTIONS_{date}.xlsx"
)
EXCEL_ENGINE = "openpyxl"

# Logging configuration
LOG_DIRECTORY = "./logs"
LOG_FILE = "application.log"
LOG_LEVEL = logging.INFO
LOG_FORMAT = (
    "%(asctime)s - %(levelname)s - %(message)s"
)

LOG_ROTATION_WHEN = "midnight"
LOG_ROTATION_INTERVAL = 1
LOG_BACKUP_COUNT = 14
