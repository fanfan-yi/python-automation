"""Configuration for Oracle EBS Account Rule Automation."""

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


EBS_LOGIN_URL = get_required_env("EBS_LOGIN_URL")
EBS_TARGET_URL = get_required_env("EBS_TARGET_URL")

AUTOMATION_INPUT_FILE = os.getenv(
    "AUTOMATION_INPUT_FILE",
    "./input/account_rules.xlsx",
)
