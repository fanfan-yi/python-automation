import time    #等待與計算執行時間

import pandas as pd    #讀 Excel
import pyperclip    #處理「會計子目」

from selenium import webdriver    #Chrome automation
from selenium.webdriver.common.by import By    #尋找網頁元素
from selenium.webdriver.common.keys import Keys    
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import (
    AUTOMATION_INPUT_FILE,
    EBS_LOGIN_URL,
    EBS_TARGET_URL,
)

WAIT_TIMEOUT = 15

def read_my_xls():
    """Read and validate the account rule input file."""

    df = pd.read_excel(AUTOMATION_INPUT_FILE)

    required_columns = {
        "name1",
        "select",
        "unitt",
    }

    missing_columns = required_columns - set(df.columns)
    
    if missing_columns:    #空值驗證
        raise ValueError(
            "Missing required columns: "
            + ", ".join(sorted(missing_columns))
        )
        
    required_value_columns = ["name1", "select"]
    
    empty_required_rows = df[
        df[required_value_columns].isna().any(axis=1)
    ]

    if not empty_required_rows.empty:
        row_numbers = (empty_required_rows.index + 2).tolist()
        raise ValueError(
            f"Missing required values in Excel rows: {row_numbers}"
        )

    missing_segment_value_rows = df[
        (df["select"] != "會計子目")
        & df["unitt"].isna()
    ]

    if not missing_segment_value_rows.empty:
        row_numbers = (missing_segment_value_rows.index + 2).tolist()
        raise ValueError(
            f"Missing segment value in Excel rows: {row_numbers}"
        )
    return df.to_dict("index"), len(df.index)

def create_account_rule(driver, row):
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    create_button = wait.until(
        EC.element_to_be_clickable((By.ID, "CreateButton"))
    )
    create_button.click()  # 點擊建立

    account_rule_value = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "AccountRuleValue")
        )
    )  # 科目規則值

    segment_name = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "SegmentName")
        )
    )  # 節段名稱

    segment_value = wait.until(
        EC.visibility_of_element_located(
            (By.ID, "SegmentValue")
        )
    )  # 節段值

    name1 = row["name1"]
    select = row["select"]
    unitt = row["unitt"]

    account_rule_value.send_keys(name1)

    segment_select = Select(segment_name)
    segment_select.select_by_visible_text(select)

    if select == "會計子目":
        pyperclip.copy("00")
        segment_value.send_keys(Keys.CONTROL, "v")

        blank = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="p_SwanPageLayout"]/div[2]/table/'
                    'tbody/tr[1]/td[2]',
                )
            )
        )
        blank.click()

        wait.until(
            EC.frame_to_be_available_and_switch_to_it(
                (By.ID, "iframelovPopUp_SegmentValue")
            )
        )

        quick_select = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    '//*[@id="SegmentValueRN:Content"]/tbody/'
                    'tr[1]/td[2]',
                )
            )
        )
        quick_select.click()

        driver.switch_to.default_content()

    else:
        segment_value.send_keys(unitt)

    apply_button = wait.until(
        EC.element_to_be_clickable((By.ID, "ApplyButton"))
    )
    apply_button.click()

def recover_to_account_rule_page(driver):
    """Return to the account rule page after a failed record."""

    driver.switch_to.default_content()
    driver.get(EBS_TARGET_URL)

    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    wait.until(
        EC.element_to_be_clickable(
            (By.ID, "CreateButton")
        )
    )

if __name__ == "__main__":
    driver = webdriver.Chrome()
    start_t = time.time()

    try:
        print("Connecting to Oracle EBS...")

        driver.get(EBS_LOGIN_URL)
        input(
            "Log in to Oracle EBS, then press Enter to continue..."
        )

        driver.get(EBS_TARGET_URL)

        excel, len_of_excel = read_my_xls()

        failed_rows = []
        successful_rows = 0

        for idx in range(len_of_excel):
            print(f"Processing row {idx + 1}...")

            current_row = excel[idx]

            try:
                create_account_rule(driver, current_row)
                successful_rows += 1

            except Exception as error:
                failed_rows.append(
                    {
                        "row": idx + 1,
                        "error": str(error),
                    }
                )

                print(
                    f"Row {idx + 1} failed: {error}"
                )

                try:
                    recover_to_account_rule_page(driver)
                
                except Exception as recovery_error:
                    print(
                        "Unable to recover the Oracle EBS page: "
                        f"{recovery_error}"
                    )
                    raise

        end_t = time.time()

        print("\n===== Execution Summary =====")
        print(f"Total rows: {len_of_excel}")
        print(f"Successful: {successful_rows}")
        print(f"Failed: {len(failed_rows)}")
        print(
            f"Execution time: "
            f"{end_t - start_t:.2f} seconds"
        )

        if failed_rows:
            print("\nFailed rows:")

            for failure in failed_rows:
                print(
                    f"- Row {failure['row']}: "
                    f"{failure['error']}"
                )

    finally:
        driver.quit()
