import os    #切換工作目錄
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

    df = pd.read_excel(AUTOMATION_INPUT_FILE)
    return df.to_dict("index"), len(df.index)

def create_account_rule(row):
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

if __name__ == '__main__':
    driver = webdriver.Chrome()
    start_t =time.time()
    
    print("連網站")
    driver.get(EBS_LOGIN_URL)  # 連網站
    input("Log in to Oracle EBS, then press Enter to continue...")
    
    driver.get(EBS_TARGET_URL)

    excel, len_of_excel = read_my_xls()
    for idx in range(len_of_excel):
        print("第", idx+1, "筆")
        current_row = excel[idx]
        create_account_rule(current_row)
    print("完成")
    end_t = time.time()
    print("執行時間:%f 秒" % (end_t - start_t))

    print("================= 第", idx+1, "筆: 結束 ====================")
    print("=============== all end! =================")
