import os    #切換工作目錄
import time    #等待與計算執行時間

import pandas as pd    #讀 Excel
import pyautogui    #目前登入操作
import pyperclip    #處理「會計子目」

from selenium import webdriver    #Chrome automation
from selenium.webdriver.common.by import By    #尋找網頁元素
from selenium.webdriver.common.keys import Keys    
from selenium.webdriver.support.ui import Select

from config import (
    AUTOMATION_INPUT_FILE,
    EBS_LOGIN_URL,
    EBS_TARGET_URL,
)

def read_my_login():
    check_excel_data = pd.read_excel('C:/Users/yvonne/Desktop/login.xlsx') #開啟excel
    row_list = [] #寫在row list
    for idx in range(len(check_excel_data.index)): #抓excel index row 數量
        a = check_excel_data.iloc[idx]
        row_list.append(a.values)
    return row_list

def read_my_xls():

    df = pd.read_excel(AUTOMATION_INPUT_FILE)
    return df.to_dict("index"), len(df.index)

def login_web():
    time.sleep(2)
    load_excel = read_my_login()
    for loging in load_excel:

        Login_name, Pssword= loging
        pyautogui.write(Login_name)
        print("輸入帳號:****")
        pyautogui.press('tab')
        pyautogui.write(str(Pssword))
        pyautogui.press('enter')
        print("輸入密碼:****")

def login_step(row):

    time.sleep(2)
    start = driver.find_element(By.ID,'CreateButton')#點擊建立
    time.sleep(2)
    start.click()

    time.sleep(2)
    your_name =driver.find_element(By.XPATH,'//*[@id="AccountRuleValue"]')#科目規則值	
    your_select =driver.find_element(By.XPATH,'//*[@id="SegmentName"]')#節段名稱
    unit = driver.find_element(By.XPATH,'//*[@id="SegmentValue"]')#節段值
    
    name1 = row["name1"]
    select = row["select"]
    unitt = row["unitt"]

    time.sleep(2)
    your_name.send_keys(name1)

    time.sleep(2)
    
    select1 = Select(your_select)
    select1.select_by_visible_text(select)

    time.sleep(3)
    if select == '會計子目':
        pyperclip.copy("00")  # put text in clipboard
        unit.send_keys(Keys.CONTROL, "v")
        blank = driver.find_element(By.XPATH,'//*[@id="p_SwanPageLayout"]/div[2]/table/tbody/tr[1]/td[2]')
        blank.click()
        time.sleep(5)
        driver.switch_to.frame('iframelovPopUp_SegmentValue')
        time.sleep(5)
        quick = driver.find_element(By.XPATH,'//*[@id="SegmentValueRN:Content"]/tbody/tr[1]/td[2]')

        quick.click()
        
        time.sleep(2)
        driver.switch_to.default_content()
    else:
        unit.send_keys(unitt)

    finish = driver.find_element(By.ID,'ApplyButton')#點擊套用
    finish.click()
    time.sleep(2)

if __name__ == '__main__':
    driver = webdriver.Chrome()
    start_t =time.time()
    
    print("連網站")
    driver.get(EBS_LOGIN_URL)  # 連網站
    print("#################登入完成#################")
    
    driver.get(EBS_TARGET_URL)
    login_web()
    excel, len_of_excel = read_my_xls()
    for idx in range(len_of_excel):
        print("第", idx+1, "筆")
        current_row = excel[idx]
        login_step(current_row)
    print("完成")
    end_t = time.time()
    print("執行時間:%f 秒" % (end_t - start_t))

    print("================= 第", idx+1, "筆: 結束 ====================")
    print("=============== all end! =================")
