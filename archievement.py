import subprocess
import sys

packages = ['webdriver_manager','gspread','oauth2client','numpy']
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def upgrade(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])

# Install/upgrade packages
install('selenium==4.1.0')
for package in packages:
    try:
        dist = __import__(package)
        print("{} ({}) is already installed".format(package, dist.__version__))
        upgrade(package)
        print("{} is upgraded".format(package))
    except ImportError:
        print("{} is NOT installed".format(package))
        install(package)
        print("{} is installed".format(package))


import gspread
from oauth2client.service_account import ServiceAccountCredentials
from oauth2client.service_account import ServiceAccountCredentials
import numpy

from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = Options()

try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    print("WebDriver has been successfully initiated.")
except Exception as e:
    print("An error occurred while initiating WebDriver:", e)

wait = WebDriverWait(driver,10)
wait_all_elements = wait.until(EC.presence_of_all_elements_located)

# OAuth2クライアントIDのJSONファイルへのパス
json_keyfile_path = r"C:\Users\yuton\Downloads\court-booking-automation-cf53a34c6969.json"

# 使用するAPI
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# 認証情報の設定
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)

# 認証情報を使用してGoogle Sheets APIに接続
gc = gspread.authorize(credentials)

# スプレッドシートのURLは省略
spreadsheet = gc.open_by_url('target_sheet')

df1 = spreadsheet.worksheet('コートカード')
df2 = spreadsheet.worksheet('予約コート')

df1_values = df1.get_all_values()
df2_values = df2.get_all_values()

print(df1_values)
print(df2_values)
print(len(df1_values))
print(len(df2_values))

blocks_length = len(df2_values)
id_sheet_length = len(df1_values)

booking_blocks = []
for i in range(12):
    base_column_number = 1 + 4*i
    #booking_block_iの[0][0]が予約コート、[1][0]が時間、[2][0~9]が予約する日、[3][0~9]が使用コートカード
    if not df2_values[4][base_column_number] == '':
        booking_block_i = []
        target_court = []
        target_time = []
        target_days = []
        user_name_for_id = []
        if not df2_values[4][base_column_number] == '':
            target_court.append(df2_values[4][base_column_number])
        if not df2_values[6][base_column_number] == '':
            target_time.append(df2_values[6][base_column_number])

        for x in range(blocks_length-4):
            if not df2_values[x+4][base_column_number+1] == '':
                target_days.append(str(df2_values[x+4][base_column_number+1]))
            if not df2_values[x+4][base_column_number+2] == '':
                user_name_for_id.append(df2_values[x+4][base_column_number+2])
        booking_block_i.append(target_court)
        booking_block_i.append(target_time)
        booking_block_i.append(target_days)
        booking_block_i.append(user_name_for_id)
        booking_blocks.append(booking_block_i)
print(booking_blocks)

#booking_blocksで未入力があるかを調べる。blank_check = 'there_is_blanks'が定義されたとき、警告文を出す。
booking_blocks_length = len(booking_blocks)
for i in range(booking_blocks_length):
    booking_block_i_length = len(booking_blocks[i])
    for j in range(booking_block_i_length):
        if len(booking_blocks[i][j]) == 0:
            blank_check = 'there_is_blanks'

import tkinter as tk
from tkinter import messagebox

if 'blank_check' in globals():
    root = tk.Tk()
    root.withdraw()  # tkinterのメインウィンドウを表示しない
    messagebox.showerror("エラー", "Excelの入力にミスがあります。必要な全ての欄に情報が入力されていることを確認してください。")

wrong_id_or_password = []
cannot_book = []

if not 'blank_check' in globals():
    root = tk.Tk()
    root.withdraw()  # tkinterのメインウィンドウを表示しない
    driver.get('https://g-kyoto.pref.kyoto.lg.jp/reserve_j/core_i/init.asp?SBT=1')
    wait_all_elements
    response = messagebox.askokcancel("確認", "予約ブロック数は"+str(len(booking_blocks))+"個ですか？\n正しく無い場合、「予約するコート」が入力されていない場合があります。\n正しい場合は「OK」を選択してください。")
    if response:
        comfirm = messagebox.askokcancel("確認","予約を実行しますか？\n実行する場合、excelファイルが閉じていることを確認してください。")
        if comfirm:
            messagebox.showinfo("実行","抽選予約を実行します。")
            #予約を実行する
            time.sleep(5)
            driver.switch_to.frame('MainFrame')
            for i in range(len(booking_blocks)):
                days_list = booking_blocks[i][2]
                park_name = booking_blocks[i][0][0]
                target_time = booking_blocks[i][1][0]
                print(days_list)
                print(park_name)
                print(target_time)
                for j in range(len(booking_blocks[i][3])):
                    #user_nameをbooking_blocksから取得して、「コートカード」で行を検索。
                    user_name = booking_blocks[i][3][j]
                    arr = numpy.array(df1_values)
                    indices = numpy.where(arr == user_name)
                    print(indices)
                    row = indices[0][0]
                    col = indices[1][0]

                    print(f"[{row}][{col}]")
                    user_id = df1_values[row][col+1]
                    user_password = df1_values[row][col+2]
                    print(user_id + user_password)
                    time.sleep(5)
                    driver.find_element_by_link_text('テニス').click()
                    wait_all_elements
                    driver.find_element_by_link_text('京都市左京区').click()
                    wait_all_elements
                    driver.find_element_by_name('btn_next').click()
                    wait_all_elements
                    if park_name == '岡崎':
                        driver.find_element_by_xpath('/html/body/form/div[2]/center/table[4]/tbody/tr/td/table/tbody/tr[1]/td[3]/input').click()
                    elif park_name == '宝':
                        driver.find_element_by_xpath('/html/body/form/div[2]/center/table[4]/tbody/tr/td/table/tbody/tr[3]/td[7]/input').click()
                    wait_all_elements
                    driver.find_element_by_xpath('/html/body/form/div[2]/div[1]/center/table[4]/tbody/tr[1]/th[3]/a').click()
                    for day in days_list:
                        driver.find_element_by_link_text(day).click()
                        time_xpath = [driver.find_element_by_xpath('/html/body/form/div[2]/div[2]/left/table[3]/tbody/tr/td/table/tbody/tr[5]/td[1]/a/img'),driver.find_element_by_xpath('/html/body/form/div[2]/div[2]/left/table[3]/tbody/tr/td/table/tbody/tr[5]/td[5]/a/img'),driver.find_element_by_xpath('/html/body/form/div[2]/div[2]/left/table[3]/tbody/tr/td/table/tbody/tr[5]/td[6]/a/img')]
                        if target_time == '8~10':
                            time_xpath[0].click()
                        elif target_time == '4~6':
                            time_xpath[1].click()
                        elif target_time == '6~9':
                            time_xpath[2].click()
                        wait_all_elements

                        original_window = driver.current_window_handle
                        wait_all_elements
                        new_window_handle = driver.window_handles[-1]
                        wait_all_elements
                        driver.switch_to.window(new_window_handle)
                        wait_all_elements
                        # `<select>`要素を特定
                        select_element = Select(driver.find_element_by_tag_name('select'))
                        # オプションを選択。ここでは値（value）によって選択します。
                        wait_all_elements
                        select_element.select_by_value('1')
                        wait_all_elements
                        driver.find_element_by_xpath('/html/body/form/p/input[1]').click()
                        wait_all_elements
                        driver.switch_to.window(original_window)
                        wait_all_elements
                        driver.switch_to.frame('MainFrame')
                        wait_all_elements
                        driver.find_element_by_css_selector('input.clsImage[name="btn_ok"]').click()
                        wait_all_elements

                        try:
                            blank_of_user_id = driver.find_element_by_xpath('/html/body/form/div[2]/center/table/tbody/tr[1]/td/input')
                            blank_of_user_id.send_keys(user_id)
                            driver.find_element_by_xpath('/html/body/form/div[2]/center/table/tbody/tr[2]/td/input').send_keys(user_password)
                            driver.find_element_by_css_selector('input.clsImage[name="btn_ok"]').click()

                        except NoSuchElementException:
                            pass
                        wait_all_elements
                        try:
                            driver.find_element_by_name("btn_next").click()
                        except:
                            messagebox.showinfo("エラー",f"{user_name}さんのコートカードは停止されている、予約制限を超えている、ID・パスワードが間違っている可能性があります。\n手動でログインして原因を確認後、スプレッドシートを修正してください。")
                        wait_all_elements
                        driver.find_element_by_name("btn_next").click()
                        wait_all_elements
                        driver.find_element_by_name("btn_cmd").click()
                        wait_all_elements
                        alert = Alert(driver)
                        alert.accept()
                        wait_all_elements
                        driver.find_element_by_name("btn_back").click()
                        try:
                            driver.find_element_by_class_name('ResultMsg')
                            messagebox.showinfo("エラー",f"{user_name}さんのコートカードではすでに同一コマが予約されている、もしくはカードの有効期限が切れている可能性があります。\n手動でログインして原因を確認後、スプレッドシートを修正してください。")
                        except:
                            pass

                    
                    #抽選予約確認のスクショを撮る
                    wait_all_elements
                    driver.find_element_by_name("btn_MoveMenu").click()
                    wait_all_elements
                    driver.find_element_by_xpath("//*[@value='抽選予約確認']").click()
                    wait_all_elements
                    driver.find_element_by_name("cmdSearch").click()

                    # JavaScriptを使ってページの最下部にスクロールします。
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

                    # スクリーンショットを撮り、保存します。
                    wait_all_elements
                    driver.save_screenshot(f'C:\\Users\\yuton\\抽選予約確認\\screenshot({user_name}).png')
                    time.sleep(3)
                    logout_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.NAME, 'btn_LogOut'))
                    )
                    logout_button.click()

        else:
            messagebox.showinfo("キャンセル","実行がキャンセルされました。")
    else:
        messagebox.showinfo("キャンセル","実行がキャンセルされました。")

print(wrong_id_or_password)
print(cannot_book)