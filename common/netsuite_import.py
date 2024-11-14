import re
from playwright.sync_api import Playwright, sync_playwright, expect
from dotenv import load_dotenv
import os

load_dotenv(".env")
NETSUITE_PASSWORD = os.getenv("NETSUITE_PASSWORD")

add_auth_question_1 = "初めての仕事はどの市町村でしましたか?"
ADD_AUTH_ANSWER_1 = os.getenv("ADD_AUTH_ANSWER_1")
add_auth_question_2 = "6年生のとき、どの学校に通いましたか?"
ADD_AUTH_ANSWER_2 = os.getenv("ADD_AUTH_ANSWER_2")
add_auth_question_3 = "一番年上の兄弟姉妹は何年何月に誕生しましたか? (例:1900年1月)"
ADD_AUTH_ANSWER_3 = os.getenv("ADD_AUTH_ANSWER_3")

def accept_alert(dialog):
    dialog.accept()

def run(playwright: Playwright, file_path) -> None:
    # ブラウザを立ち上げる
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    # インポートする際アラートが表示されるので、それをOKする処理を定義しておく
    page.on('dialog', accept_alert)

    # netsuiteにアクセス
    page.goto("https://4136251.app.netsuite.com/")
    page.get_by_placeholder("Email address").click()
    page.get_by_placeholder("Email address").fill("shinya.arai@uzabase.com")
    page.get_by_placeholder("Password").click()
    page.get_by_placeholder("Password").fill(NETSUITE_PASSWORD)
    page.get_by_role("button", name="Log In").click()

    # 追加認証の質問内容を取得
    question = page.locator(".smalltextnolink.text-opensans").nth(2).text_content()

    # 質問内容によって条件分岐させてログインする
    if(question == add_auth_question_1):
        page.locator("#null").fill(ADD_AUTH_ANSWER_1)
    elif(question == add_auth_question_2):
        page.locator("#null").fill(ADD_AUTH_ANSWER_2)
    elif(question == add_auth_question_3):
        page.locator("#null").fill(ADD_AUTH_ANSWER_3)
    else:
        print(question)

    page.get_by_role("button", name="送信").click()
    # ログイン完了

    # インポートする
    page.get_by_role("link", name="保存したCSVインポート").click()
    page.get_by_role("link", name="Production_仕訳Import").click()
    page.locator("#fileupload0").set_input_files(file_path)
    page.get_by_role("button", name="次へ >").click()
    page.wait_for_timeout(2000) 
    page.get_by_role("button", name="次へ >").click()
    page.wait_for_timeout(2000) 
    page.get_by_role("button", name="次へ >").click()
    page.wait_for_timeout(2000)
    page.get_by_role("cell", name="名前をつけて保存&実行", exact=True).get_by_role("img").click()
    page.get_by_role("link", name="実行").click()
    page.wait_for_timeout(3000) # アラートが表示されるので念のため3秒待機

    # インポート完了

    # ---------------------
    context.close()
    browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright,"")
