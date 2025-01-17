import configparser  # 設定ファイル(config.ini)を読み書きするためのモジュールをインポートする
from datetime import datetime, timedelta  # 日時と時間差を扱うクラスをインポートする

import pandas as pd  # 表形式データ(Excel/CSVなど)を簡単に操作するライブラリをインポートする
from colorama import init, Fore  # ターミナル上で文字色などを制御するライブラリをインポートする
from selenium import webdriver  # SeleniumのWebDriverを操作するためのモジュールをインポートする
from selenium.common.exceptions import NoSuchElementException  # 要素が見つからない場合の例外クラスをインポートする
from selenium.webdriver.common.action_chains import ActionChains  # マウスやキーボード操作をまとめて扱うクラスをインポートする

config = configparser.ConfigParser()  # configparserオブジェクトを生成する
config.read('config.ini')  # config.iniファイルを読み込む

TEST_LOG_EXCEL = config['PATHS']['test_log_excel']  # テスト結果を出力するExcelファイルのパスを取得する
TEST_LOG_HTML = config['PATHS']['test_log_html']  # テスト結果のHTMLレポートを出力するファイルパスを取得する

BROWSER_TYPE = config['BROWSER']['browser']  # 使用ブラウザ(chrome/firefox)を設定ファイルから取得する
IMPLICIT_WAIT = int(config['BROWSER']['implicit_wait'])  # 要素検索の暗黙的待機時間を整数にして取得する

init(autoreset=True)  # coloramaを初期化し、文字色のリセットを自動的に行う

current_date = datetime.now().strftime("%Y/%m/%d")  # 現在の日付を「YYYY/MM/DD」形式の文字列で取得する
today_date = datetime.now().day  # 今日の日(1~31)を取得する
tomorrow_date = datetime.now() + timedelta(days=1)  # 今日に1日を加算して明日の日付を取得する
date_after_7_days = datetime.now() + timedelta(days=7)  # 今日に7日を加算した日時を取得する
end_day = date_after_7_days.day  # 7日後の日(1~31)だけを取り出す
is_next_month = datetime.now().month != date_after_7_days.month  # 7日後が翌月かどうかを判定する
this_week_monday = datetime.now() - timedelta(days=datetime.now().weekday())  # 今週の月曜日を計算して取得する
current_hour = datetime.now().hour  # 現在の時刻(0~23)を整数で取得する
current_minute = datetime.now().minute  # 現在の分(0~59)を整数で取得する


def get_current_time():
    # 現在の日時を"YYYY-MM-DD HH:MM:SS"のフォーマットで文字列として返す
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class SeleniumActions:
    def __init__(self, driver, test_cases_df):
        # コンストラクタ: WebDriverとテストケースDataFrameを受け取り、ローカル変数を初期化する
        self.driver = driver
        self.local_test_cases_df = test_cases_df
        self.test_log = []
        self.step_counter = 1
        self.original_window = driver.current_window_handle  # 元のブラウザウィンドウハンドルを保持する

    def log_step(self, step_description, color, result):
        # テストステップの内容を記録し、コンソールにも出力する
        current_time = get_current_time()  # 現在時刻を取得する
        print(f"{self.step_counter}: " + Fore.YELLOW + f"[{current_time}] " + color + step_description)  # カウンタと時刻付きで出力
        self.test_log.append({
            "No": self.step_counter,        # ステップ番号
            "Time": current_time,           # 実行時刻
            "Description": step_description,# ステップ説明
            "Result": result                # 成否などの結果
        })
        self.step_counter += 1  # ステップ番号をインクリメントする

    def current_step(self, color=Fore.LIGHTGREEN_EX):
        # 現在のステップ番号に対応するテストケースの内容(case列)を取得し、成功ログを記録する
        try:
            step_description = self.local_test_cases_df.loc[self.local_test_cases_df['No'] == self.step_counter, 'case'].values[0]
            if pd.isna(step_description) or step_description == '':
                raise ValueError("ケース説明が空です。")  # ケース説明が空の場合は例外を投げる
            self.log_step(step_description, color, "成功")  # 成功時のステップログを記録する
        except (IndexError, ValueError) as e:
            self.log_step(str(e), Fore.RED, "失敗")  # 該当するステップ説明がなければ失敗として記録

    def enter_text(self, by, identifier, text):
        # 要素を取得してテキストを入力し、ログを記録する
        element = self.driver.find_element(by, identifier)  # 指定の探索方法・識別子で要素を探す
        element.send_keys(text)  # 要素に文字列を入力する
        self.current_step()      # ログにステップ実行を記録

    def click_element(self, by, identifier):
        # 要素をクリックし、ログを記録する
        element = self.driver.find_element(by, identifier)  # 要素を見つける
        element.click()                                     # 要素をクリックする
        self.current_step()                                 # ステップ実行をログに記録する

    def double_click_element(self, by, identifier):
        # 要素をダブルクリックし、ログを記録する
        element = self.driver.find_element(by, identifier)  # 要素を取得する
        actions = ActionChains(self.driver)                 # アクションチェーンを生成する
        actions.double_click(element).perform()             # ダブルクリック操作を実行
        self.current_step()                                 # ログに記録する

    def move_to_element(self, by, identifier):
        # 要素までマウスカーソルを移動し、ログを記録する
        element = self.driver.find_element(by, identifier)  # 要素を取得する
        actions = ActionChains(self.driver)                 # アクションチェーンを用意する
        actions.move_to_element(element).perform()          # 要素上へマウス移動アクションを実行
        self.current_step()                                 # ステップ実行をログに残す

    def press_key(self, by, identifier, key):
        # 要素にキー入力を行い、ログを記録する
        element = self.driver.find_element(by, identifier)  # 要素を取得する
        element.send_keys(key)                              # 引数で渡されたキーを入力
        self.current_step()                                 # ログを記録する

    def switch_to_original_tab(self):
        # 最初に取得したタブ(ウィンドウ)に切り替え、ログを記録する
        self.driver.switch_to.window(self.original_window)  # ハンドルを使って元ウィンドウに戻る
        self.current_step()                                 # ログに記録する

    def generate_html_report(self, total_steps, success_steps, fail_steps):
        # テスト完了後にHTML形式のレポートを生成する
        table_rows = ""  # テーブルの行文字列をまとめるための変数を用意する
        for step in self.test_log:
            row_class = "success" if step["Result"] == "成功" else "fail"  # 成否によってCSSクラスを変更する
            table_rows += f"""
            <tr class="{row_class}">
                <td>{step["No"]}</td>
                <td>{step["Time"]}</td>
                <td>{step["Description"]}</td>
                <td>{step["Result"]}</td>
            </tr>
            """  # HTMLテーブルの行を文字列として追記する

        if fail_steps > 0:  # 失敗ステップが1つでもあればFail表示
            test_result_class = "test-result-fail"
            test_result_text = "Fail"
            first_fail = next((item for item in self.test_log if item["Result"] != "成功"), None)  # 最初の失敗を取得
            if first_fail:
                fail_step_info = f"（No {first_fail['No']}: {first_fail['Description']} でエラー）"
            else:
                fail_step_info = ""
        else:  # 失敗がなければPass表示
            test_result_class = "test-result-pass"
            test_result_text = "Pass"
            fail_step_info = ""

        with open('report_template.html', 'r', encoding='utf-8') as f:  # テンプレートHTMLを読み込む
            template = f.read()

        html = template.format(
            report_time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),  # レポート作成時刻
            report_env=config['DEFAULT']['url'],                      # テスト対象URL
            report_browser=BROWSER_TYPE,                              # 使用ブラウザ名
            report_account=config['DEFAULT']['email'],                # 使用アカウント(email)
            total_steps=total_steps,                                  # 総ステップ数
            success_steps=success_steps,                              # 成功ステップ数
            fail_steps=fail_steps,                                    # 失敗ステップ数
            table_rows=table_rows,                                    # 上で組み立てたテーブル行
            test_result_class=test_result_class,                      # 成否に応じたCSSクラス
            test_result_text=test_result_text,                        # "Pass"または"Fail"
            fail_step_info=fail_step_info                             # 失敗ステップ情報
        )

        return html  # 完成したHTML文字列を返す

    def save_test_results(self):
        # テストログをExcelとHTMLレポートに保存する
        test_log_df = pd.DataFrame(self.test_log)      # テスト結果のリストをDataFrameに変換する
        test_log_df.to_excel(TEST_LOG_EXCEL, index=False)  # Excelに書き込む(インデックスは不要)
        print(Fore.LIGHTMAGENTA_EX + "テスト結果がExcelに保存されました。")  # 保存完了のメッセージを表示

        total_steps = len(self.test_log)  # 総ステップ数を取得する
        success_steps = sum(1 for item in self.test_log if item["Result"] == "成功")  # "成功"の数を数える
        fail_steps = sum(1 for item in self.test_log if item["Result"] != "成功")    # "成功"以外の数を数える

        html_report = self.generate_html_report(total_steps, success_steps, fail_steps)  # HTMLレポート生成
        with open(TEST_LOG_HTML, 'w', encoding='utf-8') as f:  # HTMLファイルを開く(上書きモード)
            f.write(html_report)                               # HTMLをファイルに書き込む
        print(Fore.LIGHTGREEN_EX + "HTMLレポートが生成されました。")  # レポート作成完了のメッセージ


def initialize_browser():
    # ブラウザの種類に応じてWebDriverを初期化し、暗黙的待機時間を設定する
    if BROWSER_TYPE.lower() == 'chrome':
        driver = webdriver.Chrome()
    elif BROWSER_TYPE.lower() == 'firefox':
        driver = webdriver.Firefox()
    else:
        raise ValueError("config.iniに未対応のブラウザが指定されています。")  # 想定外のブラウザ指定ならエラーを起こす

    driver.implicitly_wait(IMPLICIT_WAIT)  # 指定秒数の暗黙待機を設定する
    return driver  # 初期化したWebDriverインスタンスを返す


def execute_test_case(driver, url, email, password, actions):
    # 具体的なテストケースを実行し、例外処理も行う
    try:
        driver.get(url)        # 指定URLへアクセスする
        actions.current_step() # URLアクセスのステップをログに記録する

        driver.maximize_window()  # ブラウザを最大化する
        actions.current_step()    # 最大化のステップをログに記録する

    except NoSuchElementException as e:
        # 要素が見つからない場合に発生する例外をキャッチし、ログに失敗を記録する
        error_message = str(e).split("\n")[0]  # エラー内容の先頭行だけを取得する
        print(Fore.RED + f"テスト中に要素が見つかりませんでした: {error_message}")  # コンソールにエラーを表示
        fail_step_no = actions.step_counter  # 現在のステップ番号を取得
        try:
            fail_step_desc = actions.local_test_cases_df.loc[actions.local_test_cases_df['No'] == fail_step_no, 'case'].values[0]
        except IndexError:
            fail_step_desc = "不明なステップ"  # 一致するステップがなければ
        actions.log_step(f"「{fail_step_no} - {fail_step_desc}」実行中に対象のボタンが見つからず、テストが失敗しました。", Fore.RED, "失敗")

    except Exception as e:
        # その他の例外をキャッチしてログに記録する
        error_message = str(e)
        print(Fore.RED + f"テスト中にエラーが発生しました: {error_message}")
        actions.log_step("不明なエラーが発生し、テストが失敗しました。", Fore.RED, "失敗")

    finally:
        # エラーの有無に関わらず、テスト結果を保存する
        actions.save_test_results()


def main():
    # メイン関数: ブラウザ初期化、テストケースの読み込み、テスト実行を行う
    driver = None       # ドライバ変数の初期化
    actions = None      # SeleniumActionsのインスタンス用変数の初期化
    try:
        driver = initialize_browser()  # ブラウザを初期化して起動する
        actions = SeleniumActions(driver, pd.read_excel('test_cases.xlsx'))  # テストケースExcelを読み込んでインスタンス生成

        execute_test_case(
            driver,
            config['DEFAULT']['url'],       # テスト対象のURL
            config['DEFAULT']['email'],     # ログインなどに使うメールアドレス
            config['DEFAULT']['password'],  # パスワード(必要なら)
            actions
        )

    except Exception as e:
        # main中に起きた例外をキャッチし、ログ保存を行う
        print(Fore.RED + f"テスト中にエラーが発生しました: {str(e)}")
        if actions:
            actions.save_test_results()  # 既にactionsが初期化されていればログを保存する
    finally:
        # 処理が終わったらブラウザを閉じる
        if driver:
            driver.quit()
            print(Fore.YELLOW + "ブラウザが正常に閉じられました。")


if __name__ == "__main__":
    main()  # メイン関数を呼び出してスクリプトを開始する
