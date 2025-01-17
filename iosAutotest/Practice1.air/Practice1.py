#!C:\Python\Python312\python.exe
# -*- encoding=utf8 -*-
__author__ = "徐"

import datetime  # 日付や時刻を扱うモジュール
import json      # JSONファイルを扱うモジュール
import logging   # ロギングを行うモジュール
import shutil    # ファイルやディレクトリ操作を簡易化するモジュール
import sys       # システム関連の機能を提供するモジュール
from json import JSONDecodeError  # JSONデコードエラー例外
from airtest.core.api import *    # Airtestの主要APIをインポート
from airtest.report.report import LogToHtml  # AirtestのHTMLレポート生成クラス
from colorama import Fore, Style, init       # ターミナルの文字色やスタイルを扱うモジュール
import os      # OS依存の機能(パス操作など)を提供するモジュール


init(autoreset=True)  # coloramaを初期化し、スタイルリセットを自動有効化する

print("Current Working Directory:", os.getcwd())  # 現在の作業ディレクトリを表示


def load_config():
    """設定ファイルを読み込む"""
    # config.jsonファイルを開き、JSONとして読み込んで返す
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)


def get_current_time_for_filename():
    """ファイル名に使用される現在の時刻文字列を取得します"""
    # ファイル名用に現在の時刻をフォーマットして返す
    return datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S")


def get_current_time_for_print():
    """操作ステップの印刷に使用される現在の時刻文字列を取得します"""
    # ログや操作ステップ出力用に現在の時刻をフォーマットして返す
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def setup_logging(log_directory, log_level):
    # すべてのロガーに対して設定されたログレベルを適用する
    for logger_name in logging.root.manager.loggerDict:
        logging.getLogger(logger_name).setLevel(getattr(logging, log_level))
    """ログ設定を行う"""
    # ログファイル名に現在時刻を使用する
    current_time = get_current_time_for_filename()
    local_logger = logging.getLogger(__name__)
    local_logger.setLevel(log_level)  # config.jsonで指定されたログレベルを設定
    log_file = os.path.join(log_directory, f"error_{current_time}.log")  # ログファイルのパスを組み立て
    file_handler = logging.FileHandler(log_file)  # ログをファイルに出力するためのハンドラー
    file_handler.setLevel(log_level)  # ファイルハンドラーのログレベルを設定
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')  # ログの出力フォーマットを定義
    file_handler.setFormatter(formatter)  # フォーマッターをファイルハンドラーに設定
    local_logger.addHandler(file_handler)  # ロガーにファイルハンドラーを追加
    return local_logger


def ensure_directory_exists(directory):
    """ディレクトリが存在しない場合に作成する"""
    # 指定したディレクトリがなければ作成し、作成を通知する
    if not os.path.exists(directory):
        os.makedirs(directory)
        custom_print(f"ディレクトリを作成しました: {directory}", color=Fore.LIGHTBLUE_EX, show_step=False, update_step=False)


class Counter:
    def __init__(self):
        # ステップカウンターを1で初期化
        self.step_counter = 1

    def increment_step(self):
        # ステップカウンターを1つ増加させて返す
        self.step_counter += 1
        return self.step_counter

    def reset_step(self):
        """歩数カウンターを初期値にリセット"""
        # ステップカウンターを1にリセット
        self.step_counter = 1


# グローバルにカウンターインスタンスを作成
counter = Counter()


def custom_print(message, color=None, delay=0, update_step=True, show_step=True):
    """カスタムフォーマットでメッセージを出力する"""
    # 現在の時刻をフォーマット取得
    current_time = get_current_time_for_print()
    # show_stepがTrueなら、ステップ番号を含める
    if show_step:
        output = f"[{current_time}] {counter.step_counter}: {message}"
    else:
        output = f"[{current_time}] {message}"
    # カラーが指定されていれば文字に色を付与
    if color:
        output = f"{color}{output}{Style.RESET_ALL}"
    print(output)  # コンソールに出力
    # delayが0以上なら指定秒数待機
    if delay > 0:
        time.sleep(delay)
    # update_stepがTrueならステップカウンターを1増やす
    if update_step:
        counter.increment_step()


def handle_action_result(action_result, action_target, stop_if_not_found, delay, logger):
    """アクションの結果を処理する"""
    # action_resultがFalseかつstop_if_not_foundがTrueならエラーをスロー
    if action_result is False and stop_if_not_found:
        error_message = f"対象が見つかりません: {action_target}"
        custom_print(error_message, color=Fore.RED, update_step=False)
        logger.error(error_message)
        raise ValueError(error_message)
    # delayが指定されていれば指定秒数待機
    if delay > 0:
        time.sleep(delay)


def format_json_line_by_line(file_path):
    # ログファイル(1行1JSON)を整形してファイルに再書き込み
    formatted_lines = []
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            if line.strip():
                json_object = json.loads(line)
                formatted_line = json.dumps(json_object, indent=4, ensure_ascii=False)
                formatted_lines.append(formatted_line)
    with open(file_path, 'w', encoding='utf-8') as file:
        for line in formatted_lines:
            file.write(line + "\n")
    custom_print('ログファイルは正常にフォーマットされました', update_step=False, show_step=False, color=Fore.LIGHTBLUE_EX, delay=2)


def perform_action(action_type, action_target, description, logger, stop_if_not_found=True, delay=0, update_step=True, use_rgb=False, **kwargs):
    """
    UIアクションを実行する

    :param action_type: どのアクションを実行するかを表す文字列
    :param action_target: テンプレート画像やアプリ名など、アクションの対象を表す
    :param description: このアクションの説明文
    :param logger: ログ出力用のロガー
    :param stop_if_not_found: Trueの場合、対象が見つからなければ処理を停止する
    :param delay: アクション後に待機する秒数
    :param update_step: Trueの場合、ステップカウンターを更新する
    :param use_rgb: Trueの場合、テンプレートマッチングでRGBを使用する
    :param kwargs: その他アクションに必要な追加パラメータ(例: swipeのstart_point, end_pointなど)
    """
    # アクションごとの実行関数をディクショナリにまとめる
    action_funcs = {
        'wait': lambda: wait(Template(action_target, rgb=use_rgb), timeout=10, **kwargs),
        'swipe': lambda: swipe(kwargs['start_point'], kwargs['end_point'], **kwargs),
        'exists': lambda: exists(Template(action_target)),
        'touch': lambda: touch(Template(action_target), **kwargs),
        'start_app': lambda: start_app(action_target),
        'text': lambda: text(action_target, **kwargs),
        'snapshot': lambda: snapshot(**kwargs),
        'double_touch': lambda: double_click(Template(action_target)),
        'key_press': lambda: keyevent(kwargs['key']),
        'assert_exists': lambda: assert_exists(Template(action_target), description),
        'stop_app': lambda: stop_app(action_target)
    }

    try:
        # アクション説明をカラー付きで出力("検出中："が含まれる場合は緑、それ以外は黄色)
        custom_print(description, color=Fore.LIGHTGREEN_EX if "検出中：" in description else Fore.LIGHTYELLOW_EX,
                     update_step=update_step)
        # action_typeに対応した関数を呼び出して結果を得る
        result = action_funcs[action_type]()
        # アクション結果をハンドリング
        handle_action_result(result, action_target, stop_if_not_found, delay, logger)
        return result
    except KeyError:
        # action_typeが定義されていない場合のエラー
        error_message = f"サポートされていないaction_type: {action_type}"
        custom_print(error_message, color=Fore.RED, update_step=False)
        logger.error(error_message)
        raise NotImplementedError(error_message)
    except Exception as e:
        # それ以外のエラーが発生した場合の処理
        error_message = f"アクション {action_type} 実行中にエラーが発生しました、対象：{action_target}、エラー：{str(e)}"
        custom_print(error_message, color=Fore.RED, update_step=False)
        logger.error(error_message)
        sys.exit(1)


def generate_video_filename():
    """ビデオファイル名を生成する"""
    # 実行中のファイル名(拡張子なし)を取得し、現在時刻を組み合わせたファイル名を作成
    current_file_name = os.path.splitext(os.path.basename(__file__))[0]
    current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    return f"{current_file_name}_{current_time}.mp4"


def generate_report(script_root, log_root, export_dir, logfile):
    """ログからHTMLレポートを生成する"""
    # レポートのエクスポート先に現在時刻付きのフォルダを用意
    current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    export_sub_dir = os.path.join(export_dir, f"report_{current_time}")

    # ディレクトリが存在しなければ作成
    if not os.path.exists(export_sub_dir):
        os.makedirs(export_sub_dir)
    # AirtestのHTMLレポートを生成するクラスを初期化
    html_reporter = LogToHtml(script_root=script_root,
                              log_root=log_root,
                              export_dir=export_sub_dir,
                              logfile=logfile,
                              lang='en')
    # レポート作成
    html_reporter.report()


def main():
    """メイン関数、自動化スクリプトを実行する"""
    # config.jsonを読み込み
    config = load_config()
    # OS判定を行い、対応する設定を取得
    platform_config = config['windows'] if sys.platform.startswith('win') else config['mac']
    log_directory_airtest = platform_config['log_root']
    log_level = platform_config['log_level']

    # ログフォルダが存在する場合は削除し、削除を通知
    if os.path.exists(log_directory_airtest):
        shutil.rmtree(log_directory_airtest)
        custom_print(f"ログフォルダが削除されました: {log_directory_airtest}", color=Fore.LIGHTBLUE_EX, show_step=False, update_step=False)

    times = 0  # テスト実行回数を格納する変数
    while True:
        try:
            # ユーザーからテスト回数を入力してもらう
            times = int(input("自動テストを実行する回数を入力してください:"))
            if times <= 0:
                raise ValueError
            break
        except ValueError:
            # 整数以外や0以下の場合のエラー処理
            print("入力エラー.0 より大きい整数を入力してください.")
            continue

    # Airtestの自動セットアップ(スクリプト、ログディレクトリ、接続デバイスなどを指定)
    auto_setup(__file__, logdir=True, devices=[platform_config['device_connection_string']])
    dev = device()  # 接続したデバイスを取得
    video_filename = generate_video_filename()  # 録画ファイル名を生成
    # 録画を開始(最大10時間、画面の向きは2など)
    dev.start_recording(output=video_filename, max_time=36000, orientation=2)
    custom_print('録画開始', update_step=False, show_step=False, color=Fore.LIGHTBLUE_EX, delay=2)
    ensure_directory_exists(platform_config['export_dir'])  # エクスポートディレクトリを作成(なければ)
    logger = setup_logging(log_directory_airtest, log_level)  # ロガーをセットアップ

    try:
        # 入力された回数だけテストを繰り返す
        for iteration in range(times):
            print(f"第{iteration + 1}回のテストを開始")
            counter.reset_step()  # ステップカウンターを1に戻す
            # Google Mapsアプリを起動して15秒待機
            perform_action("start_app", "com.google.Maps", "アプリ起動", logger, delay=15)
            # マップの検索ボックスをタッチ
            perform_action("touch", "search.png", "検索ボックスをタッチ", logger)
            # テキスト入力して検索
            perform_action("text", "東京タワー", "テキスト入力", logger, text="東京タワー")
            # Google Mapsアプリを停止して3秒待機
            perform_action("stop_app", "com.google.Maps", "アプリ閉じる", logger, delay=3)
            print(f"第{iteration + 1}回のテストを終了")
    except Exception as exc:
        # スクリプトエラーが発生した場合
        print(f"スクリプトエラー: {str(exc)}")
        sys.exit(1)
    finally:
        # テスト終了かエラー発生で録画停止
        dev.stop_recording()
        custom_print('録画終了', update_step=False, show_step=False, color=Fore.LIGHTBLUE_EX, delay=2)
        # ログからHTMLレポートを生成
        generate_report(platform_config['script_root'],
                        platform_config['log_root'],
                        platform_config['export_dir'],
                        platform_config['logfile'])
        # ログファイルを整形
        format_json_line_by_line(platform_config['logfile'])


if __name__ == "__main__":
    main()
