# -*- coding: utf-8 -*-
import os
import click
import time
import win32gui
import pywintypes
from datetime import datetime
# print出力に色付ける
from termcolor import colored
# 関数、定数をインポート
from commons import *


@click.command()
@click.option('--development', '-dev', is_flag = True, help = "開発モード(outputディレクトリに解析に使用した画像を保存)")
@click.option('--output', is_flag = True, help = "指定したときコンソールに暫定の解析結果の出力を行います")
@click.option('--capture-period', default = 1.0, type = click.FloatRange(0.0, 10.0), help = "画面のキャプチャ間隔(秒)")
@click.option('--print-exec-time', is_flag = True, help = "テンプレートマッチング処理全体の処理時間を出力します。この数値を見てcapture-periodを設定するのを推奨")
def main(development, output, capture_period, print_exec_time):

    print_usage()
    try:

        # コンフィグ初期化
        config = config_init()

        # 地霊殿のハンドルを起動、起動してなかったら起動
        th11_handle = execute_th11(config)

        # ウィンドウサイズチェック
        rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)
        if (rect_left < 0 or rect_top < 0 or rect_right < 0 or rect_bottom < 0):
            print(colored("東方地霊殿のウィンドウが最小化されていて取得できませんでした", "red", attrs=['bold']))
            exit(0)
        elif (rect_right - rect_left < TH11_WINDOW_ALLOW_WIDTH or rect_bottom - rect_top < TH11_WINDOW_ALLOW_HEIGHT):
            print(colored("東方地霊殿は1280x960のウィンドウサイズで起動してください", "red", attrs=['bold']))
            exit(0)

        # capture関数自体を丸ごと呼びなおしてスペルカードのサンプルデータなんかロードも全てやり直す
        while (True):
            csv_save = capture(th11_handle, development, output, capture_period, print_exec_time)
            if (csv_save):
                time.sleep(2)
                print("\n続けてキャプチャする場合は c と\n終了する場合はそれ以外のキーを入力してください。")
                cmd = input("入力：")
                if cmd != "c":
                    exit(0)
            else:
                # CSVを出力せずに抜けてきた場合はこのままプログラム自体を終了する
                exit(0)

    except KeyboardInterrupt:
        exit(0)


def capture(th11_handle, development, output, capture_period, print_exec_time):
    # キャプチャ処理
    # CSVを保存したらTrue、しなかったらFalseを返す

    # スコア、残機の辺りを抽出する
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    results = []
    try:
        print(colored("キャプチャ準備完了", "green", attrs=['bold']))
        current_difficulty = None
        sleep_second = capture_period
        capture_count = 0
        while(True):
            if (sleep_second > 0):
                time.sleep(sleep_second)

            if (print_exec_time):
                start_time = time.time()

            rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)

            # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
            capture_area = ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom)

            # 指定した領域内をクリッピング
            current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
            original_frame = get_original_frame(capture_area, current_time, development)

            # グレースケール化
            work_frame = edit_frame(original_frame)

            # 難易度についてテンプレートマッチング
            difficulty = analyze_difficulty(work_frame)

            # 難易度が見つからないとき後続の処理を全てスキップ
            if (difficulty is None):
                continue

            # 初めて難易度を見つけたとき or 難易度が変更されたのが確認できたときスペルカードのサンプルデータを初期化
            if (current_difficulty is None or difficulty != current_difficulty):
                current_difficulty = difficulty
                load_spell_card_binaries(difficulty)

            # スコアの数字についてテンプレートマッチング
            score = analyze_score(work_frame)

            # 残機についてテンプレートマッチング
            remain = analyze_remain(work_frame)

            # グレイズの数字についてテンプレートマッチング
            graze = analyze_graze(work_frame)

            # ボス名についてテンプレートマッチング
            boss_name = analyze_boss_name(original_frame)

            # ボス戦中かどうかチェック
            is_boss_attack = check_is_boss_attack(boss_name, work_frame)

            # ボス残機についてテンプレートマッチング
            boss_remain = analyze_boss_remain(original_frame, boss_name)

            # スペルカードについてテンプレートマッチング
            spell_card = analyze_spell_card(original_frame, boss_name)

            # 矛盾した状態が発生したときスキップする
            if (inconsistency_check(output, difficulty, boss_name, is_boss_attack, boss_remain, spell_card) is False):
                continue

            # ラストスペルのときステージクリアの瞬間をキャプチャするため一時的に処理間隔を短くする
            if (capture_period > SLEEP_SECOND_TURBO):
                is_last_spell = check_is_last_spell(spell_card)
                if (is_last_spell):
                    sleep_second = SLEEP_SECOND_TURBO

            # ステージクリアについてテンプレートマッチング
            is_stage_clear = analyze_stage_clear(original_frame, work_frame)
            if (is_stage_clear):
                sleep_second = capture_period

            # ステージクリア判定
            current_position = convert_stage_clear(is_stage_clear)

            # コンソール出力
            output_console(output, capture_count, current_time, difficulty, score, remain, graze, boss_name, boss_remain, spell_card, current_position)

            # 結果を格納
            results.append([difficulty, score, remain, graze, boss_name, boss_remain, spell_card, current_position])

            if (print_exec_time):
                print("1回あたりの実行時間: {0}".format(time.time() - start_time))

            capture_count += 1

    except pywintypes.error:
        print(colored("\n\n東方地霊殿が終了したのでプログラムも終了します", "green", attrs=['bold']))
        if (len(results) > 0):
            output_csv(results)

    except KeyboardInterrupt:
        print(colored("\n\nキャプチャを終了します", "green", attrs=['bold']))
        if (len(results) > 0):
            output_csv(results)
            return True

    return False


if __name__ == '__main__':
    main()


