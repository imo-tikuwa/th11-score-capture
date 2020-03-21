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
@click.option('--development','-dev',is_flag=True) # 開発モードのとき解析に使用した画像を保管する
@click.option('--output','-out',is_flag=True)      # コンソールにログを出力する
def main(development, output):

    # コンフィグ初期化
    config = config_init()

    # 地霊殿のハンドルを起動、起動してなかったら起動
    th11_handle = execute_th11(config)

    # スコア、残機の辺りを抽出する
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    results = []
    try:
        current_difficulty = None
        sleep_second = 3
        while(True):
            if (sleep_second > 0):
                time.sleep(sleep_second)

            rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)

            # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
            capture_area = ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom)

            # 指定した領域内をクリッピング
            current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
            original_frame = get_original_frame(capture_area, current_time, development)

            # 二値化
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

            # ボス残機についてテンプレートマッチング
            boss_remain = analyze_boss_remain(original_frame, boss_name)

            # スペルカードについてテンプレートマッチング
            spell_card = analyze_spell_card(original_frame, boss_name)

            # ラストスペルのときステージクリアの瞬間をキャプチャするため一時的に処理間隔を0.5秒に変更
            is_last_spell = check_is_last_spell(spell_card)
            if (is_last_spell):
                sleep_second = 0.5

            # ステージクリアについてテンプレートマッチング
            is_stage_clear = analyze_stage_clear(work_frame)
            if (is_stage_clear):
                sleep_second = 3

            # ステージクリア判定
            current_position = convert_stage_clear(is_stage_clear)

            # コンソール出力
            if (output):
                output_console(current_time, difficulty, score, remain, graze, boss_name, boss_remain, spell_card, current_position)

            # 結果を格納
            results.append([difficulty, score, remain, graze, boss_name, boss_remain, spell_card, current_position])

    except pywintypes.error:
        print(colored("\n\n東方地霊殿が終了したのでプログラムも終了します", "green"))

    except KeyboardInterrupt:
        print(colored("プログラムを終了します", "green"))

    # CSV出力
    if (len(results) > 0):

        # 重複を含めたすべてのデータをCSV出力
        save_datetime = datetime.now().strftime('%Y%m%d%H%M%S')
        save_csv(save_datetime + '_all_result.csv', results)

        # ため込んだデータをforループして末尾のデータから順番に検証、以下の処理を実施
        # 1. スペルカードが重複するデータを削除する(古いものが残るようにする)
        # 2. currentのボス名が空、prevとnextのボス名が存在するときcurrentを会話中orデータ取得に失敗とみなし削除
        results_length = len(results)
        for current_index in reversed(range(results_length)):
            # current_numのレコードを中心に前、現在、後の3レコードを取得
            current = results[current_index]
            prev = next = None
            if (current_index > 0):
                prev = results[current_index - 1]
            if (current_index < len(results) - 1):
                next = results[current_index + 1]

            # prevのレコードとcurrentのレコードのボス名、ボス残機、スペルカードが一致していたらcurrentは重複と見なして削除
            if (prev is not None and current[CSV_INDEX_BOSS_NAME] == prev[CSV_INDEX_BOSS_NAME] and current[CSV_INDEX_BOSS_REMAIN] == prev[CSV_INDEX_BOSS_REMAIN] and current[CSV_INDEX_SPELL_CARD] == prev[CSV_INDEX_SPELL_CARD]):
                del results[current_index]
                continue

            # currentのボス名が空、prevとnextのボス名が存在するとき会話中orデータ取得に失敗とみなしてcurrentを削除
            if (prev is not None and next is not None):
                if (current[CSV_INDEX_BOSS_NAME] == '' and prev[CSV_INDEX_BOSS_NAME] != '' and next[CSV_INDEX_BOSS_NAME] != '' and prev[CSV_INDEX_BOSS_NAME] == next[CSV_INDEX_BOSS_NAME]):
                    del results[current_index]
                    continue

        # 重複を削除したデータをCSV出力
        save_csv(save_datetime + '_result.csv', results)
        print(colored("結果をCSVに出力しました。", "green"))



if __name__ == '__main__':
    main()

