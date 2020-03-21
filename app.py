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
def main(development):

    # コンフィグ初期化
    config = config_init()

    # 地霊殿のハンドルを起動、起動してなかったら起動
    th11_handle = execute_th11(config)

    # スコア、残機の辺りを抽出する
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    results = []
    try:
        current_difficulty = None
        speep_second = 3
        while(True):
            if (speep_second > 0):
                time.sleep(speep_second)

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
            boss_remain = analyze_boss_remain(original_frame)

            # スペルカードについてテンプレートマッチング
            spell_card = analyze_spell_card(original_frame)

            # コンソール出力
            difficulty = convert_difficulty(difficulty)
            score = str(score)
            remain = str(remain)
            boss_name = convert_boss_name(boss_name)
            boss_remain = convert_boss_remain(boss_remain, boss_name)
            spell_card = convert_spell_card(spell_card)
            print('----- ' + current_time + '.png -----')
            print("難易度　 ： " + difficulty)
            print("スコア　 ： " + score)
            print("残機　　 ： " + remain)
            print("グレイズ ： " + graze)
            print("ボス　　 ： " + boss_name)
            print("ボス残機 ： " + boss_remain)
            print("スペル　 ： " + spell_card)

            # 結果を格納
            results.append([difficulty, score, remain, graze, boss_name, boss_remain, spell_card])

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
            if (prev is not None and current[4] == prev[4] and current[5] == prev[5] and current[6] == prev[6]):
                del results[current_index]
                continue

            # currentのボス名が空、prevとnextのボス名が存在するとき会話中orデータ取得に失敗とみなしてcurrentを削除
            if (prev is not None and next is not None):
                if (current[4] == '' and prev[4] != '' and next[4] != '' and prev[4] == next[4]):
                    del results[current_index]
                    continue

        # 重複を削除したデータをCSV出力
        save_csv(save_datetime + '_result.csv', results)
        print(colored("結果をCSVに出力しました。", "green"))



if __name__ == '__main__':
    main()

