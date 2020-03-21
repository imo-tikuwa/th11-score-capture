# -*- coding: utf-8 -*-
import numpy as np
import os
import re
import click
import time
import cv2
import win32gui
import ctypes
import pywintypes
from PIL import ImageGrab
import sys
from datetime import datetime
# print出力に色付ける
from termcolor import colored
# 関数、定数をインポート
from commons import *
# CSV出力
# import csv

@click.command()
@click.option('--development','-dev',is_flag=True) # 開発モードのとき解析に使用した画像を保管する
def main(development):

    # コンフィグ初期化
    config = config_init()

    # 地霊殿のハンドルを起動、起動してなかったら起動
    th11_handle = execute_th11(config)

    # とりあえず5秒毎に画面をキャプチャ
    # スコア、残機の辺りを抽出する
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    try:
        current_difficulty = None
        while(True):
            time.sleep(5)

            rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)

            # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
            cap_left, cap_top, cap_right, cap_bottom = ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom)

            # 指定した領域内をクリッピング
            current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
            img = ImageGrab.grab(bbox=(cap_left, cap_top, cap_right, cap_bottom))
            if development:
                img.save(OUTPUT_DIR + current_time + '.png')
            original_frame = np.array(img)


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

            print('----- ' + current_time + '.png -----')
            print("スコア　 ： " + str(score))
            print("残機　　 ： " + str(remain))
            print("グレイズ ： " + graze)
            print("難易度　 ： " + convert_difficulty(difficulty))
            print("ボス　　 ： " + convert_boss_name(boss_name))
            print("ボス残機 ： " + convert_boss_remain(boss_remain))
            print("スペル　 ： " + convert_spell_card(spell_card))

    except pywintypes.error:
        print(colored("\n\n東方地霊殿が終了したのでプログラムも終了します", "green"))
        exit(0)

    except KeyboardInterrupt:
        print(colored("プログラムを終了します", "green"))
        exit(0)

    # CSV出力
#     csv_data = [[1234, 1, 100, 'Lunatic', '', ''],[12345, 1.2, 200, 'Lunatic', '', '']]
#     with open(OUTPUT_DIR + datetime.now().strftime('%Y%m%d%H%M%S') + '.csv', 'w') as file:
#         writer = csv.writer(file)
#         writer.writerows(csv_data)

if __name__ == '__main__':
    main()

