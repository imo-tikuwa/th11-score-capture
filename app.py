# -*- coding: utf-8 -*-
import numpy as np
import os
import re
import click
import time
import cv2
import win32gui
import ctypes
from PIL import ImageGrab
import sys
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

    # とりあえず5秒毎に画面をキャプチャ
    # スコア、残機の辺りを抽出する
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    try:
        while(True):
            time.sleep(5)

            rect_left,rect_top,rect_right,rect_bottom = win32gui.GetWindowRect(th11_handle)

            # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
            cap_left = rect_left + 3
            cap_top = rect_top + 26
            cap_right = cap_left + 1280
            cap_bottom = cap_top + 960

            # 指定した領域内をクリッピング
            current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
            img = ImageGrab.grab(bbox=(cap_left, cap_top, cap_right, cap_bottom))
            if development:
                img.save(OUTPUT_DIR + current_time + '.png')
            original_frame = np.array(img)


            # 二値化
            work_frame = edit_frame(original_frame)

            # スコアの数字についてテンプレートマッチング
            score = analyze_score(work_frame)

            # 残機についてテンプレートマッチング
            remain = analyze_remain(work_frame)

            # グレイズの数字についてテンプレートマッチング
            graze = analyze_graze(work_frame)

            # 難易度についてテンプレートマッチング
            difficulty = analyze_difficulty(work_frame)

            # ボス名についてテンプレートマッチング
            boss_name = analyze_boss_name(original_frame)

            # ボス残機についてテンプレートマッチング
            boss_remain = analyze_boss_remain(original_frame)

            print('----- ' + current_time + '.png -----')
            print("スコア　 ： " + str(score))
            print("残機　　 ： " + str(remain))
            print("グレイズ ： " + graze)
            print("難易度　 ： " + convert_difficulty(difficulty))
            print("ボス　　 ： " + convert_boss_name(boss_name))
            print("ボス残機 ： " + convert_boss_remain(boss_remain))

    except KeyboardInterrupt:
        print(colored("プログラムを終了します", "green"))
        exit(0)

if __name__ == '__main__':
    main()

