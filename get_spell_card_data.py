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
def main():

    # コンフィグ初期化
    config = config_init()

    # 地霊殿のハンドルを起動、起動してなかったら起動
    th11_handle = execute_th11(config)

    capture_count = 0
    while(True):

        # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
        rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)

        # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
        cap_left, cap_top, cap_right, cap_bottom = ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom)

        # 指定した領域内をクリッピング
        current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
        img = ImageGrab.grab(bbox=(cap_left,cap_top,cap_right,cap_bottom))
    #     img.save(OUTPUT_DIR + current_time + '.png')
        original_frame = np.array(img)

        # スペルカードのの画像を保存
        roi = (340, 64, 820, 97)
        clopped_frame = original_frame[roi[1]:roi[3], roi[0]:roi[2]]
        file_name = OUTPUT_DIR + current_time + '_spell.png'
        Image.fromarray(clopped_frame).save(file_name)

        # 3回キャプチャしたら処理を抜ける
        capture_count += 1
        if (capture_count >= 3):
            break
        time.sleep(0.5)

    print("スペルカードのサンプルデータを取得しました。\n左右の余分なピクセルは必要に応じて切り取ってください")


if __name__ == '__main__':
    main()

