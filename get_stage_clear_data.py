# -*- coding: utf-8 -*-
import os
import click
import time
import win32gui
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

        rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)

        # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
        capture_area = ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom)

        # 指定した領域内をクリッピング
        current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
        original_frame = get_original_frame(capture_area, current_time, True)

        # ステージクリアの画像を保存
        clopped_frame = original_frame[STAGE_CLEAR_ROI[1]:STAGE_CLEAR_ROI[3], STAGE_CLEAR_ROI[0]:STAGE_CLEAR_ROI[2]]
        file_name = OUTPUT_DIR + current_time + '_stage_clear.png'
        Image.fromarray(clopped_frame).save(file_name)

        # 3回キャプチャしたら処理を抜ける
        capture_count += 1
        if (capture_count >= 3):
            break
        time.sleep(0.5)

    print("ステージクリアのサンプルデータを取得しました。\n左右の余分なピクセルは必要に応じて切り取ってください")


if __name__ == '__main__':
    main()

