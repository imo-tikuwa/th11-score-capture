# -*- coding: utf-8 -*-
import os
import click
import time
import cv2
import win32gui
from datetime import datetime
# print出力に色付ける
from termcolor import colored
# 関数、定数をインポート
from commons import *

@click.command()
@click.option('--development','-dev',is_flag=True)
def main(development):

    # コンフィグ初期化
    config = config_init()

    # 地霊殿のハンドルを起動、起動してなかったら起動
    th11_handle = execute_th11(config)

    # とりあえず5秒毎に画面をキャプチャ
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    try:
        while(True):
            time.sleep(5)

            rect_left, rect_top, rect_right, rect_bottom = win32gui.GetWindowRect(th11_handle)

            # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
            capture_area = ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom)

            # 指定した領域内をクリッピング
            current_time = datetime.now().strftime('%Y%m%d%H%M%S%f')
            original_frame = get_original_frame(capture_area, current_time, development)

            # スコアの画像を保存（グレイズにも使用）
            for index, roi in enumerate(SCORE_ROIS):
                clopped_frame = original_frame[roi[1]:roi[3], roi[0]:roi[2]]
                file_name = OUTPUT_DIR + current_time + '_score_' + str(index) + '.png'
                # OpenCvはBGR、PillowはRGBなのでOpenCvで保存するときはモードを指定しないと色が変わってしまう
#                 clopped_frame = cv2.cvtColor(clopped_frame, cv2.COLOR_BGR2RGB)
#                 cv2.imwrite(file_name, clopped_frame)
                # Pillowで保存
                Image.fromarray(clopped_frame).save(file_name)

            # スコアの画像を保存（グレイズにも使用）
            clopped_frame = original_frame[TIME_REMAIN_ROI[1]:TIME_REMAIN_ROI[3], TIME_REMAIN_ROI[0]:TIME_REMAIN_ROI[2]]
            file_name = OUTPUT_DIR + current_time + '_time_remain.png'
            Image.fromarray(clopped_frame).save(file_name)

            # 残機の画像を保存
            for index, roi in enumerate(REMAIN_ROIS):
                clopped_frame = original_frame[roi[1]:roi[3], roi[0]:roi[2]]
                file_name = OUTPUT_DIR + current_time + '_remain_' + str(index) + '.png'
                Image.fromarray(clopped_frame).save(file_name)

            # 難易度の画像を保存
            roi = (990, 44, 1130, 84)
            clopped_frame = original_frame[roi[1]:roi[3], roi[0]:roi[2]]
            file_name = OUTPUT_DIR + current_time + '_difficulty.png'
            Image.fromarray(clopped_frame).save(file_name)

            print("スコア、残機、難易度のサンプルデータを取得しました")

    except KeyboardInterrupt:
        print(colored("プログラムを終了します", "green"))
        exit(0)

if __name__ == '__main__':
    main()

