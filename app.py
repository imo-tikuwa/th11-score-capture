# -*- coding: utf-8 -*-
import numpy as np
import os
import re
import click
import time
import cv2
import subprocess
import win32gui
import ctypes
from PIL import ImageGrab
import sys
import tkinter, tkinter.filedialog, tkinter.messagebox
from datetime import datetime
# print出力に色付ける
from termcolor import colored
import colorama
colorama.init()
# 設定ファイルを利用する
import configparser
# 関数をインポート
from functions import *


# 変数(定数扱いする変数)
# --------------------------------------------------
TH11_WINDOW_NAME = '東方地霊殿　～ Subterranean Animism. ver 1.00a'
CONFIG_FILE_NAME = 'settings.ini'
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'output' + os.sep
SAMPLE_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'sample_data' + os.sep
SAMPLE_NUMBERS_DIR = SAMPLE_DIR + 'number' + os.sep
SAMPLE_REMAINS_DIR = SAMPLE_DIR + 'remain' + os.sep

# スコアのROI配列(10億、1億、1000万...の順)
SCORE_ROIS = []
width = 24
height = 32
roi_start_position = (1020, 144)
for index in range(10):
    left = roi_start_position[0] + (width * index)
    top = roi_start_position[1]
    right = left + width
    bottom = top + height
    SCORE_ROIS.append((left, top, right, bottom))

# 残機のROI配列(残1、残2、残3...の順)
REMAIN_ROIS = []
width = 24
height = 32
roi_start_position = (1046, 208)
for index in range(9):
    left = roi_start_position[0] + (width * index)
    top = roi_start_position[1]
    right = left + width
    bottom = top + height
    REMAIN_ROIS.append((left, top, right, bottom))

# スコアのサンプルデータ(0～9)
BINARY_NUMBERS = []
for index in range(10):
    img = cv2.imread(SAMPLE_NUMBERS_DIR + str(index) + '.png', cv2.IMREAD_GRAYSCALE) #グレースケールで読み込み
    BINARY_NUMBERS.append(img)

# 残機のサンプルデータ(1、1/5、2/5、3/5、4/5)
BINARY_REMAINS = []
for index in range(5):
    img = cv2.imread(SAMPLE_REMAINS_DIR + str(index) + '.png', cv2.IMREAD_GRAYSCALE) #グレースケールで読み込み
    BINARY_REMAINS.append(img)

# --------------------------------------------------

@click.command()
def main():

    # コンフィグにexeファイルの設定が存在するか確認
    config = configparser.ConfigParser()
    config_section = 'config'
    config.read(CONFIG_FILE_NAME, 'cp932')

    # 地霊殿が起動してるかチェック
    th11_handle = win32gui.FindWindow(None, TH11_WINDOW_NAME)
    if th11_handle <= 0:

        # 設定ファイルチェック
        th11_exe_file = ""
        if config.has_section(config_section):
            th11_exe_file = config.get(config_section, 'exe_path')
            if not os.path.isfile(th11_exe_file):
                th11_exe_file = ""

        # ダイアログを開いて地霊殿の実行ファイルを指定してもらう
        if th11_exe_file == "":
            # th11.exe指定
            root = tkinter.Tk()
            root.withdraw()

            file_type = [("東方地霊殿の実行ファイル", "th11.exe"),("全てのファイル", "*.*")]
            initial_dir = os.path.abspath(os.path.dirname(__file__))
            th11_exe_file = tkinter.filedialog.askopenfilename(filetypes=file_type, initialdir=initial_dir)

            if th11_exe_file == "" or os.path.basename(th11_exe_file) != 'th11.exe':
                print(colored("東方地霊殿の実行ファイルを指定してください。", "red"))
                sys.exit(1)

            # 設定ファイルにexeファイルのパス保存
            if not config.has_section(config_section):
                config.add_section(config_section)
            config.set(config_section, 'exe_path', th11_exe_file)

            # ファイルに書き出し
            with open(CONFIG_FILE_NAME, 'w') as config_file:
                config.write(config_file)

        # 東方地霊殿を起動
        th11_exe_dir = os.path.dirname(th11_exe_file)
        th11_exe_name = os.path.basename(th11_exe_file)
        subprocess.Popen('cd /D ' + th11_exe_dir + ' && start ' + th11_exe_name, shell=True) # 移動してから起動しないと設定ファイルが読み込まれないみたい
        time.sleep(5)

        # ウィンドウ名でハンドル取得
        while(True):
            th11_handle = win32gui.FindWindow(None, TH11_WINDOW_NAME)
            if th11_handle > 0:
                break

            print(colored("東方地霊殿が起動してないよー", "red"))
            time.sleep(3)

    print(colored("東方地霊殿が起動してるよー", "green"))
    print(colored("東方地霊殿のハンドル：" + str(th11_handle), "green"))
    print(colored("Ctrl+Cで終了します", "green"))

    # とりあえず10秒毎に画面をキャプチャ
    # スコア、残機の辺りを抽出する
    # スコアの表示は確認したところ等間隔なので1文字ずつ抽出した方が簡単そう
    try:
        while(True):
            time.sleep(10)

            rect_left,rect_top,rect_right,rect_bottom = win32gui.GetWindowRect(th11_handle)

            # ウィンドウの外枠＋数ピクセル余分にとれちゃうので1280x960の位置補正
            cap_left = rect_left + 3
            cap_top = rect_top + 26
            cap_right = cap_left + 1280
            cap_bottom = cap_top + 960

            # 指定した領域内をクリッピング
            current_time = datetime.now().strftime('%Y%m%d%H%M%S')
            img = ImageGrab.grab(bbox=(cap_left,cap_top,cap_right,cap_bottom))
            img.save(OUTPUT_DIR + current_time + '.png')
            original_frame = np.array(img)

            # 二値化
            work_frame = edit_frame(original_frame)

            # スコアの数字についてテンプレートマッチング
            score = analyze_score(SCORE_ROIS, BINARY_NUMBERS, work_frame)

            # 残機についてテンプレートマッチング
            remain = analyze_remain(REMAIN_ROIS, BINARY_REMAINS, work_frame)

            print("スコア ： " + score)
            print("残機　 ： " + str(remain))

    except KeyboardInterrupt:
        print(colored("プログラムを終了します", "green"))
        exit(0)

if __name__ == '__main__':
    main()

