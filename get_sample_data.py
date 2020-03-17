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
# 関数、定数をインポート
from commons import *

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

    # とりあえず5秒毎に画面をキャプチャ
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
            current_time = datetime.now().strftime('%Y%m%d%H%M%S')
            img = ImageGrab.grab(bbox=(cap_left,cap_top,cap_right,cap_bottom))
            img.save(OUTPUT_DIR + current_time + '.png')
            original_frame = np.array(img)

            get_sample_data(original_frame, current_time)

    except KeyboardInterrupt:
        print(colored("プログラムを終了します", "green"))
        exit(0)

if __name__ == '__main__':
    main()

