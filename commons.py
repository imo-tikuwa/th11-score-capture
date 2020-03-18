
import os
import time
import cv2
import subprocess
from datetime import datetime
from PIL import Image
import win32gui
import tkinter, tkinter.filedialog, tkinter.messagebox
# print出力に色付ける
from termcolor import colored
import colorama
colorama.init()
# 設定ファイルを利用する
import configparser


# 変数(定数扱いする変数)
TH11_WINDOW_NAME = '東方地霊殿　～ Subterranean Animism. ver 1.00a'
CONFIG_FILE_NAME = 'settings.ini'
CONFIG_SECTION_NAME = 'config'
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'output' + os.sep
SAMPLE_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'sample_data' + os.sep
SAMPLE_NUMBERS_DIR = SAMPLE_DIR + 'number' + os.sep
SAMPLE_REMAINS_DIR = SAMPLE_DIR + 'remain' + os.sep
SAMPLE_DIFFICULTIES_DIR = SAMPLE_DIR + 'difficulty' + os.sep
SAMPLE_BOSSNAMES_DIR = SAMPLE_DIR + 'bossname' + os.sep
DIFFICULTY_HASHMAP = {
                      None: '',
                      0: 'Easy',
                      1: 'Normal',
                      2: 'Hard',
                      3: 'Lunatic',
                      4: 'Extra'
}
BOSSNAME_HASHMAP = {
                      None: '',
                      0: 'キスメ',
                      1: '黒谷ヤマメ',
                      2: '水橋パルスィ',
                      3: '星熊勇儀',
                      4: '古明地さとり',
                      5: '火焔猫燐',
                      6: '霊烏路空',
                      7: '東風谷早苗',
                      8: '古明地こいし'
}
BOSSNAME2STAGE_HASHMAP = {
                          0: '1面',
                          1: '1面',
                          2: '2面',
                          3: '3面',
                          4: '4面',
                          5: '5面',
                          6: '6面',
                          7: 'EX面',
                          8: 'EX面',
}

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

# グレイズのROI配列(一万の桁、千の桁、百の桁...の順)
GRAZE_ROIS = []
width = 24
height = 32
roi_start_position = (1044, 304)
for index in range(5):
    left = roi_start_position[0] + (width * index)
    top = roi_start_position[1]
    right = left + width
    bottom = top + height
    GRAZE_ROIS.append((left, top, right, bottom))

# 難易度のROI
DIFFICULTY_ROI = (972, 37, 1136, 90)

# ボス名のROI
BOSSNAME_ROI = (76, 56, 246, 72)

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

# 難易度のサンプルデータ(Easy～Extra)
BINARY_DIFFICULTIES = []
for index in range(5):
    img = cv2.imread(SAMPLE_DIFFICULTIES_DIR + str(index) + '.png', cv2.IMREAD_GRAYSCALE) #グレースケールで読み込み
    BINARY_DIFFICULTIES.append(img)

# ボス名のサンプルデータ(キスメ～古明地こいし)
# ボス名は単色なのでinRangeで色を指定して二値化してから使用する
BINARY_BOSSNAMES = []
for index in range(9):
    img = cv2.imread(SAMPLE_BOSSNAMES_DIR + str(index) + '.png')
    img = cv2.inRange(img, (255, 255, 119), (255, 255, 119))
    BINARY_BOSSNAMES.append(img)


def config_init():
    # コンフィグを初期化
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE_NAME, 'cp932')

    return config


def execute_th11(config):
    # 東方地霊殿を起動してハンドルを返す。
    # 実行ファイルのパスが存在しない場合はダイアログを開き設定する
    th11_handle = win32gui.FindWindow(None, TH11_WINDOW_NAME)
    if th11_handle <= 0:

        # 設定ファイルチェック(exeファイルの設定が存在するか確認)
        th11_exe_file = ""
        if config.has_section(CONFIG_SECTION_NAME):
            th11_exe_file = config.get(CONFIG_SECTION_NAME, 'exe_path')
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
            if not config.has_section(CONFIG_SECTION_NAME):
                config.add_section(CONFIG_SECTION_NAME)
            config.set(CONFIG_SECTION_NAME, 'exe_path', th11_exe_file)

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

    return th11_handle


def edit_frame(frame):
    # フレームを二値化
    work_frame = frame

    work_frame = cv2.cvtColor(work_frame, cv2.COLOR_RGB2GRAY)
#     work_frame = cv2.threshold(work_frame, FRAME_THRESH, 255, cv2.THRESH_BINARY)[1]
#     work_frame = cv2.bitwise_not(work_frame)

    return work_frame

def analyze_score(work_frame):
    # スコアをテンプレートマッチングにより取得
    score_results = []
    for index, roi in enumerate(SCORE_ROIS):
        clopped_frame = work_frame[roi[1]:roi[3], roi[0]:roi[2]]

        for num, template_img in enumerate(BINARY_NUMBERS):
            res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#             print(min_val, max_val, min_loc, max_loc)
            if (max_val > 0.98):
                score_results.append(num)
                break

    score = "".join(map(str, score_results))
    # 1億未満のとき先頭に0がついてしまうので数値に変換
    if (score):
        score = int(score)

    return score


def analyze_remain(work_frame):
    # 残機をテンプレートマッチングにより取得
    remain = 0
    for index, roi in enumerate(REMAIN_ROIS):
        clopped_frame = work_frame[roi[1]:roi[3], roi[0]:roi[2]]

        temp_max_val = current_remain = 0;
        for num, template_img in enumerate(BINARY_REMAINS):
            res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#             print(min_val, max_val, min_loc, max_loc)
            if (max_val > 0.99):
                if (num == 0):
                    current_remain = 1
                elif (num > 0):
                    current_remain = num * 0.2 # 残機画像の1.png～4.pngは0.2をかけて表示
                break

            # 欠片はきっちり結果が分かれない模様。。
            # 0.9以上でループ中の画像より類似度が高い場合は結果を置き換える
            elif (max_val > 0.9 and temp_max_val < max_val):
                temp_max_val = max_val
                current_remain = num * 0.2

        remain += current_remain

    return remain


def analyze_graze(work_frame):
    # グレイズをテンプレートマッチングにより取得
    graze_results = []
    for index, roi in enumerate(GRAZE_ROIS):
        clopped_frame = work_frame[roi[1]:roi[3], roi[0]:roi[2]]

        for num, template_img in enumerate(BINARY_NUMBERS):
            res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#             print(min_val, max_val, min_loc, max_loc)
            if (max_val > 0.98):
                graze_results.append(num)
                break

    graze = "".join(map(str, graze_results))

    return graze


def analyze_difficulty(work_frame):
    # 難易度をテンプレートマッチングにより取得
    difficulty = None
    clopped_frame = work_frame[DIFFICULTY_ROI[1]:DIFFICULTY_ROI[3], DIFFICULTY_ROI[0]:DIFFICULTY_ROI[2]]

    for num, template_img in enumerate(BINARY_DIFFICULTIES):
        res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#             print(min_val, max_val, min_loc, max_loc)
        if (max_val > 0.99):
            difficulty = num
            break

    return difficulty


def analyze_current(original_frame, work_frame):
    # 現在の場所をテンプレートマッチングにより取得
    # ボス名、ボス名下の星の数、スペルカード名などからキャプチャした画像がどこかを判定する
    current = None

    # ボス名は単色（R:119、G:255、B:255）なので色を指定して二値化することで完全一致に近いマッチングが可能
    # 二値化の色指定について
    # OpenCv(cv2)の色の並びはBGRだけど、ここのorifinal_frameはPillowで取得したRGB画像の配列データなのでRGBの順で指定
    bossname_frame = original_frame[BOSSNAME_ROI[1]:BOSSNAME_ROI[3], BOSSNAME_ROI[0]:BOSSNAME_ROI[2]]
    bossname_frame = cv2.inRange(bossname_frame, (119, 255, 255), (119, 255, 255))
#     Image.fromarray(bossname_frame).save(OUTPUT_DIR + 'bossname.png')

    bossname = None
    for num, template_img in enumerate(BINARY_BOSSNAMES):
        res = cv2.matchTemplate(bossname_frame, template_img, cv2.TM_CCORR_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#         print(min_val, max_val, min_loc, max_loc)
        if (max_val > 0.99):
            bossname = num
            break

    # 暫定でボス名までセット
    current = bossname
    return current


def get_sample_data(original_frame, current_time):
    # サンプルデータ用の切り抜いた画像を取得

    # スコアの画像を保存（グレイズにも使用）
    for index, roi in enumerate(SCORE_ROIS):
        clopped_frame = original_frame[roi[1]:roi[3], roi[0]:roi[2]]
        file_name = OUTPUT_DIR + current_time + '_score_' + str(index) + '.png'
        # OpenCvはBGR、PillowはRGBなのでOpenCvで保存するときはモードを指定しないと色が変わってしまう
#         clopped_frame = cv2.cvtColor(clopped_frame, cv2.COLOR_BGR2RGB)
#         cv2.imwrite(file_name, clopped_frame)
        # Pillowで保存
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

    print("取得しました")
    return


def convert_difficulty(difficulty):
    # 数値の難易度を文字列に変換
    return DIFFICULTY_HASHMAP[difficulty]


def convert_bossname(bossname):
    # 数値のボス名を文字列に変換
    return BOSSNAME_HASHMAP[bossname]