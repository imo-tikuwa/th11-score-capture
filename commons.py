
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
# スペルカードディレクトリ内の特定のファイル数をカウントするのに利用
import glob


# 変数(定数扱いする変数)
TH11_WINDOW_NAME = '東方地霊殿　～ Subterranean Animism. ver 1.00a'
CONFIG_FILE_NAME = 'settings.ini'
CONFIG_SECTION_NAME = 'config'
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'output' + os.sep
SAMPLE_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'sample_data' + os.sep
SAMPLE_NUMBERS_DIR = SAMPLE_DIR + 'number' + os.sep
SAMPLE_REMAINS_DIR = SAMPLE_DIR + 'remain' + os.sep
SAMPLE_DIFFICULTIES_DIR = SAMPLE_DIR + 'difficulty' + os.sep
SAMPLE_BOSS_NAMES_DIR = SAMPLE_DIR + 'boss_name' + os.sep
SAMPLE_BOSS_REMAINS_DIR = SAMPLE_DIR + 'boss_remain' + os.sep
SAMPLE_SPELL_CARDS_DIR = SAMPLE_DIR + 'spell_card' + os.sep
DIFFICULTY_EASY = 0
DIFFICULTY_NORMAL = 1
DIFFICULTY_HARD = 2
DIFFICULTY_LUNATIC = 3
DIFFICULTY_EXTRA = 4
DIFFICULTY_HASHMAP = {
                      None: '',
                      DIFFICULTY_EASY: 'Easy',
                      DIFFICULTY_NORMAL: 'Normal',
                      DIFFICULTY_HARD: 'Hard',
                      DIFFICULTY_LUNATIC: 'Lunatic',
                      DIFFICULTY_EXTRA: 'Extra'
}
BOSS_KISUME = 0
BOSS_YAMAME = 1
BOSS_PARSEE = 2
BOSS_YUGI = 3
BOSS_SATORI = 4
BOSS_ORIN = 5
BOSS_UTSUHO = 6
BOSS_SANAE = 7
BOSS_KOISHI = 8
BOSS_NAME_HASHMAP = {
                      None: '',
                      BOSS_KISUME: 'キスメ',
                      BOSS_YAMAME: '黒谷ヤマメ',
                      BOSS_PARSEE: '水橋パルスィ',
                      BOSS_YUGI: '星熊勇儀',
                      BOSS_SATORI: '古明地さとり',
                      BOSS_ORIN: '火焔猫燐',
                      BOSS_UTSUHO: '霊烏路空',
                      BOSS_SANAE: '東風谷早苗',
                      BOSS_KOISHI: '古明地こいし'
}
BOSSNAME2STAGE_HASHMAP = {
                          None: '',
                          BOSS_KISUME: '1面',
                          BOSS_YAMAME: '1面',
                          BOSS_PARSEE: '2面',
                          BOSS_YUGI: '3面',
                          BOSS_SATORI: '4面',
                          BOSS_ORIN: '5面',
                          BOSS_UTSUHO: '6面',
                          BOSS_SANAE: 'EX面',
                          BOSS_KOISHI: 'EX面',
}
BOSS_SHORT_NAME_HASHMAP = {
                      None: '',
                      BOSS_KISUME: 'kisume',
                      BOSS_YAMAME: 'yamame',
                      BOSS_PARSEE: 'parsee',
                      BOSS_YUGI: 'yugi',
                      BOSS_SATORI: 'satori',
                      BOSS_ORIN: 'orin',
                      BOSS_UTSUHO: 'utsuho',
                      BOSS_SANAE: 'sanae',
                      BOSS_KOISHI: 'koishi'
}
SPELL_CARD_NAME_HASHMAP = {
                           None: '',
                           'easy': [],
                           'normal': [],
                           'hard': [],
                           'lunatic': [],
                           'extra': [
                                     '秘法「九字刺し」',
                                     '奇跡「ミラクルフルーツ」',
                                     '神徳「五穀豊穣ライスシャワー」',
                                     '表象「夢枕にご先祖総立ち」',
                                     '表象「弾幕パラノイア」',
                                     '本能「イドの解放」',
                                     '抑制「スーパーエゴ」',
                                     '反応「妖怪ポリグラフ」',
                                     '無意識「弾幕のロールシャッハ」',
                                     '復燃「恋の埋火」',
                                     '深層「無意識の遺伝子」',
                                     '「嫌われ者のフィロソフィ」',
                                     '「サブタレイニアンローズ」'
                                     ]
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
BOSS_NAME_ROI = (76, 56, 246, 72)

# ボス残機のROI配列
BOSS_REMAIN_ROIS =[]
width = 18
height = 18
roi_start_position = (72, 74)
padding = 2 # 星と星の間の間隔
for index in range(9):
    left = roi_start_position[0] + (width * index) + (padding * index)
    top = roi_start_position[1]
    right = left + width
    bottom = top + height
    BOSS_REMAIN_ROIS.append((left, top, right, bottom))


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
BINARY_BOSS_NAMES = []
for index in range(9):
    img = cv2.imread(SAMPLE_BOSS_NAMES_DIR + str(index) + '.png')
    img = cv2.inRange(img, (255, 255, 119), (255, 255, 119))
    BINARY_BOSS_NAMES.append(img)

# ボス残機のサンプルデータ(緑色の星画像で固定)
# ボス残機は薄い緑（R:233、G:244、B:225）～濃い緑（R:89、G:172、B:21）なので色を指定して二値化してから使用する
BINARY_BOSS_REMAIN = cv2.imread(SAMPLE_BOSS_REMAINS_DIR + '0.png')
BINARY_BOSS_REMAIN = cv2.inRange(BINARY_BOSS_REMAIN, (21, 172, 89), (225, 244, 233))

# スペルカードのROI
SPELL_CARD_ROI = (340, 64, 820, 97)

# スペルカードのサンプルデータ(難易度を元に動的に切り替え)
BINARY_SPELL_CARDS = []

# スペルカード名情報(BINARY_SPELL_CARDSと同じタイミングで定義)
SPELL_CARD_NAMES = []


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


def ajust_capture_position(rect_left, rect_top, rect_right, rect_bottom):
    # キャプチャ位置修正
    cap_left = rect_left + 3
    cap_top = rect_top + 26
    cap_right = cap_left + 1280
    cap_bottom = cap_top + 960

    return cap_left, cap_top, cap_right, cap_bottom


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
        cv2.imwrite(OUTPUT_DIR + 'per_score' + str(index) + '.png', clopped_frame)

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


def analyze_boss_name(original_frame):
    # ボス名をテンプレートマッチングにより取得
    boss_name = None

    # ボス名は単色（R:119、G:255、B:255）なので色を指定して二値化することで完全一致に近いマッチングが可能
    # 二値化の色指定について
    # OpenCv(cv2)の色の並びはBGRだけど、ここのorifinal_frameはPillowで取得したRGB画像の配列データなのでRGBの順で指定
    boss_name_frame = original_frame[BOSS_NAME_ROI[1]:BOSS_NAME_ROI[3], BOSS_NAME_ROI[0]:BOSS_NAME_ROI[2]]
    boss_name_frame = cv2.inRange(boss_name_frame, (119, 255, 255), (119, 255, 255))
#     Image.fromarray(boss_name_frame).save(OUTPUT_DIR + 'boss_name.png')

    for num, template_img in enumerate(BINARY_BOSS_NAMES):
        res = cv2.matchTemplate(boss_name_frame, template_img, cv2.TM_CCORR_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#         print(min_val, max_val, min_loc, max_loc)
        if (max_val > 0.99):
            boss_name = num
            break

    return boss_name


def analyze_boss_remain(original_frame):
    # ボス残機をテンプレートマッチングにより取得
    boss_remain = None

    # ボス残機は薄い緑（R:233、G:244、B:225）～濃い緑（R:89、G:172、B:21）なので色を指定して二値化する
    # 二値化の色指定について
    # OpenCv(cv2)の色の並びはBGRだけど、ここのorifinal_frameはPillowで取得したRGB画像の配列データなのでRGBの順で指定
    for index, roi in enumerate(BOSS_REMAIN_ROIS):
        boss_remain_frame = original_frame[roi[1]:roi[3], roi[0]:roi[2]]
        boss_remain_frame = cv2.inRange(boss_remain_frame, (89, 172, 21), (233, 244, 225))
#         Image.fromarray(boss_remain_frame).save(OUTPUT_DIR + str(index) + 'boss_remain.png')

        res = cv2.matchTemplate(boss_remain_frame, BINARY_BOSS_REMAIN, cv2.TM_CCORR_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#         print(min_val, max_val, min_loc, max_loc)

        # 画面の左側から探索してるため見つからない場合は誤検出でない限りは次も見つからない(存在しない)ため処理を抜ける
        # ボス残機は1色ではないため星の下に入り込んだ弾幕の色なんかによって類似度が下がりがち
        # そのため少し許容する値を落とす
        # 星の中の一番多い単色とかに変えた方が類似度上がるかもしれない
        if (max_val < 0.8):
            break

        if (boss_remain is None):
            boss_remain = 0
        boss_remain += 1

    return boss_remain


def load_spell_card_binaries(difficulty):
    # スペルカードのサンプルデータ読み込み
    global BINARY_SPELL_CARDS
    global SPELL_CARD_NAMES
    BINARY_SPELL_CARDS = []
    SPELL_CARD_NAMES = []

    # 難易度名取得(easy～extra)
    difficulty_name = ''
    if difficulty in DIFFICULTY_HASHMAP:
        difficulty_name = str.lower(DIFFICULTY_HASHMAP[difficulty])

    # 難易度名で始まるファイルの数を取得
    spell_card_len = len(glob.glob(SAMPLE_SPELL_CARDS_DIR + difficulty_name + '_*.png'))

    # 難易度に対応したスペルカードのサンプルデータのみグレースケールで読み込み
    for index in range(spell_card_len):
        img = cv2.imread(SAMPLE_SPELL_CARDS_DIR + difficulty_name + '_' + str(index).zfill(2) + '.png', cv2.IMREAD_GRAYSCALE)
        BINARY_SPELL_CARDS.append(img)

    # スペルカード名のマップ初期化
    SPELL_CARD_NAMES = SPELL_CARD_NAME_HASHMAP[difficulty_name]

    return BINARY_SPELL_CARDS # app内では使わない(予定)だけど返しておく


def analyze_spell_card(original_frame):
    # スペルカードをテンプレートマッチングにより取得
    spell_card = None
    clopped_frame = original_frame[SPELL_CARD_ROI[1]:SPELL_CARD_ROI[3], SPELL_CARD_ROI[0]:SPELL_CARD_ROI[2]]
    clopped_frame = cv2.cvtColor(clopped_frame, cv2.COLOR_RGB2GRAY)
#     cv2.imwrite(OUTPUT_DIR + 'grayscale_spell_card.png', clopped_frame)

    for num, template_img in enumerate(BINARY_SPELL_CARDS):
        res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
#         print(min_val, max_val, min_loc, max_loc)

        # 閾値は暫定
        if (max_val > 0.85):
            spell_card = num
            break

    return spell_card


def convert_difficulty(difficulty):
    # 数値の難易度を文字列に変換
    return DIFFICULTY_HASHMAP[difficulty]


def convert_boss_name(boss_name):
    # 数値のボス名を文字列に変換
    return BOSS_NAME_HASHMAP[boss_name]


def convert_boss_remain(boss_remain):
    # ボス残機を文字列に変換
    if (boss_remain is None):
        return ''
    return str(boss_remain)


def convert_spell_card(spell_card):
    # スペルカードを文字列に変換
    if (spell_card is not None and spell_card < len(SPELL_CARD_NAMES)):
        return SPELL_CARD_NAMES[spell_card]
    return ''