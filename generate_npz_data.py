import numpy
import os
import time
import cv2
from datetime import datetime
from PIL import Image, ImageGrab
from builtins import enumerate
import glob


# 変数(定数扱いする変数)
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'output' + os.sep
SAMPLE_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'sample_data' + os.sep
SAMPLE_NUMBERS_DIR = SAMPLE_DIR + 'number' + os.sep
SAMPLE_REMAINS_DIR = SAMPLE_DIR + 'remain' + os.sep
SAMPLE_DIFFICULTIES_DIR = SAMPLE_DIR + 'difficulty' + os.sep
SAMPLE_BOSS_NAMES_DIR = SAMPLE_DIR + 'boss_name' + os.sep
SAMPLE_BOSS_REMAINS_DIR = SAMPLE_DIR + 'boss_remain' + os.sep
SAMPLE_SPELL_CARDS_DIR = SAMPLE_DIR + 'spell_card' + os.sep
SAMPLE_STAGE_CLEARS_DIR = SAMPLE_DIR + 'stage_clear' + os.sep
SAMPLE_ENEMY_ICONS_DIR = SAMPLE_DIR + 'enemy_icon' + os.sep
SAMPLE_TIME_REMAINS_DIR = SAMPLE_DIR + 'time_remain' + os.sep
NPZ_FILE = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'npz_data' + os.sep + 'bundle.npz'
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
img = cv2.imread(SAMPLE_BOSS_REMAINS_DIR + '0.png')
BINARY_BOSS_REMAIN = cv2.inRange(img, (21, 172, 89), (225, 244, 233))

# スペルカードのサンプルデータ(難易度を元に動的に切り替え)
BINARY_SPELL_CARDS = [[],[],[],[],[]]
for index, difficulty in enumerate(DIFFICULTY_HASHMAP):
    if (index is None):
        continue
    # 難易度名
    difficulty_name = str.lower(DIFFICULTY_HASHMAP[difficulty])

    # 難易度名で始まるファイルの数を取得
    spell_card_len = len(glob.glob(SAMPLE_SPELL_CARDS_DIR + difficulty_name + '_*.png'))

    # 難易度に対応したスペルカードのサンプルデータのみグレースケールで読み込み
    for index in range(spell_card_len):
        img = cv2.imread(SAMPLE_SPELL_CARDS_DIR + difficulty_name + '_' + str(index).zfill(2) + '.png', cv2.IMREAD_GRAYSCALE)
        BINARY_SPELL_CARDS[difficulty].append(img)


# ステージクリアのサンプルデータ(1～6面およびリプレイ再生時のALLクリア用の2種類)
# 0.pngは1～5面用、精度を上げるため赤（R:255、G:0、B:0）～濃い赤（R:136、G:0、B:0）の範囲で二値化する
BINARY_STAGE_CLEARS = []
img = cv2.imread(SAMPLE_STAGE_CLEARS_DIR + '0.png')
img = cv2.inRange(img, (0, 0, 136), (0, 0, 255))
# cv2.imwrite(OUTPUT_DIR + 'stage_clear0.png', img)
BINARY_STAGE_CLEARS.append(img)
# 1.pngリプレイ再生時のALLクリア用、特に弄らずにグレースケール化したものと一致するみたい
img = cv2.imread(SAMPLE_STAGE_CLEARS_DIR + '1.png', cv2.IMREAD_GRAYSCALE) #グレースケールで読み込み
BINARY_STAGE_CLEARS.append(img)


# エネミーアイコンのサンプルデータ
BINARY_ENEMY_ICONS = []
for index in range(2):
    img = cv2.imread(SAMPLE_ENEMY_ICONS_DIR + str(index) + '.png', cv2.IMREAD_GRAYSCALE) #グレースケールで読み込み
    BINARY_ENEMY_ICONS.append(img)


# 残り時間のサンプルデータ
BINARY_TIME_REMAINS = []
for index in range(10):
    img = cv2.imread(SAMPLE_TIME_REMAINS_DIR + str(index) + '.png', cv2.IMREAD_GRAYSCALE) #グレースケールで読み込み
    BINARY_TIME_REMAINS.append(img)


# npz保存
numpy.savez_compressed(NPZ_FILE,
            number = BINARY_NUMBERS,
            remain = BINARY_REMAINS,
            difficulty = BINARY_DIFFICULTIES,
            boss_name = BINARY_BOSS_NAMES,
            boss_remain = BINARY_BOSS_REMAIN,
            spell_card = BINARY_SPELL_CARDS,
            stage_clear = BINARY_STAGE_CLEARS,
            enemy_icon = BINARY_ENEMY_ICONS,
            time_remain = BINARY_TIME_REMAINS
            )
