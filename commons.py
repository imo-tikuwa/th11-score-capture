
import os
import cv2
from datetime import datetime
from PIL import Image

# 変数(定数扱いする変数)
TH11_WINDOW_NAME = '東方地霊殿　～ Subterranean Animism. ver 1.00a'
CONFIG_FILE_NAME = 'settings.ini'
OUTPUT_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'output' + os.sep
SAMPLE_DIR = os.path.abspath(os.path.dirname(__file__)) + os.sep + 'sample_data' + os.sep
SAMPLE_NUMBERS_DIR = SAMPLE_DIR + 'number' + os.sep
SAMPLE_REMAINS_DIR = SAMPLE_DIR + 'remain' + os.sep
SAMPLE_DIFFICULTIES_DIR = SAMPLE_DIR + 'difficulty' + os.sep
DIFFICULTY_HASHMAP = {
                      None: '',
                      0: 'Easy',
                      1: 'Normal',
                      2: 'Hard',
                      3: 'Lunatic',
                      4: 'Extra'
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
