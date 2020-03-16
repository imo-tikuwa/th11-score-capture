
import cv2

def edit_frame(frame):
    # フレームを二値化
    work_frame = frame

    work_frame = cv2.cvtColor(work_frame, cv2.COLOR_RGB2GRAY)
#     work_frame = cv2.threshold(work_frame, FRAME_THRESH, 255, cv2.THRESH_BINARY)[1]
#     work_frame = cv2.bitwise_not(work_frame)

    return work_frame

def analyze_score(SCORE_ROIS, BINARY_NUMBERS, work_frame):
    # スコアをテンプレートマッチングにより取得
    score_results = []
    for index, roi in enumerate(SCORE_ROIS):
        clopped_frame = work_frame[roi[1]:roi[3], roi[0]:roi[2]]
#         file_name = OUTPUT_DIR + current_time + '_score_' + str(index) + '.png'
#         cv2.imwrite(file_name, clopped_frame)

        for num, template_img in enumerate(BINARY_NUMBERS):
            res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            if (max_val > 0.99):
                score_results.append(num)
                break

    score = "".join(map(str, score_results))

    return score

def analyze_remain(REMAIN_ROIS, BINARY_REMAINS, work_frame):
    # 残機をテンプレートマッチングにより取得
    remain = 0
    for index, roi in enumerate(REMAIN_ROIS):
        clopped_frame = work_frame[roi[1]:roi[3], roi[0]:roi[2]]
#         file_name = OUTPUT_DIR + current_time + '_remain_' + str(index) + '.png'
#         cv2.imwrite(file_name, clopped_frame)

        temp_max_val = current_remain = 0;
        for num, template_img in enumerate(BINARY_REMAINS):
            res = cv2.matchTemplate(clopped_frame, template_img, cv2.TM_CCORR_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
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
