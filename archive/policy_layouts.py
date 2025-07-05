import cv2
from paddleocr import PaddleOCR
import numpy as np

# 範例：針對常見保單格式定義欄位區塊（左上x, 左上y, 右下x, 右下y）
# 單位為像素，需根據實際掃描尺寸微調
POLICY_LAYOUT = {
    'policy_number': (200, 120, 600, 180),
    'insured_name': (200, 200, 600, 260),
    'insurance_company': (100, 60, 500, 110),
    'premium_amount': (900, 600, 1200, 660),
    'coverage_amount': (900, 680, 1200, 740),
    # ...可依實際需求擴充...
}

ocr = PaddleOCR(use_angle_cls=True, lang='ch', show_log=False)

def detect_fields_by_layout(image_path, layout=POLICY_LAYOUT):
    """
    根據欄位區塊設定裁切並辨識各欄位
    Args:
        image_path: 圖片路徑
        layout: 欄位區塊設定
    Returns:
        欄位辨識結果 dict
    """
    image = cv2.imread(image_path)
    results = {}
    for field, (x1, y1, x2, y2) in layout.items():
        crop = image[y1:y2, x1:x2]
        # OCR辨識
        ocr_result = ocr.ocr(crop)
        texts = []
        if ocr_result and ocr_result[0]:
            for line in ocr_result[0]:
                if line and len(line) >= 2:
                    text = line[1][0]
                    conf = line[1][1]
                    texts.append((text, conf))
        # 取信心度最高的結果
        if texts:
            texts.sort(key=lambda x: x[1], reverse=True)
            results[field] = texts[0][0]
        else:
            results[field] = ''
    return results 