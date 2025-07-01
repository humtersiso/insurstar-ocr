import cv2
import pytesseract
import numpy as np
import os
import difflib

# 設定 tesseract 執行檔路徑（Windows 預設安裝路徑）
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def imread_unicode(path):
    with open(path, 'rb') as f:
        img_array = np.asarray(bytearray(f.read()), dtype=np.uint8)
        img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    return img

def find_best_match_line(d, keyword):
    # 將每一行的文字合併，找最接近的行
    lines = {}
    for i, line_num in enumerate(d['line_num']):
        if line_num not in lines:
            lines[line_num] = []
        lines[line_num].append(d['text'][i])
    best_score = 0
    best_line = None
    best_y = None
    for line_num, words in lines.items():
        line_text = ''.join(words).replace(' ', '')
        score = difflib.SequenceMatcher(None, line_text, keyword).ratio()
        if score > best_score:
            best_score = score
            best_line = line_text
            # 取該行第一個字的 top
            idx = [i for i, ln in enumerate(d['line_num']) if ln == line_num][0]
            best_y = d['top'][idx] + d['height'][idx]
    return best_y, best_line, best_score

def crop_table_by_best_match(image_path, output_path, keyword):
    img = imread_unicode(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    d = pytesseract.image_to_data(gray, lang='chi_tra+eng', output_type=pytesseract.Output.DICT)
    y, line, score = find_best_match_line(d, keyword)
    if y is None:
        print(f"找不到接近的關鍵字：{keyword}")
        return
    print(f"最佳匹配行：{line}，相似度：{score}")
    h = img.shape[0] - y
    crop = img[y:y+h, :]
    os.makedirs(output_path, exist_ok=True)
    out_path = os.path.join(output_path, "table_cropped.png")
    cv2.imwrite(out_path, crop)
    print(f"已切割並存檔：{out_path}")

def crop_table_by_coords(image_path, output_path, x, y, w, h):
    img = imread_unicode(image_path)
    crop = img[y:y+h, x:x+w]
    os.makedirs(output_path, exist_ok=True)
    out_path = os.path.join(output_path, "table_cropped.png")
    cv2.imwrite(out_path, crop)
    print(f"已依指定座標切割並存檔：{out_path}")

def batch_crop_tables():
    src_dir = 'converted_images'
    dst_dir = 'static/table_preview'
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    # 統一座標
    crop_x, crop_y, crop_w, crop_h = 85, 1156, 1794, 743
    for fname in os.listdir(src_dir):
        if fname.lower().endswith('.png'):
            src_path = os.path.join(src_dir, fname)
            img = imread_unicode(src_path)
            if img is None:
                print(f'無法讀取圖片: {src_path}')
                continue
            cropped = img[crop_y:crop_y+crop_h, crop_x:crop_x+crop_w]
            base, ext = os.path.splitext(fname)
            out_path = os.path.join(dst_dir, f'{base}_cropped.png')
            cv2.imwrite(out_path, cropped)
            print(f'已切割: {out_path}')

def draw_table_lines():
    src_dir = 'converted_images'
    dst_dir = 'static/table_preview'
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    # 統一座標
    crop_x, crop_y, crop_w, crop_h = 85, 1156, 1794, 743
    n_rows = 6  # 預設行數，可依實際表格調整
    n_cols = 5  # 預設欄數，可依實際表格調整
    for fname in os.listdir(src_dir):
        if fname.lower().endswith('.png'):
            src_path = os.path.join(src_dir, fname)
            img = imread_unicode(src_path)
            if img is None:
                print(f'無法讀取圖片: {src_path}')
                continue
            # 只針對表格區塊處理
            roi = img[crop_y:crop_y+crop_h, crop_x:crop_x+crop_w].copy()
            h, w = roi.shape[:2]
            preview = roi.copy()
            # 畫橫線（行）
            for i in range(n_rows + 1):
                y = int(i * h / n_rows)
                cv2.line(preview, (0, y), (w, y), (0, 0, 255), 2)  # 改為紅色
            # 畫直線（欄）
            for j in range(n_cols + 1):
                x = int(j * w / n_cols)
                cv2.line(preview, (x, 0), (x, h), (0, 0, 255), 2)  # 改為紅色
            base, ext = os.path.splitext(fname)
            out_path = os.path.join(dst_dir, f'{base}_table_lines.png')
            cv2.imwrite(out_path, preview)
            print(f'已產生表格線條預覽: {out_path}')

def ocr_row_lines():
    src_path = 'converted_images/RCE-9915_page_01.png'
    out_path = 'static/table_preview/RCE-9915_page_01_ocr_lines.png'
    img = imread_unicode(src_path)
    if img is None:
        print(f'無法讀取圖片: {src_path}')
        return
    # 統一座標
    crop_x, crop_y, crop_w, crop_h = 85, 1156, 1794, 743
    roi = img[crop_y:crop_y+crop_h, crop_x:crop_x+crop_w].copy()
    h, w = roi.shape[:2]
    preview = roi.copy()
    # 取得每一行的 y 座標
    boxes = pytesseract.image_to_data(roi, lang='chi_tra', output_type=pytesseract.Output.DICT)
    line_tops = []
    for i in range(len(boxes['level'])):
        if boxes['text'][i].strip() != '':
            top = boxes['top'][i]
            height = boxes['height'][i]
            line_tops.append((top, top+height))
    # 合併重疊的行
    line_tops.sort()
    merged_lines = []
    for t, b in line_tops:
        if not merged_lines or t > merged_lines[-1][1]:
            merged_lines.append([t, b])
        else:
            merged_lines[-1][1] = max(merged_lines[-1][1], b)
    # 畫紅色橫線
    for t, b in merged_lines:
        cv2.line(preview, (0, t), (w, t), (0, 0, 255), 2)
        cv2.line(preview, (0, b), (w, b), (0, 0, 255), 2)
    cv2.imwrite(out_path, preview)
    print(f'已產生 OCR 行分割預覽: {out_path}')

def projection_row_lines():
    src_path = 'converted_images/RCE-9915_page_01.png'
    out_path = 'static/table_preview/RCE-9915_page_01_proj_lines.png'
    img = imread_unicode(src_path)
    if img is None:
        print(f'無法讀取圖片: {src_path}')
        return
    # 統一座標
    crop_x, crop_y, crop_w, crop_h = 85, 1156, 1794, 743
    roi = img[crop_y:crop_y+crop_h, crop_x:crop_x+crop_w].copy()
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    # 橫向投影（行分割）
    projection = np.sum(binary, axis=1)
    threshold = np.max(projection) * 0.2
    in_text = False
    lines = []
    for y, val in enumerate(projection):
        if val > threshold and not in_text:
            start = y
            in_text = True
        elif val <= threshold and in_text:
            end = y
            lines.append((start, end))
            in_text = False
    preview = roi.copy()
    h, w = roi.shape[:2]
    for t, b in lines:
        cv2.line(preview, (0, t), (w, t), (0, 0, 255), 2)
        cv2.line(preview, (0, b), (w, b), (0, 0, 255), 2)
    # 縱向投影（欄分割）
    v_projection = np.sum(binary, axis=0)
    v_threshold = np.max(v_projection) * 0.2
    in_col = False
    v_lines = []
    for x, val in enumerate(v_projection):
        if val > v_threshold and not in_col:
            start = x
            in_col = True
        elif val <= v_threshold and in_col:
            end = x
            v_lines.append((start, end))
            in_col = False
    for l, r in v_lines:
        cv2.line(preview, (l, 0), (l, h), (0, 0, 255), 2)
        cv2.line(preview, (r, 0), (r, h), (0, 0, 255), 2)
    cv2.imwrite(out_path, preview)
    print(f'已產生投影法行+欄分割預覽: {out_path}')

def ocr_col_lines():
    src_path = 'converted_images/RCE-9915_page_01.png'
    out_path = 'static/table_preview/RCE-9915_page_01_ocr_cols.png'
    img = imread_unicode(src_path)
    if img is None:
        print(f'無法讀取圖片: {src_path}')
        return
    # 統一座標
    crop_x, crop_y, crop_w, crop_h = 85, 1156, 1794, 743
    roi = img[crop_y:crop_y+crop_h, crop_x:crop_x+crop_w].copy()
    h, w = roi.shape[:2]
    preview = roi.copy()
    # OCR 取得所有文字座標
    boxes = pytesseract.image_to_data(roi, lang='chi_tra', output_type=pytesseract.Output.DICT)
    keywords = ['保險種類', '保險金額', '自負額', '簽單保費']
    x_positions = []
    for i, text in enumerate(boxes['text']):
        t = text.strip().replace(' ', '')
        for kw in keywords:
            if kw in t:
                x = boxes['left'][i]
                x_positions.append(x)
    # 去除重複並排序
    x_positions = sorted(set(x_positions))
    # 畫紅色直線
    for x in x_positions:
        cv2.line(preview, (x, 0), (x, h), (0, 0, 255), 2)
    cv2.imwrite(out_path, preview)
    print(f'已產生 OCR 關鍵字欄分割預覽: {out_path}')

if __name__ == "__main__":
    batch_crop_tables()
    draw_table_lines()
    ocr_row_lines()
    projection_row_lines()
    ocr_col_lines()
    print('批次切割、表格線條、OCR行分割、投影法行分割與OCR欄分割預覽完成！')

    img_path = "converted_images/RAS-5879_sample.png"
    out_dir = "static/table_preview"
    # 直接用 sample 座標切割（請根據實際需求調整座標）
    crop_table_by_coords(img_path, out_dir, x=0, y=1200, w=2481, h=800)

    # 測試用
    img = imread_unicode(img_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray, lang='chi_tra+eng')
    print("=== OCR 文字內容 ===")
    print(text)
    print("==================")
    # 你可以根據這些內容選擇正確的關鍵字
    # crop_table_by_keywords(img_path, out_dir, '承保內容', '續下頁，投保駕駛人傷害險') 