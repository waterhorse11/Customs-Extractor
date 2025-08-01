import fitz  # PyMuPDF
import pdfplumber
# import pprint
from difflib import SequenceMatcher
from PIL import Image
import io
import re
# import pytesseract

# pytesseract.pytesseract.tesseract_cmd = r"D:\Program Files\Tesseract-OCR\tesseract.exe"

class CustomsFormCorrector:
    """
    一个用于修正海关表单中特定字符错误的类。

    该类通过结合 PyMuPDF (fitz) 和 pdfplumber 的优点，
    实现了一个"精确换砖"的逻辑：
    1. 使用 pdfplumber 获取准确的表格结构（地基）。
    2. 使用 fitz 获取包含可修正字符的文本内容（好砖）。
    3. 通过坐标定位，用"好砖"替换掉"地基"中对应的"坏砖"，
       从而实现对PDF表格的精确、最小化修正。
    """
    def __init__(self, pdf_path, char_to_find='\x15', char_to_replace='2'):
        """
        初始化修正器。
        :param pdf_path: 要处理的PDF文件路径。
        :param char_to_find: 在fitz提取的文本中要查找的特殊字符。
        :param char_to_replace: 用于替换特殊字符的正确字符。
        """
        # print("[INFO] 初始化 CustomsFormCorrector...")
        self.pdf_path = pdf_path
        self.char_to_find = char_to_find
        self.char_to_replace = char_to_replace

        self.fitz_doc = fitz.open(self.pdf_path)
        self.fitz_page = self.fitz_doc[0]

        self.correction_count = 0
        print("[INFO] 初始化完成。")

    def _get_fitz_blocks(self, page_num):
        """
        (私有方法) 使用 PyMuPDF 的 get_text('blocks') 提取所有修正后的文本块。
        """
        # print("[INFO] 步骤 1/3: 从 PyMuPDF 提取修正后的文本块 ('好砖')...")
        self.fitz_page = self.fitz_doc[page_num]
        blocks_data = []
        for block in self.fitz_page.get_text("blocks"):
            x0, y0, x1, y1, text, _, _ = block
            corrected_text = text.strip().replace(self.char_to_find, self.char_to_replace)
            if corrected_text:
                blocks_data.append({"text": corrected_text, "bbox": (x0, y0, x1, y1)})
        # print(f"[INFO] 找到 {len(blocks_data)} 个 '好砖'.")
        return blocks_data

    def correct(self, page_num, plumber_page):
        """
        执行核心的"精确换砖"修正流程。
        :return: 修正后的、与 pdfplumber.extract_tables() 格式相同的表格数据。
        """
        self.plumber_page = plumber_page
        # 步骤1: 获取"好砖"
        fitz_blocks = self._get_fitz_blocks(page_num)

        # 步骤2: 获取"地基"
        # print("[INFO] 步骤 2/3: 从 pdfplumber 提取原始表格结构和内容 ('地基')...")
        plumber_tables = self.plumber_page.find_tables()
        if not plumber_tables:
            print("错误：pdfplumber 未能在此页面上找到任何表格。")
            return None
        
        # 将原始表格数据提取出来，作为我们最终要修改的"画布"
        final_tables_rebuilt = [table.extract() for table in plumber_tables]
        final_tables_changed_count = [[[0 for _ in range(len(table.columns))] for _ in range(len(table.rows))] for table in plumber_tables]
        final_tables_ocr_text = [[["" for _ in range(len(table.columns))] for _ in range(len(table.rows))] for table in plumber_tables]

        # 步骤3: "精确换砖"
        # print("[INFO] 步骤 3/3: 开始 '精确换砖' 流程...")
        for fitz_block in fitz_blocks:
            # if self.char_to_replace not in fitz_block['text']:
            if not bool(re.search(r'\d', fitz_block['text'])):
                continue
            # print('fitz_block:',fitz_block['text'])
            # 定位"好砖"所属的单元格
            block_center_x = (fitz_block['bbox'][0] + fitz_block['bbox'][2]) / 2
            block_center_y = (fitz_block['bbox'][1] + fitz_block['bbox'][3]) / 2
            
            found_cell_coords = None
            for t_idx, table in enumerate(plumber_tables):
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell_bbox in enumerate(row.cells):
                        if (cell_bbox and
                            cell_bbox[0] <= block_center_x < cell_bbox[2] and
                            cell_bbox[1] <= block_center_y < cell_bbox[3]):

                            # if iou > 0.95:
                            found_cell_coords = (t_idx, r_idx, c_idx)
                            # print(fitz_block['text'])
                            # if '1, 00.000 C6' == fitz_block['text']:
                                # print("fitz_block", fitz_block['text'])
                            #     clip = fitz.Rect(cell_bbox[0], cell_bbox[1], cell_bbox[2], cell_bbox[3])
                            #     pix = self.fitz_page.get_pixmap(clip=clip, dpi=300)
                            #     show_image = Image.open(io.BytesIO(pix.tobytes()))
                            #     show_image.show()
                            break
                    if found_cell_coords: break
                if found_cell_coords: break
            
            if not found_cell_coords: continue

            t, r, c = found_cell_coords
            
            # bad_brick_text = self.plumber_page.crop(fitz_block['bbox']).extract_text()
            # original_cell_text = final_tables_rebuilt[t][r][c]
            
            good_brick_text = fitz_block['text']
            corrected_text = good_brick_text
            
            if final_tables_changed_count[t][r][c] == 0:
                final_tables_rebuilt[t][r][c] = corrected_text
            else:
                if bool(re.fullmatch(r'^[\x20-\x7E]*$', corrected_text)) and len(corrected_text) < 20:
                    # if final_tables_ocr_text[t][r][c] == "":
                    #     pix = self.fitz_page.get_pixmap(clip=cell_bbox, dpi=300)
                    #     img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    #     img.show()
                    #     try:
                    #         ocr_text = pytesseract.image_to_string(img, lang="eng").strip()
                    #         final_tables_ocr_text[t][r][c] = ocr_text
                    #     except pytesseract.TesseractError:
                    #         ocr_text = ""

                    
                    # matcher = SequenceMatcher(None, corrected_text, final_tables_ocr_text[t][r][c])
                    # similarity = matcher.ratio()
                    # if similarity > 0.8:
                    #     final_tables_rebuilt[t][r][c] = corrected_text

                    # else:
                    #     final_tables_rebuilt[t][r][c] = final_tables_ocr_text[t][r][c]
                    final_tables_rebuilt[t][r][c] = corrected_text

                    
                else:
                    if 'KGM' in final_tables_rebuilt[t][r][c] or 'KGM' in corrected_text:
                        final_tables_rebuilt[t][r][c] += '\n' + corrected_text
                    else:
                        final_tables_rebuilt[t][r][c] = corrected_text
            
            self.correction_count += 1
            final_tables_changed_count[t][r][c] += 1


            # print(f'{original_cell_text} -> {corrected_text}')
            # print("==========================")
        
        # print(f"[INFO] '精确换砖' 流程完成。")
        return final_tables_rebuilt
    
    def clean_string(self, text):
        """仅用于文本匹配的清洗函数，移除空格和换行符"""
        if not isinstance(text, str): return ""
        return text.replace(" ", "").replace("\n", "")

    def calculate_change_ratio(self, matcher, original_len):
        """计算需要改动的字符数占原始字符串总长度的比例"""
        if original_len == 0: return 1.0
        changed_chars = 0
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag != 'equal':
                changed_chars += (i2 - i1)
        return changed_chars / original_len
    
    def compute_small_in_large_ratio(self, rect_small, rect_large):
        # rect: (x1, y1, x2, y2)
        x1, y1, x2, y2 = rect_small
        x3, y3, x4, y4 = rect_large

        # 计算交集坐标
        xi1 = max(x1, x3)
        yi1 = max(y1, y3)
        xi2 = min(x2, x4)
        yi2 = min(y2, y4)

        # 计算交集宽高
        inter_width = max(0, xi2 - xi1)
        inter_height = max(0, yi2 - yi1)
        inter_area = inter_width * inter_height

        # 计算小矩形面积
        area_small = max(0, x2 - x1) * max(0, y2 - y1)

        if area_small == 0:
            return 0.0
        return inter_area / area_small

    def patch_text(self, original_text, ocr_text, char_to_patch='2'):
        """
        [新] 更智能的文本合并函数。
        只在必要时插入或替换为目标字符，并尽量保留原始格式。
        """
        original_cleaned = self.clean_string(original_text)
        ocr_cleaned = self.clean_string(ocr_text)
        matcher = SequenceMatcher(None, original_cleaned, ocr_cleaned)
        
        opcodes = matcher.get_opcodes()
        result_str = list(original_text) 
        
        # 创建从清洗后文本索引到原始文本索引的映射
        original_map = [i for i, char in enumerate(original_text) if not char.isspace()]
        
        # 从后往前应用更改，避免索引失效
        for tag, i1, i2, j1, j2 in reversed(opcodes):
            if tag == 'replace':
                ocr_segment = ocr_cleaned[j1:j2]
                # 只有当替换的片段包含目标字符，并且长度相同时，我们才进行操作
                # 这是一个启发式规则，可以防止错误的替换
                if char_to_patch in ocr_segment and len(original_cleaned[i1:i2]) == len(ocr_segment):
                    # 遍历这个片段，只替换非空格字符
                    for k in range(len(ocr_segment)):
                        # 找到原始字符串中需要被替换的字符的实际索引
                        original_char_index = original_map[i1 + k]
                        result_str[original_char_index] = ocr_segment[k]

        return "".join(result_str)

    def close(self):
        """关闭所有打开的PDF文件句柄。"""
        self.fitz_doc.close()
        # print("[INFO] PDF文件句柄已关闭。")

# --- 示例用法 ---
if __name__ == "__main__":
    # PDF文件路径
    pdf_file_path = r"E:\Code\2025\Customs-Extractor\template\OLC报关单模板.pdf"
    
    # 1. 创建修正器实例
    plumber_page = pdfplumber.open(pdf_file_path).pages[1]
    corrector = CustomsFormCorrector(pdf_path=pdf_file_path)
    
    # 2. 执行修正
    final_tables = corrector.correct(page_num=1, plumber_page=plumber_page)
    
    # 3. 关闭文件
    corrector.close()

    # --- 结果展示 ---
    print(f"\n\n======= 修正完成，共计 {corrector.correction_count} 处 '换砖' =======")
    if final_tables:
        # 打印第一个表格作为示例
        print(final_tables[0])
    else:
        print("未能解析到任何表格。")