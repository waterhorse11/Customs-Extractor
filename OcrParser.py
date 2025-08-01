import pdfplumber
import os
import re
import json
from PIL import Image
import numpy as np
from PIL import Image as PILImage
from paddleocr import PaddleOCR
import logging
import multiprocessing
from tqdm import tqdm
import cv2
from CustomsFormCorrector import CustomsFormCorrector

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logging.disable(logging.DEBUG)  # 关闭DEBUG日志的打印
logging.disable(logging.WARNING)  # 关闭WARNING日志的打印
# 为每个工作进程设置的全局OCR实例
# 这是必要的，因为实例(self)本身不能被传递给子进程
_process_ocr_instance = None
_process_ocr_lang = None

class OcrParser:
    def __init__(self, lang='en', use_corrector=False):
        self.lang = lang
        self.use_corrector = use_corrector
        self.logger = logging.getLogger("OcrParser")
        self.progress_queue = None  # 用于向UI报告进度的队列

    @staticmethod
    def _initialize_worker(lang: str):
        """
        为每个工作进程初始化OCR模型。
        这是一个静态方法，以便可以安全地传递给Pool的initializer。
        """
        global _process_ocr_instance, _process_ocr_lang
        if _process_ocr_instance is None or _process_ocr_lang != lang:
            logging.info(f"进程 {os.getpid()}: 初始化语言为 '{lang}' 的OCR模型...")
            _process_ocr_instance = PaddleOCR(use_angle_cls=False, lang=lang, use_gpu=False, use_tensorrt=False, show_log=False)
            _process_ocr_lang = lang
            logging.info(f"进程 {os.getpid()}: OCR模型初始化完成。")

    @staticmethod
    def _process_page_groups_worker(page_data: tuple):
        """
        在工作进程中处理单个页面的所有组。
        这是一个静态方法，以便可以安全地被 pool.imap 调用。
        """
        page_num, img_original, img_scale, cell_coords, groups, original_groups, color_threshold = page_data
        
        logger = logging.getLogger(f"Worker-Page-{page_num+1}")
        logger.info(f"开始在进程 {os.getpid()} 中处理页面 {page_num + 1}...")

        img_data = np.array(img_original)
        page_groups = []

        for group_idx, (start_row, end_row) in enumerate(groups):
            group_cells = cell_coords[start_row:end_row+1]
            group_text_rows = []

            for row_cells in group_cells:
                row_texts = []
                for cell in row_cells:
                    if cell:
                        x0, y0, x1, y1 = cell
                        x0_img, y0_img = int(x0 * img_scale), int(y0 * img_scale)
                        x1_img, y1_img = int(x1 * img_scale), int(y1 * img_scale)
                        
                        cell_img_np_rgb = img_data[y0_img:y1_img, x0_img:x1_img]
                        
                        if cell_img_np_rgb.size > 0:
                            try:
                                # 颜色空间转换：从Pillow的RGB格式转换为OpenCV的BGR格式
                                cell_img_np = cv2.cvtColor(cell_img_np_rgb, cv2.COLOR_RGB2BGR)

                                # --- 图像预处理：将非黑色像素替换为白色 ---
                                # 创建一个图像的副本进行处理
                                processed_img = cell_img_np.copy()
                                
                                # 将BGR图像转换为灰度图以创建阈值掩码
                                gray_img = cv2.cvtColor(cell_img_np, cv2.COLOR_BGR2GRAY)
                                
                                # 找到所有不够黑的像素点 (亮度大于等于阈值)
                                # 这些是我们想要变成白色的区域
                                light_pixels_mask = gray_img >= color_threshold
                                
                                # 将这些不够黑的像素在原彩色图副本中设置为白色
                                processed_img[light_pixels_mask] = [255, 255, 255]

                                # 使用处理后的BGR图像进行OCR识别
                                result = _process_ocr_instance.ocr(processed_img, cls=True)

                                # cv2.imshow("processed_img", processed_img)
                                # cv2.waitKey(0)
                                
                                cell_text = ""
                                if result and len(result) > 0 and result[0]:
                                    texts = [line[1][0] for line in result[0] if line and line[1] and line[1][0].strip()]
                                    cell_text = "\n".join(texts)
                                row_texts.append(cell_text)
                            except Exception as e:
                                logger.error(f"处理单元格时出错: {e}")
                                row_texts.append('')
                        else:
                            row_texts.append('')
                    # else:
                    #     row_texts.append('')
                group_text_rows.append(row_texts)
            
            page_groups.append({
                'group_idx': group_idx + 1,
                'rows': group_text_rows,
                'original_rows': original_groups[group_idx]
            })
            
        return page_num, page_groups

    def extract_group_text(self, pdf_path, output_dir=None, page_numbers=None, group_size=4, lang='en', max_workers=None, save_json=True, color_threshold=10):
        """
        使用PaddleOCR从PDF的表格分组中提取文本。

        参数:
            pdf_path (str): PDF文件路径。
            output_dir (str, optional): 输出目录。默认为PDF同目录下的一个子文件夹。
            page_numbers (list, optional): 要处理的页面列表（0-indexed）。默认为所有页面。
            group_size (int, optional): 每个分组的行数。默认为4。
            lang (str, optional): OCR语言。默认为 'en'。
            max_workers (int, optional): 最大工作进程数。默认为CPU核心数。
            save_json (bool, optional): 是否保存JSON结果。默认为True。
            color_threshold (int, optional): 颜色过滤阈值 (0-255)。低于此值的像素被视为文本。默认为50。
            use_corrector (bool, optional): 是否使用海关表单修正器。默认为False。
        """
        if max_workers is None:
            max_workers = multiprocessing.cpu_count()
        max_workers = min(max_workers, 8)

        if output_dir is None:
            pdf_dir = os.path.dirname(os.path.abspath(pdf_path))
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = os.path.join(pdf_dir, f"{pdf_name}_table_groups_text")
        os.makedirs(output_dir, exist_ok=True)

        if self.use_corrector:
            corrector = CustomsFormCorrector(pdf_path)
        page_data_to_process = []
        with pdfplumber.open(pdf_path) as pdf:
            if page_numbers is None:
                page_numbers = range(len(pdf.pages))
            else:
                page_numbers = [p for p in page_numbers if 0 <= p < len(pdf.pages)]

            for page_num in page_numbers:
                self.logger.info(f"准备第 {page_num + 1} 页数据...")
                page = pdf.pages[page_num]
                
                img = page.to_image(resolution=300)
                
                extracted_tables = page.extract_tables()
                if self.use_corrector:
                    extracted_tables = corrector.correct(page_num, page)
                    self.logger.info(extracted_tables)
                if not extracted_tables or not extracted_tables[0]:
                    self.logger.warning(f"第 {page_num + 1} 页未找到表格。")
                    continue
                
                table_text_data = extracted_tables[0]
                
                start_index = -1
                for i, row in enumerate(table_text_data):
                    first_cell_text = row[0]
                    if isinstance(first_cell_text, str) and (
                        re.match(r'^\d+\n', first_cell_text) or
                        re.search(r'ราย\s*การ', first_cell_text) # 更宽松的泰语匹配
                    ):
                        start_index = i
                        break
                
                if start_index == -1:
                    self.logger.warning(f"第 {page_num + 1} 页未找到报关单项目起始点。")
                    continue
                
                table_objects = page.find_tables()
                if not table_objects:
                    self.logger.warning(f"第 {page_num + 1} 页未找到表格坐标对象。")
                    continue
                table_obj = table_objects[0]

                groups = []
                original_groups = []
                for i in range(start_index, len(table_text_data), group_size):
                    if i + group_size <= len(table_text_data) and table_text_data[i][0] is not None and table_text_data[i][0].strip() != '':
                        groups.append((i, i + group_size - 1))
                        fixed_table_text_data = []
                        for row in table_text_data[i:i+group_size]:
                            for cell in row:
                                if cell is None:
                                    row.remove(cell)
                            fixed_table_text_data.append(row)
                        original_groups.append(fixed_table_text_data)
                
                cell_coords = [row.cells for row in table_obj.rows]
                
                page_data_to_process.append((
                    page_num,
                    img.original,
                    img.scale,
                    cell_coords,
                    groups,
                    original_groups,
                    color_threshold
                ))

                # break

        all_pages_groups = {}
        if page_data_to_process:
            self.logger.info(f"使用 {max_workers} 个进程开始OCR处理...")
            total_pages = len(page_data_to_process)
            
            with multiprocessing.Pool(processes=max_workers, initializer=OcrParser._initialize_worker, initargs=(lang,)) as pool:
                # 使用 imap_unordered 以便在任务完成时立即获得结果，这对于进度更新更及时
                results_iterator = pool.imap_unordered(OcrParser._process_page_groups_worker, page_data_to_process)
                
                # 手动迭代结果并更新进度条
                for i, result in enumerate(results_iterator):
                    page_num, page_groups = result
                    if page_groups:
                        # 对结果进行排序，因为imap_unordered不保证顺序
                        all_pages_groups[page_num] = page_groups
                    
                    # 如果UI传递了进度队列，则更新进度
                    if self.progress_queue:
                        # 计算进度百分比
                        progress_percentage = int(((i + 1) / total_pages) * 100)
                        self.progress_queue.put(progress_percentage)
        
        # 注意：由于我们使用了imap_unordered，如果需要按页面顺序处理结果，
        # 在这里需要对 all_pages_groups 字典按键进行排序。
        # 对于保存为JSON，字典键的顺序通常不重要。
        if save_json and all_pages_groups:
            # 保存每个页面的JSON
            # for page_num, page_groups in all_pages_groups.items():
            #     page_json_path = os.path.join(output_dir, f"page_{page_num + 1}_groups_text.json")
            #     with open(page_json_path, 'w', encoding='utf-8') as f:
            #         json.dump(page_groups, f, ensure_ascii=False, indent=2)
            
            # 保存一个包含所有页面的总JSON文件
            all_json_path = os.path.join(output_dir, "all_pages_groups_text.json")
            with open(all_json_path, 'w', encoding='utf-8') as f:
                json.dump(all_pages_groups, f, ensure_ascii=False, indent=2)
            self.logger.info(f"所有页面的合并结果已保存到: {all_json_path}")
            
        return all_pages_groups

def main():
    """脚本的命令行入口。"""
    # 在Windows上, 'spawn'是更安全的多进程启动方式
    if os.name == 'nt':
        multiprocessing.set_start_method('spawn', force=True)

    import argparse
    parser = argparse.ArgumentParser(description="使用OCR从PDF文档的表格中提取文本。")
    parser.add_argument("pdf_path", help="PDF文件路径。")
    parser.add_argument("-o", "--output", help="输出目录路径。默认为PDF旁边的新建文件夹。")
    parser.add_argument("-p", "--pages", type=int, nargs="+", help="要处理的页码列表 (从1开始)。")
    parser.add_argument("--group-size", type=int, default=4, help="每个分组的行数 (默认: 4)。")
    parser.add_argument("--lang", default="en", help="OCR识别语言 (例如 'en', 'ch', 'th')。默认: 'en'。")
    parser.add_argument("--processes", type=int, default=4, help="工作进程数 (默认: CPU核心数)。")
    parser.add_argument("--no-json", action="store_true", help="不保存JSON输出文件。")
    parser.add_argument("--color-threshold", type=int, default=10, help="颜色过滤的亮度阈值 (0-255)。数值越低，只识别越黑的文本。默认: 10。")
    args = parser.parse_args()

    # 将用户输入的1-based页码转换为0-based
    page_numbers = [p - 1 for p in args.pages] if args.pages else None

    # 初始化并运行解析器
    ocr_parser = OcrParser(lang=args.lang)
    all_pages_groups = ocr_parser.extract_group_text(
        args.pdf_path,
        output_dir=args.output,
        page_numbers=page_numbers,
        group_size=args.group_size,
        lang=args.lang,
        max_workers=args.processes,
        save_json=not args.no_json,
        color_threshold=args.color_threshold
    )
    
    logging.info(f"\n提取完成。共处理了 {len(all_pages_groups)} 个页面。")
    if all_pages_groups and not args.no_json:
        output_dir = args.output or os.path.join(os.path.dirname(os.path.abspath(args.pdf_path)), os.path.splitext(os.path.basename(args.pdf_path))[0] + '_table_groups_text')
        logging.info(f"输出文件已保存到: {output_dir}")
    else:
        logging.info("未生成输出文件。")

if __name__ == "__main__":
    # 这对于打包成可执行文件（如使用PyInstaller）很重要
    multiprocessing.freeze_support()
    main()
