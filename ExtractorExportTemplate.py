from FieldsExtractor import ExportFields, ExportFieldsExtractor
import re

class TianShiExtractor(ExportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en'):
        """
        天狮模板的构造函数。
        调用父类的构造函数以完成基本初始化。
        """
        super().__init__(pdf_path, output_dir, lang)
        self.logger.info("初始化 TianShi 模板提取器。")

    def _parse_group_to_fields(self, group_data: dict) -> ExportFields:
        """
        重写父类的方法以适应TianShi模板的特定字段布局。
        （如果布局与基类完全相同，则此方法可以不重写）
        """
        self.logger.info("使用 TianShi 特定的解析逻辑...")
        # 调用父类的解析方法作为基础，然后可以进行微调
        item = super()._parse_group_to_fields(group_data)
        
        # --- 针对TianShi模板的特定调整 ---
        # 示例：如果TianShi模板的型号在不同位置
        # item.MODEL = self._get_cell_from_rows(group_data['rows'], 3, 2) 
        
        return item

class HlsExtractor(ExportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en'):
        """
        HLS模板的构造函数。
        """
        super().__init__(pdf_path, output_dir, lang)
        self.logger.info("初始化 HLS 模板提取器。")

    def _parse_group_to_fields(self, group_data: dict) -> ExportFields:

        def get_original_cell(r, c, default=''):
            try:
                return group_data.get('original_rows', [])[r][c].strip()
            except (IndexError, AttributeError):
                return default
            
        # 调用父类的解析方法作为基础，然后可以进行微调
        item = super()._parse_group_to_fields(group_data)

        item.AMOUNT_USD = self.get_digital_value(item.AMOUNT_USD)
        item.AMOUNT_THB = self.get_digital_value(item.AMOUNT_THB)

        # 重量cell
        if len(item.WEIGHT_UNIT.split('\n')) > 1:
            total_nw_text = item.WEIGHT_UNIT.split('\n')[1]
            nw_match = re.match(r'([0-9,.]+)\s*([A-Z]+)', total_nw_text, re.IGNORECASE)
            if nw_match:
                item.WEIGHT_UNIT = nw_match.group(2).replace('KGV', 'KGM')
                item.TOTAL_N_W = nw_match.group(1)
            else:
                item.WEIGHT_UNIT = total_nw_text

        item.NO = self.get_digital_value(get_original_cell(0, 0))

        return item


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="从PDF报关单中提取结构化字段。")
    parser.add_argument("pdf_path", help="PDF文件路径。")
    parser.add_argument("-o", "--output", help="输出目录路径。默认为PDF旁边的新建文件夹。")
    parser.add_argument("--lang", default="en", help="OCR识别语言 (例如 'en', 'ch', 'th')。默认: 'en'。")
    parser.add_argument("--template", default="TianShi", help="模板类型 (例如 'TianShi', 'LSS', 'HLS')。默认: 'TianShi'。")
    args = parser.parse_args()
    

    if args.template == "TianShi":
        extractor = TianShiExtractor(args.pdf_path, args.output, args.lang)
    elif args.template == "HLS":
        extractor = HlsExtractor(args.pdf_path, args.output, args.lang)

    # 执行提取流程
    extractor.extract_items()