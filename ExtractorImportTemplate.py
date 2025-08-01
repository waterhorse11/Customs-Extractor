from FieldsExtractor import ImportFieldsExtractor, ImportFields
import re

class TianShiExtractor(ImportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en'):
        """
        天狮模板的构造函数。
        调用父类的构造函数以完成基本初始化。
        """
        super().__init__(pdf_path, output_dir, lang)
        self.logger.info("初始化 TianShi 模板提取器。")

    def _parse_group_to_fields(self, group_data: dict) -> ImportFields:
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


class LssExtractor(ImportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en'):
        """
        LSS模板的构造函数。
        """
        super().__init__(pdf_path, output_dir, lang)
        self.logger.info("初始化 LSS 模板提取器。")

    def _parse_description_block(self, item: ImportFields, original_block_text: str, ocr_block_text: str):
        """
        使用正则表达式解析复杂的描述信息块（通常在单元格(2,7)）。
        """
        if not original_block_text:
            return

        remaining_block = original_block_text

        
        # 提取原产国 (例如 CN, TH, US)
        country_match = re.search(r'\b(CN|TH|US|JP|DE|VN)\b', remaining_block)
        if country_match:
            item.COUNTRY_OF_ORIGIN = country_match.group(1)
            # 避免将 'CN' 作为描述的一部分
            remaining_block = remaining_block.replace(item.COUNTRY_OF_ORIGIN, '', 1)

        # 提取使用规则 (例如 Origin Criteria PE)
        usage_match = re.search(r'Origin Criteria\s*(.*)|OriginCer\s*(.*)', remaining_block, re.IGNORECASE)
        if usage_match:
            item.USAGE_RULES = usage_match.group(0).strip()
            remaining_block = remaining_block.replace(usage_match.group(0), '', 1).strip()

        # 提取型号
        remaining_text_list = list(filter(lambda x: "BRAND" not in x.upper() and len(x.strip()) > 0 
                                and "ชนิดของ" not in x , remaining_block.replace('null', '').split('\n')))
        for text in remaining_text_list:
            if 'MODEL' in text:
                # 检查是否只有大写字母和数字-和/
                match = re.search(r'[A-Z0-9-/]+$', text.strip())
                if match:
                    item.MODEL = match.group(0).strip()
                    remaining_text_list.remove(text)
                    break
        
        # 提取证书号和日期
        # 示例 FE Cert: 7PEB191001 10/07/2568, TISI: E256183384090202 10/07/2568
        for text in remaining_text_list:
            date_match = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
            self.logger.info(text)
            if date_match:
                item.FE_CERTIFICATE_NO_DATE = text.strip()
                remaining_text_list.remove(text)
                break
        
        # 泰文描述
        for text in remaining_text_list:
            # 是否含有泰文
            if re.search(r'[\u0E00-\u0E7F]', text):
                item.DESCRIPTION_TH = self.replace_pua_thai(text.strip()).replace('null', '').strip()
                remaining_text_list.remove(text)
                break

        # 将剩余部分作为描述
        item.DESCRIPTION = ' '.join(remaining_text_list).strip().replace('"', '').replace('null', '').strip()
        # 删除Permit No及其之后部分
        match = re.search(r'Permit No.*', item.DESCRIPTION)
        if match:
            item.DESCRIPTION = item.DESCRIPTION.replace(match.group(0), '').strip()
        match = re.search(r'FORM E NO', item.DESCRIPTION)
        if match:
            item.DESCRIPTION = item.DESCRIPTION.replace(match.group(0), '').strip()


class HlsExtractor(ImportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en'):
        """
        HLS模板的构造函数。
        """
        super().__init__(pdf_path, output_dir, lang)
        self.logger.info("初始化 HLS 模板提取器。")


class OlcExtractor(ImportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en', use_corrector: bool = False):
        """
        OLC模板的构造函数。
        """
        super().__init__(pdf_path=pdf_path, output_dir=output_dir, lang=lang, use_corrector=use_corrector)
        self.logger.info("初始化 OLC 模板提取器。")

    def _parse_description_block(self, item: ImportFields, original_block_text: str, ocr_block_text: str):
        if not original_block_text:
            return

        remaining_block = original_block_text
        
        # 提取原产国 (例如 CN, TH, US)
        country_match = re.search(r'\b(CN|TH|US|JP|DE|VN)\b', remaining_block)
        if country_match:
            item.COUNTRY_OF_ORIGIN = country_match.group(1)
            # 避免将 'CN' 作为描述的一部分
            remaining_block = remaining_block.replace(item.COUNTRY_OF_ORIGIN, '', 1)

        # 提取使用规则 (例如 Origin Criteria PE)
        usage_match = re.search(r'Origin Criteria\s*(.*)', remaining_block, re.IGNORECASE)
        if usage_match:
            item.USAGE_RULES = usage_match.group(0).strip()
            remaining_block = remaining_block.replace(usage_match.group(0), '', 1).strip()

        # 提取型号
        remaining_text_list = list(filter(lambda x: "BRAND" not in x.upper() and len(x.strip()) > 0 and "ชนิดของ" not in x, remaining_block.split('\n')))
        for text in remaining_text_list:
            # 检查是否只有大写字母和数字-和/
            if re.match(r'^[A-Z0-9-/]+$', text):
                item.MODEL = text.strip()
                remaining_text_list.remove(text)
                break
        
        # 提取证书号和日期
        # 示例 FE Cert: 7PEB191001 10/07/2568, TISI: E256183384090202 10/07/2568
        for text in remaining_text_list:
            date_match = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
            self.logger.info(text)
            if date_match:
                item.FE_CERTIFICATE_NO_DATE = text.strip()
                remaining_text_list.remove(text)
                break
        
        # 泰文描述
        for text in remaining_text_list:
            # 是否含有泰文
            if re.search(r'[\u0E00-\u0E7F]', text):
                item.DESCRIPTION_TH = self.replace_pua_thai(text.strip().replace('r', ' ์'))
                remaining_text_list.remove(text)
                break

        # 将剩余部分作为描述
        item.DESCRIPTION = ' '.join(remaining_text_list).strip().replace('"', '')

class SnpExtractor(ImportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en', use_corrector: bool = False):
        """
        SNP模板的构造函数。
        """
        super().__init__(pdf_path=pdf_path, output_dir=output_dir, lang=lang, use_corrector=use_corrector)
        self.logger.info("初始化 SNP 模板提取器。")


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
    elif args.template == "LSS":
        extractor = LssExtractor(args.pdf_path, args.output, args.lang)
    elif args.template == "HLS":
        extractor = HlsExtractor(args.pdf_path, args.output, args.lang)
    elif args.template == "OLC":
        extractor = OlcExtractor(args.pdf_path, args.output, args.lang, use_corrector=True)
    elif args.template == "SNP":
        extractor = SnpExtractor(args.pdf_path, args.output, args.lang)

    # 执行提取流程
    extractor.extract_items()
