import json
import re
import os
from openpyxl import Workbook


from OcrParser import OcrParser


class ImportFields:
    """一个数据类，用于存放从报关单单个项目中提取的字段。"""
    def __init__(self):
        self.NO = ''    # NO (项号)
        self.MODEL = ''    # MODEL (型号)
        self.DESCRIPTION = ''    # DESCRIPTION(英文描述)
        self.DESCRIPTION_TH = ''    # DESCRIPTION(泰文描述)
        self.HS_CODE = ''    # HS CODE (海关编码)
        self.QTY = ''    # QTY (数量)
        self.QTY_UNIT = ''    # QTY Unit (数量单位)
        self.UNIT_CODE_1 = ''    # Unit Code 1 (单位代码1)
        self.UNIT_CODE_2 = ''    # Unit Code 2 (单位代码2)
        self.PRIVILEGE_CODE = ''   # Privilege Code (优惠代码)
        self.AMOUNT_USD = ''   # AMOUNT(USD) (美元金额)
        self.AMOUNT_THB = ''   # AMOUNT(THB) (泰铢金额)
        self.TOTAL_N_W = ''   # TOTAL.N.W (总净重)
        self.WEIGHT_UNIT = ''   # Weight unit (重量单位)
        self.TAX_RATE = ''   # Tax rate (进口税率)
        self.CUSTOMS_DUTIES_PAYABLE = ''   # Customs duties payable (应缴关税)
        self.DUTY_PAID = ''   # Duty paid (已缴关税)
        self.INV = ''   # Inv. (发票号)
        self.FEE = ''   # Fee (费用)
        self.EXCISE_PRODUCT_CODE = ''   # Excise Product Code (消费税产品代码)
        self.EXCISE_TAX_RATE = ''   # Excise tax rate (消费税率)
        self.EXCISE_TAX = ''   # Excise tax (消费税额)
        self.OTHER_TAXES = ''   # Other taxes (其他税费)
        self.MINISTRY_OF_INTERIOR_TAX = ''   # Taxes for the Ministry of Interior (内政部税)
        self.VALUE_ADDED_TAX_BASE = ''   # Value Added Tax Base (增值税基础)
        self.VAT = ''   # VAT (增值税)
        self.FE_CERTIFICATE_NO_DATE = ''   # FE Certificate No./Date (FE证书号/日期)
        self.TISI_CERTIFICATE_NO_DATE = ''   # TISI Certificate No./Date (TISI证书号/日期)
        self.EXPLANATION = ''   # Explanation (解释说明)
        self.COUNTRY_OF_ORIGIN = ''   # Country of Origin (原产国)
        self.USAGE_RULES = ''   # Usage Rules (使用规则)
        
class ImportFieldsExtractor:
    """
    负责从OCR解析后的文本中提取结构化字段。
    此类不存储字段，而是生成一个包含多个Fields对象的列表。
    """
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en', save_json: bool = False, save_excel: bool = True, use_corrector: bool = False):
        self.pdf_path = pdf_path
        self.output_dir = output_dir if output_dir else self._get_default_output_dir()
        self.lang = lang
        self.save_json = save_json
        self.use_corrector = use_corrector
        self.ocr_parser = OcrParser(lang=self.lang, use_corrector=self.use_corrector)
        self.logger = self.ocr_parser.logger

        self.replacement_map = {
            '\uf700': 'ำ',    # sara am
            '\uf701': '์',    # thanthakhat
            '\uf702': '้',    # mai tho
            '\uf706': '้',    # mai tho
            '\uf707': '๊',    # mai tri
            '\uf712': 'เ',    # sara e
            '\uf713': 'ใ',    # sara ai maimuan
            '\uf714': 'ไ',    # sara ai maimalai
            '\uf715': '์',    # thanthakhat
            '\uf716': '่',    # mai ek
            '\uf717': '๊',    # mai tri
            '\uf718': '๋',    # mai chattawa
            '\uf70a': '่',    # mai ek (声调符号)
            '\uf70e': '์',    # thanthakhat (静音符号)
            '\uf70b': '้',    # mai tho (声调符号)
            '\x0b': '(',
            '\x0c': ')',
            '\x9e': 'ป',
            '\x9f': 'ผ',
            '\x9a': 'ท',
            '\x8d': 'ช',
            'Ã': 'โ',
        }

    def _get_default_output_dir(self) -> str:
        """根据PDF文件名生成一个默认的输出目录。"""
        pdf_dir = os.path.dirname(os.path.abspath(self.pdf_path))
        pdf_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
        return os.path.join(pdf_dir, f"{pdf_name}_extracted_fields")

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
                item.DESCRIPTION_TH = self.replace_pua_thai(text.strip())
                remaining_text_list.remove(text)
                break

        # 将剩余部分作为描述
        item.DESCRIPTION = ' '.join(remaining_text_list).strip().replace('"', '')


    def fix_thai_thanthakhat(self, text):
        # 泰语辅音字母Unicode范围，包含所有辅音
        consonants = (
            'กขฃคฅฆงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรฤลฦวศษสหฬอฮ'
        )
    
        # 构造正则表达式，匹配辅音 + 空白(0个或多个) + 音杀符(์ U+0E4C)
        pattern = re.compile(
            f'([{consonants}])\\s*์'
        )
    
        # 替换为辅音+์，去掉空白
        fixed_text = pattern.sub(r'\1์', text)
        return fixed_text


    def replace_pua_thai(self, text):
        for pua_char, std_char in self.replacement_map.items():
            text = text.replace(pua_char, std_char)
            text = self.fix_thai_thanthakhat(text)
            
        return text

    def _parse_group_to_fields(self, group_data: dict) -> ImportFields:
        """将单个分组的OCR结果解析并填充到一个Fields对象中。"""
        item = ImportFields()
        rows = group_data.get('rows', [])

        def get_cell(r, c, default=''):
            try:
                return rows[r][c].strip()
            except (IndexError, AttributeError):
                return default
            
        def get_original_cell(r, c, default=''):
            try:
                return group_data.get('original_rows', [])[r][c].strip()
            except (IndexError, AttributeError):
                return default
            
        def get_digital_value(text):
            # 提取数字
            value_list = text.split('\n')
            for value in value_list:
                match = re.search(r'([0-9,.%]+)', value)
                if match:
                    return match.group(1)
            return text
            
        # --- 按单元格位置解析字段 ---
        # 第1行
        item.NO = get_cell(0, 0).split('\n')[0]
        item.HS_CODE = get_cell(0, 1)
        item.AMOUNT_USD = get_cell(0, 2).replace('USD\n', '').replace('USD ', '').strip()
        item.TAX_RATE = get_cell(0, 3)
        item.CUSTOMS_DUTIES_PAYABLE = get_digital_value(get_cell(0, 4))
        item.FEE = get_digital_value(get_cell(0, 5))
        item.EXCISE_PRODUCT_CODE = get_cell(0, 6)
        item.EXCISE_TAX = get_digital_value(get_cell(0, 7))
        item.VALUE_ADDED_TAX_BASE = get_cell(0, 8)

        # 第2行
        unit_codes = get_cell(1, 0).split('/')
        if len(unit_codes) == 2:
            item.UNIT_CODE_1 = unit_codes[0].strip()
            item.UNIT_CODE_2 = unit_codes[1].strip()
        item.AMOUNT_THB = get_cell(1, 1)
        item.DUTY_PAID = get_cell(1, 2)
        item.OTHER_TAXES = get_digital_value(get_cell(1, 3))
        item.EXCISE_TAX_RATE = get_digital_value(get_cell(1, 4))
        item.MINISTRY_OF_INTERIOR_TAX = get_digital_value(get_cell(1, 5))
        item.VAT = get_cell(1, 6)

        # 第3行
        item.PRIVILEGE_CODE = get_cell(2, 0)
        total_nw_text = get_cell(2, 1)
        nw_match = re.match(r'([0-9,.]+)\s*([A-Z]+)', total_nw_text, re.IGNORECASE)
        if nw_match:
            item.WEIGHT_UNIT = nw_match.group(2).replace('KGV', 'KGM')
            item.TOTAL_N_W = nw_match.group(1)
        else:
            item.WEIGHT_UNIT = total_nw_text

        qty_unit_text = get_cell(2, 2)
        qty_unit_match = re.match(r'([0-9,.]+)\s*([A-Z]+[0-9]*)', qty_unit_text, re.IGNORECASE)
        if qty_unit_match:
            item.QTY = qty_unit_match.group(1)
            item.QTY_UNIT = qty_unit_match.group(2)
        else:
            item.QTY = qty_unit_text

        # 解析描述块,找到字符串最长的单元格
        max_length = 0
        max_length_row = None
        for row in group_data.get('original_rows', [])[2]:
            if row is not None and len(row) > max_length and re.search(r'[\u0E00-\u0E7F]', row):
                max_length = len(row)
                max_length_row = row

        description_block = max_length_row
        description_block = self.replace_pua_thai(description_block)
        self._parse_description_block(item, original_block_text=description_block, ocr_block_text=get_cell(2, 3))

        # 第4行
        inv_text = get_original_cell(3, 0)
        inv_text_list = inv_text.split('\n')
        for text in inv_text_list:
            if re.search(r'[\u0E00-\u0E7F]', text):
                text = re.sub(r'[\u0E00-\u0E7F]+', '', text).strip()
            if 'T8' in text:
                #提取T8及之后的字符
                inv_match = re.search(r'(T8\S*)', text)
                if inv_match:
                    item.INV = inv_match.group(0).strip()
                else:
                    item.INV = text.strip()

        return item

    def save_to_json(self, items: list, filename: str = "extracted_fields.json"):
        """将提取出的字段列表保存为JSON文件。"""
        os.makedirs(self.output_dir, exist_ok=True)
        filepath = os.path.join(self.output_dir, filename)
        
        # 将Fields对象列表转换为字典列表以便序列化
        items_as_dicts = [item.__dict__ for item in items]
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(items_as_dicts, f, ensure_ascii=False, indent=2)
        self.logger.info(f"已将提取的字段保存到: {filepath}")

    def save_to_excel(self, items: list, filename: str = "extracted_fields.xlsx"):
        """将提取出的字段列表保存为Excel文件。"""
        os.makedirs(self.output_dir, exist_ok=True)
        filepath = os.path.join(self.output_dir, filename)
        
        # 创建Excel工作簿
        workbook = Workbook()
        sheet = workbook.active
        
        # 添加表头
        sheet.append([
                'NO', 	'MODEL',	'DESCRIPTION(英文描述)',	'DESCRIPTION(泰文描述)',
                'HS CODE',	'QTY',	'Unit',	'Unit Code 1', 'Unit Code 2',	'Privilege Code',	'AMOUNT(USD)',	'AMOUNT(THB)',
                'TOTAL.N.W', 'Weight unit',	'Tax rate',	'Customs duties payable',	'Duty paid',	'Inv.',
                'Fee',	'Excise Product Code',	'Excise tax rate',	'Excise tax',	'Other taxes',
                'Taxes for the Ministry of Interior',	'Value Added Tax Base',	'VAT',
                'FE Certificate No./Date',	'TISI Certificate No./Date',	'Explanation',	'Country of Origin',	'Usage Rules'
            ])
        
        # 添加数据
        for item in items:
            sheet.append([
                item.NO, item.MODEL, item.DESCRIPTION, item.DESCRIPTION_TH,
                item.HS_CODE, item.QTY, item.QTY_UNIT, item.UNIT_CODE_1, 
                item.UNIT_CODE_2, item.PRIVILEGE_CODE, item.AMOUNT_USD, 
                item.AMOUNT_THB, item.TOTAL_N_W, item.WEIGHT_UNIT, item.TAX_RATE, 
                item.CUSTOMS_DUTIES_PAYABLE, item.DUTY_PAID, item.INV, item.FEE, 
                item.EXCISE_PRODUCT_CODE, item.EXCISE_TAX_RATE, item.EXCISE_TAX, 
                item.OTHER_TAXES, item.MINISTRY_OF_INTERIOR_TAX,
                item.VALUE_ADDED_TAX_BASE, item.VAT, item.FE_CERTIFICATE_NO_DATE, 
                item.TISI_CERTIFICATE_NO_DATE, item.EXPLANATION, item.COUNTRY_OF_ORIGIN, 
                item.USAGE_RULES
            ])

        # 保存Excel文件
        workbook.save(filepath)
        self.logger.info(f"已将提取的字段保存到: {filepath}")


    def extract_items(self):
        """执行完整的提取流程：OCR -> 解析 -> 保存。"""
        self.logger.info("开始执行字段提取流程...")
        # 1. 使用OcrParser提取原始文本
        all_pages_groups = self.ocr_parser.extract_group_text(
            self.pdf_path,
            output_dir=self.output_dir,
            lang=self.lang,
            save_json=self.save_json,
        )

        if not all_pages_groups:
            self.logger.warning("OCR未能从PDF中提取任何分组，提取流程终止。")
            return None

        # 2. 遍历所有分组并解析字段
        extracted_items = []
        sorted_pages = sorted(all_pages_groups.keys(), key=int)
        for page_num_str in sorted_pages:
            for group_data in all_pages_groups[page_num_str]:
                item_fields = self._parse_group_to_fields(group_data)
                extracted_items.append(item_fields)
        
        self.logger.info(f"成功从 {len(all_pages_groups)} 个页面中解析出 {len(extracted_items)} 个项目。")

        # 3. 保存结果到JSON文件
        if self.save_json and extracted_items:
            filename = f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_extracted_fields.json"
            self.save_to_json(extracted_items, filename=filename)

        # 4. 保存结果到Excel文件
        if extracted_items:
            filename = f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_extracted_fields.xlsx"
            self.save_to_excel(extracted_items, filename=filename)

        return extracted_items

class ExportFields:
    """一个数据类，用于存放从报关单单个项目中提取的字段。"""
    def __init__(self):
        self.NO = ''    # NO (项号)
        self.MODEL = ''    # MODEL (型号)
        self.DESCRIPTION = ''    # DESCRIPTION(英文描述)
        self.DESCRIPTION_TH = ''    # DESCRIPTION(泰文描述)
        self.HS_CODE = ''    # HS CODE (海关编码)
        self.QTY = ''    # QTY (数量)
        self.QTY_UNIT = ''    # QTY Unit (数量单位)
        self.UNIT_CODE_1 = ''    # Unit Code 1 (单位代码1)
        self.UNIT_CODE_2 = ''    # Unit Code 2 (单位代码2)
        self.PACKAGE_QTY = ''    # Package Qty (包装数量)
        self.PACKAGE_TYPE = ''    # Package Type (包装类型)
        self.AMOUNT_USD = ''   # AMOUNT(USD) (美元金额)
        self.AMOUNT_THB = ''   # AMOUNT(THB) (泰铢金额)
        self.TOTAL_N_W = ''   # TOTAL.N.W (总净重)
        self.WEIGHT_UNIT = ''   # Weight unit (重量单位)
        self.TAX_RATE = ''   # Tax rate (出口税率)
        self.EXPORT_TAX = ''   # Export tax (出口税)
        self.CUSTOMS_DUTIES_PAYABLE = ''   # Customs duties payable (关税评估价格)
        self.PRIVILEGE_CODE = ''   # Privilege Code (优惠代码)
        self.INV = ''   # Inv. (发票号)
        self.VAT = ''   # VAT (增值税)
        self.FE_CERTIFICATE_NO_DATE = ''   # FE Certificate No./Date (FE证书号/日期)
        self.TISI_CERTIFICATE_NO_DATE = ''   # TISI Certificate No./Date (TISI证书号/日期)
        self.COUNTRY_OF_ORIGIN = ''   # Country of Origin (原产国)
        self.COUNTRY_OF_DESTINATION = ''   # Country of Destination (目的国)

class ExportFieldsExtractor(ImportFieldsExtractor):
    def __init__(self, pdf_path: str, output_dir: str = None, lang: str = 'en', save_json: bool = True, save_excel: bool = True, use_corrector: bool = False):
        super().__init__(pdf_path, output_dir, lang, save_json, save_excel, use_corrector)

    def get_digital_value(self, text):
            # 提取数字
            value_list = text.split('\n')
            for value in value_list:
                match = re.match(r'([0-9,.%]+)', value)
                if match:
                    return match.group(0)
            return text

    def _parse_group_to_fields(self, group_data: dict) ->  ExportFields:
        item = ExportFields()
        rows = group_data.get('rows', [])

        def get_cell(r, c, default=''):
            try:
                return rows[r][c].strip()
            except (IndexError, AttributeError):
                return default
            
        def get_original_cell(r, c, default=''):
            try:
                return group_data.get('original_rows', [])[r][c].strip()
            except (IndexError, AttributeError):
                return default
            

            
        # --- 按单元格位置解析字段 ---
        # HS编码和单位cell
        cell_values = get_cell(6, 0).split('\n')
        item.HS_CODE = self.get_digital_value(get_cell(6, 0))
        if item.HS_CODE:
            cell_values.remove(item.HS_CODE)
        if len(cell_values) > 0:
            unit_codes = cell_values[0].strip().split('/')
            if len(unit_codes) == 2:
                item.UNIT_CODE_1 = unit_codes[0].strip()
                item.UNIT_CODE_2 = unit_codes[1].strip()

        # 包装数量和包装类型cell
        cell_values = get_original_cell(0, 2).split('\n')
        for text in cell_values:
            if re.search(r'[\u0E00-\u0E7F]', text):
                cell_values.remove(text)
        item.PACKAGE_QTY = self.get_digital_value(get_original_cell(0, 2))
        if item.PACKAGE_QTY:
            cell_values.remove(item.PACKAGE_QTY)
        if len(cell_values) > 0:
            item.PACKAGE_TYPE = cell_values[0].strip()

        # 重量cell
        total_nw_text = get_cell(0, 3)
        nw_match = re.match(r'([0-9,.]+)\s*([A-Z]+)', total_nw_text, re.IGNORECASE)
        if nw_match:
            item.WEIGHT_UNIT = nw_match.group(2).replace('KGV', 'KGM')
            item.TOTAL_N_W = nw_match.group(1)
        else:
            item.WEIGHT_UNIT = total_nw_text

        # QTY 和数量单位cell
        qty_unit_text = get_cell(1, 0)
        qty_unit_match = re.match(r'([0-9,.]+)\s*([A-Z]+[0-9]*)', qty_unit_text, re.IGNORECASE)
        if qty_unit_match:
            item.QTY = qty_unit_match.group(1)
            item.QTY_UNIT = qty_unit_match.group(2)
        else:
            item.QTY = qty_unit_text

        # 美元金额cell
        item.AMOUNT_USD = get_cell(0, 4).replace('USD\n', '').replace('USD ', '').strip()

        # 泰铢金额cell
        item.AMOUNT_THB = get_cell(2, 0).replace('THB\n', '').replace('THB ', '').strip()

        # 出口税率cell
        item.TAX_RATE = get_cell(4, 0)

        # 出口税cell
        item.EXPORT_TAX = get_cell(6, 1)

        # 关税评估价格cell
        item.CUSTOMS_DUTIES_PAYABLE = get_cell(5, 0)

        # 优惠代码cell
        item.PRIVILEGE_CODE = get_cell(0, 5)

        item.NO = get_cell(0, 0).split('\n')[0]
        

        # 解析描述块,找到字符串最长的单元格
        max_length = 0
        max_length_row = None
        for row in group_data.get('original_rows', [])[3]:
            if row is not None and len(row) > max_length and re.search(r'[\u0E00-\u0E7F]', row):
                max_length = len(row)
                max_length_row = row

        description_block = max_length_row
        description_block = self.replace_pua_thai(description_block)
        self._parse_description_block(item, original_block_text=description_block, ocr_block_text=get_cell(3, 0))

        return item
    
    def _parse_description_block(self, item: ExportFields, original_block_text: str, ocr_block_text: str):
        if not original_block_text:
            return

        remaining_block = original_block_text

        
        # 提取原产国 (例如 CN, TH, US)
        country_match = re.search(r'\bOrigin : (CN|TH|US|JP|DE|VN)\b', remaining_block)
        if country_match:
            item.COUNTRY_OF_ORIGIN = country_match.group(1).strip()
            # 避免将 'CN' 作为描述的一部分
            remaining_block = remaining_block.replace(country_match.group(0), '', 1)

        # 提取目的地国 (例如 Origin Criteria PE)
        destination_country_match = re.search(r'\bPur.Country : (CN|TH|US|JP|DE|VN)\b', remaining_block, re.IGNORECASE)
        if destination_country_match:
            item.COUNTRY_OF_DESTINATION = destination_country_match.group(1).strip()
            remaining_block = remaining_block.replace(destination_country_match.group(0), '', 1)

        # 提取型号
        remaining_text_list = list(filter(lambda x: "BRAND" not in x.upper() and len(x.strip()) > 0 and "ชนิดของ" not in x, remaining_block.split('\n')))
        matched_texts = [text for text in remaining_text_list if re.match(r'^(?=.*[A-Z])[A-Z0-9-/]+$', text)]
        if matched_texts:
            if len(matched_texts) >= 2:
                chosen_text = matched_texts[1]  # 取第二个
            else:
                chosen_text = matched_texts[0]  # 取唯一一个
            
            item.MODEL = chosen_text.strip()
            remaining_text_list.remove(chosen_text)  # 从原列表移除
        
        # 泰文描述
        for text in remaining_text_list:
            # 是否含有泰文
            if re.search(r'[\u0E00-\u0E7F]', text):
                item.DESCRIPTION_TH = self.replace_pua_thai(text.strip())
                remaining_text_list.remove(text)
                break
        # 发票号
        for text in remaining_text_list:
            if re.search(r'[\u0E00-\u0E7F]', text):
                text = re.sub(r'[\u0E00-\u0E7F]+', '', text).strip()
            if 'BOI' in text:
                #提取-BOI及之前的字符
                inv_match = re.search(r'(\S*BOI)', text)
                if inv_match:
                    item.INV = inv_match.group(0).strip()
                    remaining_text_list.remove(text)
                else:
                    item.INV = text.strip()
                    remaining_text_list.remove(text)

        # 将剩余部分作为描述
        for text in remaining_text_list:
            # 是否含有泰文
            if re.search(r'[\u0E00-\u0E7F]', text):
                remaining_text_list.remove(text)
        item.DESCRIPTION = ' '.join(remaining_text_list).strip().replace('"', '')
    
    def extract_items(self):
        """执行完整的提取流程：OCR -> 解析 -> 保存。"""
        self.logger.info("开始执行字段提取流程...")
        # 1. 使用OcrParser提取原始文本
        all_pages_groups = self.ocr_parser.extract_group_text(
            self.pdf_path,
            output_dir=self.output_dir,
            lang=self.lang,
            save_json=self.save_json,
            group_size=8
        )

        if not all_pages_groups:
            self.logger.warning("OCR未能从PDF中提取任何分组，提取流程终止。")
            return None

        # 2. 遍历所有分组并解析字段
        extracted_items = []
        sorted_pages = sorted(all_pages_groups.keys(), key=int)
        for page_num_str in sorted_pages:
            for group_data in all_pages_groups[page_num_str]:
                item_fields = self._parse_group_to_fields(group_data)
                extracted_items.append(item_fields)
        
        self.logger.info(f"成功从 {len(all_pages_groups)} 个页面中解析出 {len(extracted_items)} 个项目。")

        # 3. 保存结果到JSON文件
        if self.save_json and extracted_items:
            filename = f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_extracted_fields.json"
            self.save_to_json(extracted_items, filename=filename)

        # 4. 保存结果到Excel文件
        if extracted_items:
            filename = f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_extracted_fields.xlsx"
            self.save_to_excel(extracted_items, filename=filename)

        return extracted_items
    
    def save_to_excel(self, items: list, filename: str = "extracted_fields.xlsx"):
        """将提取出的字段列表保存为Excel文件。"""
        os.makedirs(self.output_dir, exist_ok=True)
        filepath = os.path.join(self.output_dir, filename)
        
        # 创建Excel工作簿
        workbook = Workbook()
        sheet = workbook.active
        
        # 添加表头
        sheet.append([
                'NO', 	'MODEL',	'DESCRIPTION(英文描述)',	'DESCRIPTION(泰文描述)',
                'HS CODE',	'QTY',	'Unit',	'Unit Code 1', 'Unit Code 2',	'Privilege Code',	
                'Package Qty', 'Package Type',	'AMOUNT(USD)',	'AMOUNT(THB)',
                'TOTAL.N.W', 'Weight unit',	'Tax rate',	'Export tax',	'Customs duties payable',	'Inv.',
                'Country of Origin',	'Country of Destination'
            ])
        
        # 添加数据
        for item in items:
            sheet.append([
                item.NO, item.MODEL, item.DESCRIPTION, item.DESCRIPTION_TH,
                item.HS_CODE, item.QTY, item.QTY_UNIT, item.UNIT_CODE_1, 
                item.UNIT_CODE_2, item.PRIVILEGE_CODE, item.PACKAGE_QTY, 
                item.PACKAGE_TYPE, item.AMOUNT_USD, item.AMOUNT_THB, 
                item.TOTAL_N_W, item.WEIGHT_UNIT, item.TAX_RATE, item.EXPORT_TAX, 
                item.CUSTOMS_DUTIES_PAYABLE, item.INV, item.COUNTRY_OF_ORIGIN, 
                item.COUNTRY_OF_DESTINATION
            ])

        # 保存Excel文件
        workbook.save(filepath)
        self.logger.info(f"已将提取的字段保存到: {filepath}")

if __name__ == '__main__':
    # 为 FieldsExtractor 添加命令行测试入口
    import argparse
    
    # 设置父解析器以共享参数
    parent_parser = argparse.ArgumentParser(add_help=False)
    parent_parser.add_argument("pdf_path", help="PDF文件路径。")
    parent_parser.add_argument("-o", "--output", help="输出目录路径。默认为PDF旁边的新建文件夹。")
    parent_parser.add_argument("--lang", default="en", help="OCR识别语言 (例如 'en', 'ch', 'th')。默认: 'en'。")
    parent_parser.add_argument("--type", default="import", help="提取类型 (例如 'import', 'export')。默认: 'import'。")

    parser = argparse.ArgumentParser(
        description="从PDF报关单中提取结构化字段。",
        parents=[parent_parser]
    )
    
    args = parser.parse_args()

    # 初始化并运行提取器
    if args.type == 'import':
        extractor = ImportFieldsExtractor(
            pdf_path=args.pdf_path,
            output_dir=args.output,
            lang=args.lang,
            save_json=True,
            save_excel=True
        )
    elif args.type == 'export':
        extractor = ExportFieldsExtractor(
            pdf_path=args.pdf_path,
            output_dir=args.output,
            lang=args.lang,
            save_json=True,
            save_excel=True
    )
    extractor.extract_items()
