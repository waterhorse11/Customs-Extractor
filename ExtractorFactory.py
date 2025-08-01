from ExtractorImportTemplate import TianShiExtractor as TianShiImportExtractor, LssExtractor as LssImportExtractor, HlsExtractor as HlsImportExtractor, SnpExtractor as SnpImportExtractor, OlcExtractor as OLCImportExtractor
from ExtractorExportTemplate import TianShiExtractor as TianShiExportExtractor, HlsExtractor as HlsExportExtractor

class ExtractorFactory:
    """
    一个工厂类，用于根据指定的模板类型创建对应的字段提取器实例。
    """
    @staticmethod
    def create_extractor(template_type: str, pdf_path: str, output_dir: str = None, lang: str = 'en', type: str = 'import'):
        """
        根据模板类型创建并返回一个具体的FieldsExtractor实例。

        Args:
            template_type (str): 模板的名称 (例如, 'TianShi', 'LSS', 'HLS', 'SNP').
            pdf_path (str): PDF文件的路径。
            output_dir (str, optional): 输出目录的路径。
            lang (str, optional): OCR语言。
            type (str, optional): 模板类型 (例如, 'import', 'export').

        Returns:
            一个FieldsExtractor的子类实例，如果模板类型未知则返回None。
        """
        if type == 'import':
            if template_type == 'TianShi':
                return TianShiImportExtractor(pdf_path, output_dir, lang)
            elif template_type == 'LSS':
                return LssImportExtractor(pdf_path, output_dir, lang)
            elif template_type == 'HLS':
                return HlsImportExtractor(pdf_path, output_dir, lang)
            elif template_type == 'OLC':
                return OLCImportExtractor(pdf_path, output_dir, lang, use_corrector=True)
            elif template_type == 'SNP':
                return SnpImportExtractor(pdf_path, output_dir, lang)
            else:
                raise ValueError(f"未知的模板类型: {template_type}") 
        elif type == 'export':
            if template_type == 'TianShi':
                return TianShiExportExtractor(pdf_path, output_dir, lang)
            elif template_type == 'HLS':
                return HlsExportExtractor(pdf_path, output_dir, lang)
            else:
                raise ValueError(f"未知的模板类型: {template_type}") 
