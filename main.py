from converter import Converter


def convert(pdf_file: str, docx_file: str = None, start: int = 0, end: int = None, pages: list = None, **kwargs):
    if not kwargs.get('zero_based_index', True):
        start = max(start - 1, 0)
        if end: end -= 1
        if pages: pages = [i - 1 for i in pages]

    cv = Converter(pdf_file)
    cv.convert(docx_file, start, end, pages, kwargs)
    cv.close()


pdf_file = "C:\\Users\\Ripsita\\Desktop\\workfiles\\ast_sci_data_tables_sample.pdf"
docx_file: str = 'C:\\Users\\Ripsita\\Desktop\\workfiles\\converted_1.docx'

convert(pdf_file, docx_file);
