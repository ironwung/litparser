"""
문서 포맷 파서들

Modern (OOXML/ZIP):
- docx_parser: docx
- pptx_parser: pptx
- xlsx_parser: xlsx
- hwpx_parser: hwpx

Legacy (OLE2/Binary):
- doc_parser: doc
- ppt_parser: ppt
- xls_parser: xls
- hwp_parser: hwp

Text:
- text_parser: txt, md
"""
from .text_parser import parse_text, parse_markdown, TextDocument
from .docx_parser import parse_docx, DocxDocument
from .pptx_parser import parse_pptx, PptxDocument
from .xlsx_parser import parse_xlsx, XlsxDocument
from .hwpx_parser import parse_hwpx, HwpxDocument

# Legacy parsers
from .doc_parser import parse_doc, DocDocument
from .ppt_parser import parse_ppt, PptDocument
from .xls_parser import parse_xls, XlsDocument
from .hwp_parser import parse_hwp, HwpDocument

__all__ = [
    # Text
    'parse_text', 'parse_markdown', 'TextDocument',
    # Modern
    'parse_docx', 'DocxDocument',
    'parse_pptx', 'PptxDocument',
    'parse_xlsx', 'XlsxDocument',
    'parse_hwpx', 'HwpxDocument',
    # Legacy
    'parse_doc', 'DocDocument',
    'parse_ppt', 'PptDocument',
    'parse_xls', 'XlsDocument',
    'parse_hwp', 'HwpDocument',
]
