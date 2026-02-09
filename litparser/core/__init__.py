"""
PDF Parser Core Module
"""
from .parser import PDFParser, PDFDocument, PDFRef, XRefEntry
from .stream_decoder import StreamDecoder
from .content_stream import (
    ContentStreamParser, ContentStreamLexer,
    TextItem, TextState, FontInfo,
    parse_tounicode_cmap
)
from .image_extractor import PDFImage, extract_images, save_image, raw_to_png
from .table_detector import Table, TableCell, detect_tables, extract_tables_from_page
from .layout_analyzer import (
    PageLayout, TextBlock, BlockType,
    analyze_layout, analyze_page_layout
)
from .struct_tree import (
    StructTreeParser, StructElement, StructTable, StructType,
    extract_tables_from_struct_tree, is_tagged_pdf
)
from .ole_parser import OLE2Reader, is_ole2_file

__all__ = [
    # Parser
    'PDFParser', 'PDFDocument', 'PDFRef', 'XRefEntry', 'StreamDecoder',
    # Content Stream
    'ContentStreamParser', 'ContentStreamLexer', 
    'TextItem', 'TextState', 'FontInfo',
    'parse_tounicode_cmap',
    # Image Extractor
    'PDFImage', 'extract_images', 'save_image', 'raw_to_png',
    # Table Detector
    'Table', 'TableCell', 'detect_tables', 'extract_tables_from_page',
    # Layout Analyzer
    'PageLayout', 'TextBlock', 'BlockType', 'analyze_layout', 'analyze_page_layout',
    # Struct Tree
    'StructTreeParser', 'StructElement', 'StructTable', 'StructType',
    'extract_tables_from_struct_tree', 'is_tagged_pdf',
    # OLE2
    'OLE2Reader', 'is_ole2_file',
]
