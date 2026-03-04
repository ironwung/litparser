"""
LitParser - Lightweight Document Parser

다양한 문서 포맷을 파싱하는 라이브러리
- PDF, DOCX, PPTX, HWPX, TXT, MD 지원
- 외부 라이브러리 없이 순수 Python으로 구현

사용법:
    from litparser import parse, to_markdown, to_json
    
    # 자동 포맷 감지
    result = parse('document.pdf')
    result = parse('report.docx')
    
    # 마크다운/JSON 변환
    md = to_markdown(result)
    json_str = to_json(result)
    
    # 개별 옵션
    result = parse('document.pdf', include_images=True)
    
    # PDF 전용 API (하위 호환)
    from litparser import parse_pdf, extract_text
    doc = parse_pdf('document.pdf')
    text = extract_text(doc, page_num=0)
"""
from pathlib import Path
from typing import Union, Optional
from dataclasses import dataclass, field
from .core import (
    PDFParser, PDFDocument, PDFRef, XRefEntry, StreamDecoder,
    ContentStreamParser, FontInfo, TextItem, parse_tounicode_cmap
)
from .core.image_extractor import PDFImage, extract_images, save_image, raw_to_png
from .core.table_detector import Table, TableCell, detect_tables, extract_tables_from_page
from .core.layout_analyzer import (
    PageLayout, TextBlock, BlockType, 
    analyze_layout, analyze_page_layout
)
from .core.struct_tree import (
    StructTreeParser, StructElement, StructTable, StructType,
    extract_tables_from_struct_tree, is_tagged_pdf
)
import re

__version__ = '0.7.0'
__all__ = [
    # 통합 API
    'parse', 'to_markdown', 'to_json', 'ParseResult',
    # Core
    'parse_pdf', 'extract_text', 'extract_text_with_positions', 
    'extract_all_text', 'get_page_count', 'get_pages', 'decode_stream',
    'PDFParser', 'PDFDocument', 'PDFRef', 'StreamDecoder', 'TextItem',
    # Stage 5: Images
    'extract_images', 'save_image', 'PDFImage', 'raw_to_png',
    # Stage 5: Tables
    'extract_tables', 'detect_tables', 'Table', 'TableCell',
    # Stage 5: Layout
    'analyze_layout', 'analyze_page_layout', 'PageLayout', 'TextBlock', 'BlockType',
    # Stage 5: StructTree
    'is_tagged_pdf', 'get_document_outline', 'get_document_structure',
    'StructTreeParser', 'StructElement', 'StructTable',
]


# =============================================================================
# 통합 ParseResult 클래스
# =============================================================================

@dataclass
class ParseResult:
    """통합 파싱 결과"""
    # 메타데이터
    filename: str = ""
    format: str = ""  # pdf, docx, pptx, hwpx, txt, md
    page_count: int = 0
    title: str = ""
    author: str = ""
    
    # 콘텐츠
    text: str = ""
    pages: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    images: list = field(default_factory=list)
    headings: list = field(default_factory=list)
    
    # 원본 문서 객체 (고급 사용)
    _doc: object = field(default=None, repr=False)


# =============================================================================
# 통합 parse() 함수
# =============================================================================

def parse(
    filepath_or_bytes: Union[str, bytes],
    filename: str = None,
    include_images: bool = False,
) -> ParseResult:
    """
    문서 파싱 (자동 포맷 감지)
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트 데이터
        filename: 바이트 입력 시 파일명 (포맷 감지용)
        include_images: 이미지 데이터 포함 여부
    
    Returns:
        ParseResult: 파싱 결과
    
    Examples:
        result = parse('document.pdf')
        result = parse('report.docx')
        result = parse(pdf_bytes, filename='doc.pdf')
        
        print(result.text)
        for table in result.tables:
            print(table['markdown'])
    
    Supported formats:
        .pdf, .docx, .pptx, .hwpx, .txt, .md
    """
    import base64
    
    # 확장자 감지
    if isinstance(filepath_or_bytes, str):
        ext = Path(filepath_or_bytes).suffix.lower()
        fname = filepath_or_bytes
    else:
        ext = Path(filename).suffix.lower() if filename else _detect_format(filepath_or_bytes)
        fname = filename or ""
    
    result = ParseResult(filename=fname, format=ext.lstrip('.'))
    
    # PDF
    if ext == '.pdf':
        doc = parse_pdf(filepath_or_bytes)
        result._doc = doc
        result.page_count = get_page_count(doc)
        
        all_text = []
        for page_num in range(result.page_count):
            page_text = extract_text(doc, page_num)
            all_text.append(page_text)
            
            page_data = {
                'page': page_num + 1,
                'text': page_text,
                'tables': [],
                'images': [],
            }
            
            # 테이블
            tables = extract_tables(doc, page_num)
            for t in tables:
                table_data = {
                    'rows': t.rows,
                    'cols': t.cols,
                    'data': t.to_list(),
                    'markdown': t.to_markdown(),
                }
                page_data['tables'].append(table_data)
                result.tables.append(table_data)
            
            result.pages.append(page_data)
        
        result.text = "\n\n".join(all_text)
        
        # 이미지
        if include_images:
            images = extract_images(doc)
            for img in images:
                img_data = {
                    'width': img.width,
                    'height': img.height,
                    'format': img.color_space,
                }
                if img.data:
                    img_data['base64'] = base64.b64encode(img.data).decode('ascii')
                result.images.append(img_data)
        
        # 개요
        try:
            outline = get_document_outline(doc)
            result.headings = [{'level': lv, 'text': txt} for lv, txt in outline]
        except:
            pass
    
    # DOCX
    elif ext == '.docx':
        from .formats.docx_parser import parse_docx
        doc = parse_docx(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = 1
        result.text = doc.get_text()
        result.headings = [{'level': lv, 'text': txt} for lv, txt in doc.get_headings()]
        
        for t in doc.tables:
            result.tables.append({
                'rows': len(t.rows),
                'cols': len(t.rows[0]) if t.rows else 0,
                'data': t.rows,
                'markdown': t.to_markdown(),
            })
        
        if include_images:
            for img in doc.images:
                img_data = {
                    'filename': img.filename,
                    'format': img.content_type,
                }
                if img.data:
                    img_data['base64'] = base64.b64encode(img.data).decode('ascii')
                result.images.append(img_data)
    
    # PPTX
    elif ext == '.pptx':
        from .formats.pptx_parser import parse_pptx
        doc = parse_pptx(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = doc.slide_count
        
        all_text = []
        for slide in doc.slides:
            page_data = {
                'page': slide.number,
                'title': slide.title,
                'text': slide.get_text(),
                'tables': [],
            }
            all_text.append(slide.get_text())
            
            for t in slide.tables:
                table_data = {
                    'rows': len(t.rows),
                    'cols': len(t.rows[0]) if t.rows else 0,
                    'data': t.rows,
                    'markdown': t.to_markdown(),
                }
                page_data['tables'].append(table_data)
                result.tables.append(table_data)
            
            result.pages.append(page_data)
            
            if slide.title:
                result.headings.append({'level': 1, 'text': slide.title})
        
        result.text = "\n\n".join(all_text)
        
        if include_images:
            for img in doc.images:
                img_data = {
                    'filename': img.filename,
                    'format': img.content_type,
                }
                if img.data:
                    img_data['base64'] = base64.b64encode(img.data).decode('ascii')
                result.images.append(img_data)
    
    # HWPX
    elif ext == '.hwpx':
        from .formats.hwpx_parser import parse_hwpx
        doc = parse_hwpx(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = 1
        result.text = doc.get_text()
        result.headings = [{'level': lv, 'text': txt} for lv, txt in doc.get_headings()]
        
        for t in doc.tables:
            result.tables.append({
                'rows': len(t.rows),
                'cols': len(t.rows[0]) if t.rows else 0,
                'data': t.rows,
                'markdown': t.to_markdown(),
            })
        
        if include_images:
            for img in doc.images:
                img_data = {
                    'filename': img.filename,
                    'format': img.content_type,
                }
                if img.data:
                    img_data['base64'] = base64.b64encode(img.data).decode('ascii')
                result.images.append(img_data)
    
    # HWP (바이너리)
    elif ext == '.hwp':
        from .formats.hwp_parser import parse_hwp
        doc = parse_hwp(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = 1
        result.text = doc.get_text()
        result.headings = [{'level': lv, 'text': txt} for lv, txt in doc.get_headings()]
        
        for t in doc.tables:
            result.tables.append({
                'rows': len(t.rows),
                'cols': len(t.rows[0]) if t.rows else 0,
                'data': t.rows,
                'markdown': t.to_markdown(),
            })
    
    # XLSX
    elif ext == '.xlsx':
        from .formats.xlsx_parser import parse_xlsx
        doc = parse_xlsx(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = doc.sheet_count
        result.text = doc.get_text()
        
        # 각 시트를 페이지로
        for sheet in doc.sheets:
            page_data = {
                'page': sheet.index + 1,
                'name': sheet.name,
                'text': sheet.get_text(),
                'tables': [],
            }
            
            # 시트 자체를 테이블로
            if sheet.cells:
                table_data = {
                    'name': sheet.name,
                    'rows': sheet.rows,
                    'cols': sheet.cols,
                    'data': sheet.to_list(),
                    'markdown': sheet.to_markdown(),
                }
                page_data['tables'].append(table_data)
                result.tables.append(table_data)
            
            result.pages.append(page_data)
            result.headings.append({'level': 1, 'text': sheet.name})
    
    # TXT / MD
    elif ext in ['.txt', '.md', '.markdown']:
        from .formats.text_parser import parse_text, parse_markdown
        
        is_md = ext in ['.md', '.markdown']
        doc = parse_markdown(filepath_or_bytes) if is_md else parse_text(filepath_or_bytes)
        result._doc = doc
        result.format = 'md' if is_md else 'txt'
        result.page_count = 1
        result.text = doc.content
        
        if is_md and doc.headings:
            result.headings = [{'level': lv, 'text': txt} for lv, txt in doc.headings]
    
    # DOC (Word 97-2003)
    elif ext == '.doc':
        from .formats.doc_parser import parse_doc
        doc = parse_doc(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = 1
        result.text = doc.get_text()
        result.headings = [{'level': lv, 'text': txt} for lv, txt in doc.get_headings()]
        
        for t in doc.tables:
            result.tables.append({
                'rows': len(t.rows),
                'cols': len(t.rows[0]) if t.rows else 0,
                'data': t.rows,
                'markdown': t.to_markdown(),
            })
    
    # PPT (PowerPoint 97-2003)
    elif ext == '.ppt':
        from .formats.ppt_parser import parse_ppt
        doc = parse_ppt(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = doc.slide_count
        result.text = doc.get_text()
        
        for slide in doc.slides:
            page_data = {
                'page': slide.number,
                'title': slide.title,
                'text': slide.get_text(),
                'tables': [],
            }
            result.pages.append(page_data)
            
            if slide.title:
                result.headings.append({'level': 1, 'text': slide.title})
        
        if include_images:
            for img in doc.images:
                img_data = {
                    'filename': img.filename,
                    'format': img.content_type,
                }
                if img.data:
                    img_data['base64'] = base64.b64encode(img.data).decode('ascii')
                result.images.append(img_data)
    
    # XLS (Excel 97-2003)
    elif ext == '.xls':
        from .formats.xls_parser import parse_xls
        doc = parse_xls(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = doc.sheet_count
        result.text = doc.get_text()
        
        for sheet in doc.sheets:
            page_data = {
                'page': sheet.index + 1,
                'name': sheet.name,
                'text': sheet.get_text(),
                'tables': [],
            }
            
            if sheet.cells:
                table_data = {
                    'name': sheet.name,
                    'rows': sheet.rows,
                    'cols': sheet.cols,
                    'data': sheet.to_list(),
                    'markdown': sheet.to_markdown(),
                }
                page_data['tables'].append(table_data)
                result.tables.append(table_data)
            
            result.pages.append(page_data)
            result.headings.append({'level': 1, 'text': sheet.name})
    
    # HWP (한글 5.0 바이너리)
    elif ext == '.hwp':
        from .formats.hwp_parser import parse_hwp
        doc = parse_hwp(filepath_or_bytes)
        result._doc = doc
        result.title = doc.title
        result.author = doc.author
        result.page_count = 1
        result.text = doc.get_text()
        result.headings = [{'level': lv, 'text': txt} for lv, txt in doc.get_headings()]
        
        for t in doc.tables:
            result.tables.append({
                'rows': len(t.rows),
                'cols': len(t.rows[0]) if t.rows else 0,
                'data': t.rows,
                'markdown': t.to_markdown(),
            })
    
    else:
        raise ValueError(f"지원하지 않는 포맷: {ext}")
    
    return result


def _detect_format(data: bytes) -> str:
    """바이트 데이터에서 포맷 감지"""
    # PDF
    if data[:5] == b'%PDF-':
        return '.pdf'
    
    # ZIP (docx, pptx, hwpx, xlsx)
    if data[:4] == b'PK\x03\x04':
        import zipfile
        from io import BytesIO
        try:
            zf = zipfile.ZipFile(BytesIO(data), 'r')
            names = zf.namelist()
            zf.close()
            
            if any('word/' in n for n in names):
                return '.docx'
            if any('ppt/' in n for n in names):
                return '.pptx'
            if any('xl/' in n for n in names):
                return '.xlsx'
            if any('Contents/' in n for n in names):
                return '.hwpx'
        except:
            pass
    
    # OLE2 (doc, ppt, xls, hwp)
    if data[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        from .core.ole_parser import OLE2Reader
        try:
            ole = OLE2Reader(data)
            streams = ole.list_all()
            
            if 'WordDocument' in streams:
                return '.doc'
            if 'PowerPoint Document' in streams:
                return '.ppt'
            if 'Workbook' in streams or 'Book' in streams:
                return '.xls'
            if 'FileHeader' in streams:
                # HWP 확인
                header = ole.get_stream('FileHeader')
                if header and header[:17].decode('utf-8', errors='ignore').startswith('HWP'):
                    return '.hwp'
        except:
            pass
    
    return '.txt'


# =============================================================================
# 출력 변환 함수
# =============================================================================

def to_markdown(result: ParseResult, include_images: bool = False) -> str:
    """
    ParseResult를 마크다운으로 변환
    
    Args:
        result: parse() 결과
        include_images: 이미지를 base64로 포함
    
    Returns:
        str: 마크다운 문자열
    """
    lines = []
    
    # 제목
    title = result.title or Path(result.filename).stem if result.filename else "Document"
    lines.append(f"# {title}")
    lines.append("")
    
    # 메타데이터
    if result.author or result.page_count > 1:
        lines.append("> **문서 정보**")
        if result.author:
            lines.append(f"> - 작성자: {result.author}")
        if result.format:
            lines.append(f"> - 포맷: {result.format.upper()}")
        if result.page_count > 1:
            lines.append(f"> - 페이지: {result.page_count}")
        lines.append("")
    
    # 목차
    if result.headings:
        lines.append("## 목차")
        lines.append("")
        for h in result.headings:
            indent = "  " * (h['level'] - 1)
            lines.append(f"{indent}- {h['text']}")
        lines.append("")
        lines.append("---")
        lines.append("")
    
    # 페이지별 내용
    if result.pages:
        for page in result.pages:
            if result.page_count > 1:
                lines.append(f"## 페이지 {page['page']}")
                lines.append("")
            
            if page.get('text'):
                lines.append(page['text'])
                lines.append("")
            
            for i, table in enumerate(page.get('tables', []), 1):
                lines.append(f"### 테이블 {i}")
                lines.append("")
                lines.append(table['markdown'])
                lines.append("")
    elif result.text:
        lines.append(result.text)
        lines.append("")
        
        for i, table in enumerate(result.tables, 1):
            lines.append(f"### 테이블 {i}")
            lines.append("")
            lines.append(table['markdown'])
            lines.append("")
    
    # 이미지
    if include_images and result.images:
        lines.append("## 이미지")
        lines.append("")
        for i, img in enumerate(result.images, 1):
            if img.get('base64'):
                mime = img.get('mime', 'image/png')
                lines.append(f"![이미지 {i}](data:{mime};base64,{img['base64']})")
                lines.append("")
    
    return "\n".join(lines)


def to_json(result: ParseResult, include_images: bool = False, indent: int = 2) -> str:
    """
    ParseResult를 JSON으로 변환
    
    Args:
        result: parse() 결과
        include_images: 이미지 base64 포함
        indent: JSON 들여쓰기
    
    Returns:
        str: JSON 문자열
    """
    import json
    
    data = {
        'metadata': {
            'filename': result.filename,
            'format': result.format,
            'page_count': result.page_count,
            'title': result.title,
            'author': result.author,
        },
        'content': {
            'text': result.text,
            'headings': result.headings,
        },
        'pages': result.pages,
        'tables': result.tables,
        'images': result.images if include_images else [
            {k: v for k, v in img.items() if k != 'base64'}
            for img in result.images
        ],
    }
    
    return json.dumps(data, ensure_ascii=False, indent=indent)


def to_dict(result: ParseResult, include_images: bool = False) -> dict:
    """ParseResult를 딕셔너리로 변환"""
    import json
    return json.loads(to_json(result, include_images))


def parse_pdf(filepath_or_bytes):
    """
    PDF 파일 또는 바이트를 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 (str) 또는 PDF 데이터 (bytes)
    
    Returns:
        PDFDocument: 파싱된 PDF 문서 객체
    
    Example:
        doc = parse_pdf('document.pdf')
        doc = parse_pdf(pdf_bytes)
    """
    if isinstance(filepath_or_bytes, str):
        with open(filepath_or_bytes, 'rb') as f:
            data = f.read()
    else:
        data = filepath_or_bytes
    
    parser = PDFParser(data)
    return parser.parse()


def get_page_count(doc: PDFDocument) -> int:
    """페이지 수 반환"""
    root_ref = doc.trailer.get('Root')
    if not isinstance(root_ref, PDFRef):
        return 0
    
    catalog = doc.objects.get((root_ref.obj_num, root_ref.gen_num))
    if not catalog:
        return 0
    
    pages_ref = catalog.get('Pages')
    if not isinstance(pages_ref, PDFRef):
        return 0
    
    pages = doc.objects.get((pages_ref.obj_num, pages_ref.gen_num))
    return pages.get('Count', 0) if pages else 0


def decode_stream(doc: PDFDocument, stream_obj: dict) -> bytes:
    """
    스트림 객체 디코딩
    
    Args:
        doc: PDFDocument 객체
        stream_obj: 스트림을 포함한 객체 딕셔너리
    
    Returns:
        bytes: 디코딩된 스트림 데이터
    """
    if '_stream_data' not in stream_obj:
        return b''
    
    filters = stream_obj.get('Filter', [])
    if not filters:
        return stream_obj['_stream_data']
    
    if isinstance(filters, str):
        filters = [filters]
    
    return StreamDecoder.decode(stream_obj['_stream_data'], filters)


def get_pages(doc: PDFDocument) -> list:
    """모든 페이지 객체 반환 (중첩된 페이지 트리 지원)"""
    root_ref = doc.trailer.get('Root')
    if not isinstance(root_ref, PDFRef):
        return []
    
    catalog = doc.objects.get((root_ref.obj_num, root_ref.gen_num))
    if not catalog:
        return []
    
    pages_ref = catalog.get('Pages')
    if not isinstance(pages_ref, PDFRef):
        return []
    
    pages_obj = doc.objects.get((pages_ref.obj_num, pages_ref.gen_num))
    if not pages_obj:
        return []
    
    # 재귀적으로 페이지 수집
    def collect_pages(node_ref):
        """페이지 트리를 재귀적으로 탐색하여 모든 Page 객체 수집"""
        if not isinstance(node_ref, PDFRef):
            return []
        
        node = doc.objects.get((node_ref.obj_num, node_ref.gen_num))
        if not node:
            return []
        
        node_type = node.get('Type')
        
        if node_type == 'Page':
            # 실제 페이지
            return [node]
        elif node_type == 'Pages':
            # 페이지 컨테이너 - Kids를 재귀 탐색
            result = []
            kids = node.get('Kids', [])
            for kid_ref in kids:
                result.extend(collect_pages(kid_ref))
            return result
        else:
            return []
    
    return collect_pages(pages_ref)


def _get_page_dimensions(doc: PDFDocument, page_num: int) -> tuple:
    """
    페이지 크기(width, height) 반환
    
    Args:
        doc: PDFDocument 객체
        page_num: 페이지 번호
    
    Returns:
        (width, height) 튜플. 기본값 (595, 842) = A4
    """
    pages = get_pages(doc)
    if page_num >= len(pages):
        return 595.0, 842.0
    
    page = pages[page_num]
    media_box = page.get('MediaBox', [0, 0, 595, 842])
    try:
        w = float(media_box[2]) - float(media_box[0])
        h = float(media_box[3]) - float(media_box[1])
    except (TypeError, IndexError, ValueError):
        w, h = 595.0, 842.0
    
    return w, h


def _build_font_map(doc: PDFDocument, page: dict) -> dict:
    """페이지의 폰트 정보 수집 (ToUnicode CMap 포함)"""
    font_map = {}
    
    resources = page.get('Resources', {})
    if isinstance(resources, PDFRef):
        resources = doc.objects.get((resources.obj_num, resources.gen_num), {})
    
    fonts = resources.get('Font', {})
    if isinstance(fonts, PDFRef):
        fonts = doc.objects.get((fonts.obj_num, fonts.gen_num), {})
    
    for font_name, font_ref in fonts.items():
        if isinstance(font_ref, PDFRef):
            font_obj = doc.objects.get((font_ref.obj_num, font_ref.gen_num))
        else:
            font_obj = font_ref
        
        if not font_obj or not isinstance(font_obj, dict):
            continue
        
        font_info = FontInfo(
            name=font_name,
            subtype=font_obj.get('Subtype', ''),
            base_font=font_obj.get('BaseFont', ''),
            encoding=str(font_obj.get('Encoding', ''))
        )
        
        # ToUnicode CMap 파싱
        tounicode_ref = font_obj.get('ToUnicode')
        if isinstance(tounicode_ref, PDFRef):
            tounicode_obj = doc.objects.get((tounicode_ref.obj_num, tounicode_ref.gen_num))
            if tounicode_obj and '_stream_data' in tounicode_obj:
                try:
                    cmap_data = decode_stream(doc, tounicode_obj)
                    font_info.to_unicode = parse_tounicode_cmap(cmap_data)
                except Exception:
                    # CMap 디코딩 실패 시 해당 폰트만 스킵 (나머지 계속 진행)
                    pass
        
        font_map[font_name] = font_info
    
    return font_map


def extract_text_with_positions(doc: PDFDocument, page_num: int = 0) -> list:
    """
    페이지에서 위치 정보와 함께 텍스트 추출 (Top-Down 좌표계로 정규화)
    
    Args:
        doc: PDFDocument 객체
        page_num: 페이지 번호 (0부터 시작)
    
    Returns:
        List[TextItem]: 텍스트 항목 리스트 (text, x, y, font_name, font_size)
                        Y좌표는 항상 페이지 상단이 0, 하단이 큰 값
    """
    pages = get_pages(doc)
    if page_num >= len(pages):
        return []
    
    page = pages[page_num]
    
    # 폰트 맵 구성 (ToUnicode 포함)
    font_map = _build_font_map(doc, page)
    
    # Content Stream 가져오기
    contents_ref = page.get('Contents')
    content_stream = b''
    
    if isinstance(contents_ref, PDFRef):
        contents_obj = doc.objects.get((contents_ref.obj_num, contents_ref.gen_num))
        if contents_obj:
            try:
                content_stream = decode_stream(doc, contents_obj)
            except Exception:
                content_stream = b''
    elif isinstance(contents_ref, list):
        # 여러 Content Stream 합치기
        for ref in contents_ref:
            if isinstance(ref, PDFRef):
                obj = doc.objects.get((ref.obj_num, ref.gen_num))
                if obj:
                    try:
                        content_stream += decode_stream(doc, obj) + b'\n'
                    except Exception:
                        pass
    
    if not content_stream:
        return []
    
    # Content Stream 파싱
    parser = ContentStreamParser(font_map)
    raw_items = parser.parse(content_stream)
    
    if not raw_items:
        return []
    
    # =================================================================
    # 좌표계 정규화 (Normalization to Top-Down)
    # =================================================================
    
    # 좌표계 방향 감지 (Auto-Detect)
    is_bottom_up = _detect_coordinate_direction(raw_items)
    
    # 좌표 범위 확인
    max_y = max(t.y for t in raw_items)
    
    # 1. 스케일 정규화 (큰 좌표계 대응)
    # 일반적인 A4 페이지 높이: 842pt
    if max_y > 1500:
        scale = max_y / 842.0
        for item in raw_items:
            item.x /= scale
            item.y /= scale
            item.font_size /= scale
        max_y = 842.0  # 스케일 후 최대값
    
    # 2. Y축 방향 변환 (Bottom-Up → Top-Down)
    if is_bottom_up:
        for item in raw_items:
            item.y = max_y - item.y
    
    # 최종 정렬 (Reading Order: Y 오름차순 → X 오름차순)
    raw_items.sort(key=lambda t: (t.y, t.x))
    
    return raw_items


def _detect_coordinate_direction(text_items: list) -> bool:
    """
    텍스트 좌표계의 Y축 방향을 감지
    
    Returns:
        True: Bottom-Up (표준 PDF) - Y가 클수록 위쪽
        False: Top-Down - Y가 작을수록 위쪽
    """
    if len(text_items) < 2:
        return True  # 기본값: Bottom-Up
    
    # 투표 시스템
    # 텍스트가 읽기 순서로 배열되어 있다고 가정
    # 다음 줄로 갈 때 Y가 작아지면 Bottom-Up (다음 줄이 위로 올라감 = Y 감소)
    vote_bottom_up = 0
    vote_top_down = 0
    
    sample_size = min(len(text_items) - 1, 50)
    for i in range(sample_size):
        curr_y = text_items[i].y
        next_y = text_items[i + 1].y
        diff = next_y - curr_y
        
        if abs(diff) > 5:  # 같은 줄이 아닐 때만
            if diff < 0:  # 다음 줄의 Y가 작아짐 -> Bottom-Up
                vote_bottom_up += 1
            else:  # 다음 줄의 Y가 커짐 -> Top-Down
                vote_top_down += 1
    
    return vote_bottom_up > vote_top_down


def _clean_punctuation_spacing(text: str) -> str:
    """
    PDF 텍스트 추출 후 구두점/기호 주위의 불필요한 공백을 정리
    
    한글 PDF에서 각 텍스트 아이템이 후행 공백을 포함하여
    구두점 앞에 불필요한 공백이 생기는 문제를 해결
    """
    import re
    
    # 1. 구두점 앞 공백 제거: "했다 ." → "했다."  "법무부 ," → "법무부,"
    text = re.sub(r' ([,.\!?;:])', r'\1', text)
    
    # 2. 닫는 괄호/따옴표 류 앞 공백 제거: "시행계획 」" → "시행계획」"
    text = re.sub(r' ([」\)）\]】』"\'\'`」])', r'\1', text)
    
    # 3. 여는 괄호/따옴표 류 뒤 공백 제거: 「 '26년 → 「'26년  ( 이하 → (이하
    text = re.sub(r'([「\(（\[【『"\'\'`]]) ', r'\1', text)
    
    # 4. 가운뎃점(·) 앞뒤 공백 제거: "수립 ·의결" → "수립·의결"
    text = re.sub(r' ·', '·', text)
    text = re.sub(r'· ', '·', text)
    
    # 5. 붙임표(-) 앞뒤 공백 - 단어 사이의 하이픈만 (줄 시작 - 제외)
    # "치료 ·사회재활" 같은 패턴은 4번에서 처리됨
    
    return text


def _extract_text_single_column(items: list) -> str:
    """
    기존 1단 텍스트 추출 로직 (하위 호환용)
    
    items는 이미 Top-Down으로 정규화된 TextItem 리스트.
    Y→X 순 정렬 후 같은 Y(±5pt)를 한 줄로 합침.
    """
    sorted_items = sorted(items, key=lambda t: (t.y, t.x))
    
    # 같은 줄의 글자들을 합치기
    lines = []
    current_line = []
    current_y = None
    
    for item in sorted_items:
        if current_y is None or abs(item.y - current_y) > 5:
            if current_line:
                lines.append(current_line)
            current_line = [item]
            current_y = item.y
        else:
            current_line.append(item)
    
    if current_line:
        lines.append(current_line)
    
    # 각 줄에서 가까운 글자들을 합쳐서 단어로 만들기
    result = []
    for line in lines:
        line.sort(key=lambda t: t.x)
        
        # font_size가 비정상적으로 작으면(≤2pt) 실제 글자 간격에서 추정
        line_font_size = max((item.font_size for item in line if item.text), default=1.0)
        if line_font_size <= 2.0 and len(line) >= 3:
            # 한글 연속 글자 간격의 중앙값으로 font_size 추정
            cjk_gaps = []
            for i in range(1, len(line)):
                if line[i].text and line[i-1].text:
                    prev_ch = line[i-1].text[-1]
                    curr_ch = line[i].text[0]
                    if ord(prev_ch) > 0x2E7F and ord(curr_ch) > 0x2E7F:
                        gap = line[i].x - line[i-1].x
                        if gap > 0:
                            cjk_gaps.append(gap)
            if cjk_gaps:
                import statistics
                line_font_size = statistics.median(cjk_gaps)
        
        words = []
        current_word = ""
        prev_x = None
        
        for item in line:
            text = item.text
            if not text:
                continue
            
            effective_fs = max(item.font_size, line_font_size) if item.font_size <= 2.0 else item.font_size
            
            if prev_x is not None:
                gap = item.x - prev_x
                # 한글(CJK): char_width ≈ font_size, threshold = font_size * 1.2
                # 영문: char_width ≈ font_size * 0.6, threshold = font_size * 0.9
                prev_char = current_word[-1] if current_word else ''
                if prev_char and ord(prev_char) > 0x2E7F:
                    threshold = effective_fs * 1.2
                else:
                    threshold = effective_fs * 0.9
                
                if gap > threshold:
                    if current_word.strip():
                        words.append(current_word)
                    current_word = text
                else:
                    current_word += text
            else:
                current_word = text
            
            text_width = 0
            for ch in text:
                if ord(ch) > 0x2E7F:
                    text_width += effective_fs * 0.9
                else:
                    text_width += effective_fs * 0.45
            prev_x = item.x + text_width
        
        if current_word.strip():
            words.append(current_word)
        
        if words:
            parts = []
            for w in words:
                if parts and not parts[-1].endswith(' ') and not w.startswith(' '):
                    parts.append(' ')
                parts.append(w)
            line_text = ''.join(parts)
        else:
            line_text = ''
        
        if line_text.strip():
            result.append(line_text)
    
    return '\n'.join(result)


def _extract_page_lines(doc: PDFDocument, page_num: int):
    """Content stream에서 수직/수평 라인(선) 추출 (표 경계 감지용)
    
    CTM(Current Transformation Matrix)을 추적하여 실제 페이지 좌표로 변환.
    PDF content stream에서 `cm` 명령으로 좌표계가 변환된 후 
    `0 0 m 486 0 l` 같은 상대 좌표로 선을 그리는 패턴을 처리.
    """
    import re as _re
    pages_list = get_pages(doc)
    if page_num >= len(pages_list):
        return [], []
    page = pages_list[page_num]
    contents_ref = page.get('Contents')
    cs = b''
    if isinstance(contents_ref, PDFRef):
        obj = doc.objects.get((contents_ref.obj_num, contents_ref.gen_num))
        if obj:
            try:
                cs = decode_stream(doc, obj)
            except Exception:
                return [], []
    elif isinstance(contents_ref, list):
        for ref in contents_ref:
            if isinstance(ref, PDFRef):
                obj = doc.objects.get((ref.obj_num, ref.gen_num))
                if obj:
                    try:
                        cs += decode_stream(doc, obj) + b'\n'
                    except Exception:
                        pass

    h_lines = []  # (y, x_min, x_max) - PDF 좌표계
    v_lines = []  # (x, y_min, y_max) - PDF 좌표계

    # CTM 스택 추적
    # CTM = [a, b, c, d, e, f] where:
    # x' = a*x + c*y + e
    # y' = b*x + d*y + f
    ctm_stack = []
    ctm = [1, 0, 0, 1, 0, 0]  # identity
    
    def apply_ctm(x, y):
        """CTM을 적용하여 페이지 좌표로 변환"""
        x2 = ctm[0] * x + ctm[2] * y + ctm[4]
        y2 = ctm[1] * x + ctm[3] * y + ctm[5]
        return x2, y2
    
    def multiply_ctm(new_ctm):
        """현재 CTM에 새 변환 행렬을 곱함"""
        a, b, c, d, e, f = ctm
        a2, b2, c2, d2, e2, f2 = new_ctm
        return [
            a*a2 + b*c2, a*b2 + b*d2,
            c*a2 + d*c2, c*b2 + d*d2,
            e*a2 + f*c2 + e2, e*b2 + f*d2 + f2
        ]
    
    # 토큰 단위로 파싱
    text = cs.decode('latin-1', errors='replace')
    
    # q/Q (save/restore), cm, m, l, re, S/s/f/F 명령어 처리
    cur_x, cur_y = 0.0, 0.0  # current point
    
    # 간단한 토큰 기반 파서
    tokens = []
    for tok in _re.finditer(rb'[-+]?\d*\.?\d+|[a-zA-Z*]+|\[|\]|\((?:[^\\)]|\\.)*\)|<[^>]*>', cs):
        tokens.append(tok.group())
    
    num_stack = []
    for tok in tokens:
        # 숫자인지 확인
        try:
            num = float(tok)
            num_stack.append(num)
            continue
        except (ValueError, UnicodeDecodeError):
            pass
        
        op = tok.decode('latin-1', errors='replace') if isinstance(tok, bytes) else tok
        
        if op == 'q':
            ctm_stack.append(ctm[:])
        elif op == 'Q':
            if ctm_stack:
                ctm = ctm_stack.pop()
        elif op == 'cm' and len(num_stack) >= 6:
            new_ctm = num_stack[-6:]
            ctm = multiply_ctm(new_ctm)
            num_stack = num_stack[:-6]
        elif op == 'm' and len(num_stack) >= 2:
            cur_x, cur_y = num_stack[-2], num_stack[-1]
            num_stack = num_stack[:-2]
        elif op == 'l' and len(num_stack) >= 2:
            lx, ly = num_stack[-2], num_stack[-1]
            num_stack = num_stack[:-2]
            # CTM 적용
            px1, py1 = apply_ctm(cur_x, cur_y)
            px2, py2 = apply_ctm(lx, ly)
            if abs(py1 - py2) < 1:  # 수평선
                h_lines.append((py1, min(px1, px2), max(px1, px2)))
            elif abs(px1 - px2) < 1:  # 수직선
                v_lines.append((px1, min(py1, py2), max(py1, py2)))
            cur_x, cur_y = lx, ly
        elif op == 're' and len(num_stack) >= 4:
            rx, ry, rw, rh = num_stack[-4:]
            num_stack = num_stack[:-4]
            # 얇은 사각형 → 선으로 처리
            if abs(rh) < 2 and abs(rw) > 5:  # 수평 사각형 → 수평선
                px, py = apply_ctm(rx, ry + rh/2)
                px2, _ = apply_ctm(rx + rw, ry + rh/2)
                h_lines.append((py, min(px, px2), max(px, px2)))
            elif abs(rw) < 2 and abs(rh) > 5:  # 수직 사각형 → 수직선
                px, py = apply_ctm(rx + rw/2, ry)
                _, py2 = apply_ctm(rx + rw/2, ry + rh)
                v_lines.append((px, min(py, py2), max(py, py2)))
        elif op in ('S', 's', 'f', 'F', 'B', 'b'):
            pass  # stroke/fill - 이미 위에서 선 처리됨
        else:
            # 그 외 텍스트/기타 명령은 num_stack 초기화
            if op in ('Tf', 'Td', 'TD', 'Tm', 'TJ', 'Tj', 'T*',
                      'w', 'd', 'J', 'j', 'M', 'ri', 'gs',
                      'g', 'G', 'rg', 'RG', 'k', 'K', 'cs', 'CS',
                      'sc', 'SC', 'scn', 'SCN', 'Do', 'BT', 'ET',
                      'n', 'W', 'h', 'c', 'v', 'y'):
                num_stack.clear()

    return h_lines, v_lines


def _find_table_col_separator(v_lines, page_width):
    """수직선에서 페이지 중앙 근처의 주요 열 분리선 X좌표를 반환.
    
    조건: 페이지 중앙 ±30% 이내, 3회 이상 출현, 총 span ≥ 200pt.
    이 조건을 만족하는 X가 없으면 None 반환 (표 열 분리 없음).
    """
    if not v_lines:
        return None

    from collections import Counter
    x_counts = Counter(round(l[0]) for l in v_lines)

    center = page_width / 2
    margin = page_width * 0.3
    best = None

    for x, cnt in x_counts.items():
        if not (center - margin < x < center + margin):
            continue
        if cnt < 3:
            continue
        total_span = sum(l[2] - l[1] for l in v_lines if abs(round(l[0]) - x) < 2)
        if total_span < 200:
            continue
        if best is None or total_span > best[1]:
            best = (x, total_span)

    return best[0] if best else None


def _extract_text_table_columns(items, col_boundaries, h_lines, page_height, col_y_range=None, v_lines=None, page_width=None):
    """수직선이 존재하는 Y영역만 컬럼 분리, 나머지는 단일 컬럼 처리.
    
    표 유형에 따라 두 가지 모드로 동작:
    - 대조표 (컬럼 경계 1~2개): 좌 컬럼 전체 → 우 컬럼 전체
    - 일반 표 (컬럼 경계 3개 이상): 행 단위 좌→우 순서
      행 경계는 수직선 Y 끝점 기반으로 정확한 셀 경계를 사용
    """
    if not isinstance(col_boundaries, (list, tuple)):
        col_boundaries = [col_boundaries]
    col_boundaries = sorted(col_boundaries)
    
    if not col_boundaries or not items:
        return _extract_text_single_column(items)
    
    # 컬럼 분리할 Y범위 결정
    if col_y_range:
        split_y_top, split_y_bot = col_y_range
    else:
        h_ys_pdf = sorted(set(round(l[1]) for l in h_lines))
        if len(h_ys_pdf) < 3:
            return _extract_text_single_column(items)
        split_y_top = page_height - max(h_ys_pdf)
        split_y_bot = page_height - min(h_ys_pdf)
    
    # 아이템을 3구간으로 분류
    above, split_items, below = [], [], []
    for item in items:
        if item.y < split_y_top - 5:
            above.append(item)
        elif item.y > split_y_bot + 5:
            below.append(item)
        else:
            split_items.append(item)
    
    parts = []
    
    # 상단: 일반 텍스트
    if above:
        parts.append(_extract_text_single_column(above))
    
    # 표 영역 처리
    if split_items:
        is_comparison_table = len(col_boundaries) <= 2
        
        if is_comparison_table:
            # ── 대조표 모드: 좌 컬럼 전체 → 우 컬럼 전체 ──
            num_cols = len(col_boundaries) + 1
            columns = [[] for _ in range(num_cols)]
            
            for item in split_items:
                col_idx = len(col_boundaries)
                for ci, b in enumerate(col_boundaries):
                    if item.x < b:
                        col_idx = ci
                        break
                columns[col_idx].append(item)
            
            col_texts = []
            for col_items in columns:
                if col_items:
                    ct = _extract_text_single_column(col_items).strip()
                    if ct:
                        col_texts.append(ct)
            
            if col_texts:
                parts.append('\n'.join(col_texts))
        else:
            # ── 일반 표 모드: 행 단위 좌→우 순서 ──
            # 행 경계: 수직선 Y 끝점 기반 (가장 정확한 셀 경계)
            row_boundaries = _compute_row_boundaries_from_vlines(
                v_lines, page_height, page_width, split_y_top, split_y_bot
            ) if v_lines and page_width else []
            
            # v_lines 기반 행 경계가 부족하면 수평선/자동 행 경계로 fallback
            if len(row_boundaries) < 3:
                h_ys_pdf = sorted(set(round(l[1]) for l in h_lines))
                row_ys_td = sorted(set(
                    round(page_height - y) for y in h_ys_pdf
                    if split_y_top - 5 <= page_height - y <= split_y_bot + 5
                ))
                auto_boundaries = _compute_row_boundaries_from_items(split_items)
                
                if len(row_ys_td) >= 3 and len(row_ys_td) >= len(auto_boundaries):
                    row_boundaries = row_ys_td
                elif len(auto_boundaries) >= 2:
                    row_boundaries = auto_boundaries
                else:
                    row_boundaries = row_ys_td if len(row_ys_td) >= 2 else auto_boundaries
            
            if len(row_boundaries) >= 2:
                assigned = set()
                row_texts = []
                for r in range(len(row_boundaries) - 1):
                    rtop = row_boundaries[r]
                    rbot = row_boundaries[r + 1]
                    row_items = []
                    for i, item in enumerate(split_items):
                        if i in assigned:
                            continue
                        if rtop - 1 <= item.y <= rbot + 1:
                            row_items.append(item)
                            assigned.add(i)
                    
                    if not row_items:
                        continue
                    
                    columns = [[] for _ in range(len(col_boundaries) + 1)]
                    for item in row_items:
                        col_idx = len(col_boundaries)
                        for ci, b in enumerate(col_boundaries):
                            if item.x < b:
                                col_idx = ci
                                break
                        columns[col_idx].append(item)
                    
                    # 희소 컬럼 병합: 아이템이 적은(<=2) 컬럼을 좌측 인접 컬럼에 합침
                    # (텍스트가 컬럼 경계를 약간 넘어가는 경우 보정)
                    for ci in range(len(columns) - 1, 0, -1):
                        if len(columns[ci]) <= 2 and len(columns[ci - 1]) > 2:
                            columns[ci - 1].extend(columns[ci])
                            columns[ci] = []
                    
                    cell_texts = []
                    for col_items in columns:
                        if col_items:
                            ct = _extract_text_single_column(col_items).strip()
                            if ct:
                                # 셀 내부 줄바꿈을 공백으로 치환 (셀 간 구분만 \n 사용)
                                ct = ' '.join(ct.split('\n'))
                                cell_texts.append(ct)
                    
                    if cell_texts:
                        row_texts.append('\n'.join(cell_texts))
                
                remaining = [split_items[i] for i in range(len(split_items)) if i not in assigned]
                if remaining:
                    row_texts.append(_extract_text_single_column(remaining).strip())
                
                if row_texts:
                    parts.append('\n'.join(row_texts))
            else:
                parts.append(_extract_text_single_column(split_items))
    
    # 하단: 일반 텍스트
    if below:
        parts.append(_extract_text_single_column(below))
    
    return '\n'.join(p for p in parts if p.strip())


def _compute_row_boundaries_from_vlines(v_lines, page_height, page_width, split_y_top, split_y_bot):
    """수직선 Y 끝점으로 행 경계 계산. 표의 실제 셀 경계와 정확히 일치."""
    if not v_lines or not page_width:
        return []
    
    v_endpoints = set()
    for x, y0, y1 in v_lines:
        xk = round(x)
        if page_width * 0.12 < xk < page_width * 0.88:
            td0 = round(page_height - y0)
            td1 = round(page_height - y1)
            if split_y_top - 5 <= td0 <= split_y_bot + 5:
                v_endpoints.add(td0)
            if split_y_top - 5 <= td1 <= split_y_bot + 5:
                v_endpoints.add(td1)
    
    return sorted(v_endpoints)


def _compute_row_boundaries_from_items(items):
    """수평선이 없을 때 아이템 Y좌표 기반으로 행 경계를 자동 계산.
    
    Y좌표를 정렬 후 큰 갭(>= font_size * 1.5)을 행 경계로 사용.
    """
    if not items:
        return []
    
    # 평균 폰트 크기 계산
    font_sizes = [getattr(it, 'font_size', 10) or 10 for it in items]
    avg_fs = sum(font_sizes) / len(font_sizes)
    gap_threshold = avg_fs * 1.5
    
    # Y좌표 수집 및 정렬
    y_vals = sorted(set(round(it.y, 1) for it in items))
    if len(y_vals) < 2:
        return [min(it.y for it in items) - 1, max(it.y for it in items) + avg_fs + 1]
    
    boundaries = [y_vals[0] - 1]
    for i in range(1, len(y_vals)):
        if y_vals[i] - y_vals[i-1] > gap_threshold:
            # 갭 중간점을 경계로
            boundaries.append((y_vals[i-1] + y_vals[i]) / 2)
    boundaries.append(y_vals[-1] + avg_fs + 1)
    
    return boundaries


def extract_text(doc: PDFDocument, page_num: int = 0) -> str:
    """
    특정 페이지에서 텍스트 추출
    
    2단(multi-column) 레이아웃 감지 시 layout_analyzer를 사용하여
    올바른 읽기 순서(left column → right column)로 텍스트를 반환합니다.
    1단 문서는 기존 로직과 동일한 결과를 반환합니다.
    
    Args:
        doc: PDFDocument 객체
        page_num: 페이지 번호 (0부터 시작)
    
    Returns:
        str: 추출된 텍스트 (읽기 순서로 정렬)
    """
    items = extract_text_with_positions(doc, page_num)
    
    if not items:
        return ""
    
    page_w, page_h = _get_page_dimensions(doc, page_num)
    
    # ─────────────────────────────────────────────────────────
    # 1차: 표 라인 기반 열 분리 (content stream의 수직선 감지)
    #   "개정 전/개정 후" 같은 다중 열 표에서 컬럼별 텍스트 분리
    # ─────────────────────────────────────────────────────────
    try:
        h_lines, v_lines = _extract_page_lines(doc, page_num)
        col_sep = _find_table_col_separator(v_lines, page_w)
        # 주요 수직선 경계 수집: 표 높이의 20% 이상을 관통하는 선만
        inner_col_xs = []
        if v_lines and h_lines:
            from litparser._grid_table import _cluster_values
            from collections import defaultdict
            h_ys_pdf = sorted(set(round(l[1]) for l in h_lines))
            table_height = max(h_ys_pdf) - min(h_ys_pdf) if len(h_ys_pdf) >= 2 else page_h
            # X 클러스터별 총 높이 계산
            x_heights = defaultdict(float)
            for x, y0, y1 in v_lines:
                x_key = round(x)
                x_heights[x_key] += (y1 - y0)
            # 표 높이의 20% 이상 관통 + 페이지 가장자리 제외
            min_height = table_height * 0.2
            for x_key, total_h in x_heights.items():
                if total_h >= min_height and page_w * 0.12 < x_key < page_w * 0.88:
                    inner_col_xs.append(float(x_key))
            inner_col_xs = sorted(inner_col_xs)
            
            # ── 텍스트 crossing 검증 ──
            # 텍스트가 경계를 60% 이상 가로지르면 가짜 경계(셀 병합)로 판단
            # 외곽 수직선(가장 왼쪽/오른쪽)은 표의 구조적 경계이므로 항상 유지
            if len(inner_col_xs) >= 3 and items:
                from collections import defaultdict as _dd
                h_ys_pdf_set = sorted(set(round(l[1]) for l in h_lines))
                _ty_top = page_h - max(h_ys_pdf_set) if h_ys_pdf_set else 0
                _ty_bot = page_h - min(h_ys_pdf_set) if h_ys_pdf_set else page_h
                _titems = [it for it in items if _ty_top - 5 <= it.y <= _ty_bot + 5]
                _ygroups = _dd(list)
                for it in _titems:
                    _ygroups[round(it.y / 3) * 3].append(it)
                
                _first = inner_col_xs[0]
                _last = inner_col_xs[-1]
                validated_xs = []
                for bx in inner_col_xs:
                    # 외곽 경계는 항상 유지
                    if bx == _first or bx == _last:
                        validated_xs.append(bx)
                        continue
                    crossing = 0
                    total = 0
                    for gy, group in _ygroups.items():
                        if len(group) < 2:
                            continue
                        total += 1
                        has_left = any(it.x < bx - 2 for it in group)
                        has_right = any(it.x > bx + 2 for it in group)
                        if has_left and has_right:
                            crossing += 1
                    # 60% 이상 crossing → 가짜 경계 제거
                    if total > 0 and crossing > total * 0.6:
                        continue
                    validated_xs.append(bx)
                inner_col_xs = validated_xs
    except Exception:
        h_lines, v_lines, col_sep, inner_col_xs = [], [], None, []
    
    if col_sep is not None and h_lines:
        # 표 판정 강화: v_lines가 h_lines보다 과도하게 많으면 표가 아님 (코드 Listing 등)
        # 일반 표: h_lines와 v_lines가 비슷한 비율
        if len(v_lines) > len(h_lines) * 10 and len(h_lines) < 15:
            col_sep = None  # 표가 아닌 것으로 판정
    
    if col_sep is not None and h_lines:
        # col_sep 수직선이 커버하는 Y범위 (top-down) 계산
        col_sep_y_top_td = page_h
        col_sep_y_bot_td = 0
        for x, y0, y1 in v_lines:
            if abs(x - col_sep) < 5:
                td_top = page_h - y1
                td_bot = page_h - y0
                col_sep_y_top_td = min(col_sep_y_top_td, td_top)
                col_sep_y_bot_td = max(col_sep_y_bot_td, td_bot)
        
        col_boundaries = inner_col_xs if len(inner_col_xs) >= 2 else [col_sep]
        col_y_range = (col_sep_y_top_td, col_sep_y_bot_td) if col_sep_y_bot_td > col_sep_y_top_td else None
        text = _extract_text_table_columns(items, col_boundaries, h_lines, page_h, col_y_range, v_lines=v_lines, page_width=page_w)
        text = _clean_punctuation_spacing(text)
        return text
    
    # ─────────────────────────────────────────────────────────
    # 2차: 레이아웃 분석 - 2단 이상이면 layout_analyzer 경유
    # ─────────────────────────────────────────────────────────
    # 최소 아이템 수: 아이템이 너무 적으면 2단 감지 불가
    MIN_ITEMS_FOR_COLUMN_DETECT = 40
    
    layout = None
    if len(items) >= MIN_ITEMS_FOR_COLUMN_DETECT:
        try:
            layout = analyze_layout(items, page_w, page_h)
        except Exception:
            layout = None
    
    if layout and layout.num_columns >= 2 and layout.blocks:
        # 2단 이상: layout blocks의 reading order 순으로 텍스트 반환
        text = '\n'.join(b.text for b in layout.blocks if b.text.strip())
    else:
        # 1단: 기존 로직 그대로 (하위 호환 보장)
        text = _extract_text_single_column(items)
    
    # 후처리: 구두점/기호 주위 불필요한 공백 제거
    text = _clean_punctuation_spacing(text)
    
    return text


def extract_all_text(doc: PDFDocument) -> str:
    """모든 페이지에서 텍스트 추출"""
    page_count = get_page_count(doc)
    all_text = []
    
    for i in range(page_count):
        text = extract_text(doc, i)
        if text:
            all_text.append(f"--- Page {i+1} ---\n{text}")
    
    return '\n\n'.join(all_text)


def extract_tables(doc: PDFDocument, page_num: int = 0, **kwargs) -> list:
    """
    페이지에서 테이블 추출
    
    Args:
        doc: PDFDocument 객체
        page_num: 페이지 번호 (0부터 시작)
        **kwargs: detect_tables 파라미터 (min_rows, min_cols, x_tolerance, y_tolerance)
    
    Returns:
        List[Table]: 감지된 테이블 목록
    
    Example:
        tables = extract_tables(doc, 0)
        for table in tables:
            print(table.to_markdown())
    """
    from .core.table_detector import extract_tables_from_page
    return extract_tables_from_page(doc, page_num, **kwargs)


def get_document_outline(doc: PDFDocument) -> list:
    """
    문서 개요 (헤딩 목록) 추출 (Tagged PDF)
    
    Args:
        doc: PDFDocument 객체
    
    Returns:
        List[Tuple[int, str]]: (헤딩 레벨, 텍스트) 리스트
    
    Example:
        outline = get_document_outline(doc)
        for level, text in outline:
            print('  ' * (level-1) + f'H{level}: {text}')
    """
    parser = StructTreeParser(doc)
    return parser.get_document_outline()


def get_document_structure(doc: PDFDocument):
    """
    문서 구조 트리 반환 (Tagged PDF)
    
    Args:
        doc: PDFDocument 객체
    
    Returns:
        StructElement: 문서 구조 루트 또는 None
    """
    parser = StructTreeParser(doc)
    return parser.parse()