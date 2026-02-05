"""
PDF Parser Library
직접 만드는 PDF 파싱 라이브러리

사용법:
    from pdf_parser import parse_pdf, extract_text
    
    doc = parse_pdf('document.pdf')
    text = extract_text(doc, page_num=0)
    
    # Stage 5: 이미지 추출
    images = extract_images(doc)
    
    # Stage 5: 테이블 감지
    tables = extract_tables(doc, page_num=0)
    
    # Stage 5: 레이아웃 분석
    layout = analyze_layout(doc, page_num=0)
    
    # Stage 5: 문서 구조 (Tagged PDF)
    from pdf_parser import is_tagged_pdf, get_document_outline
    if is_tagged_pdf(doc):
        outline = get_document_outline(doc)
"""
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

__version__ = '0.4.0'
__all__ = [
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
                cmap_data = decode_stream(doc, tounicode_obj)
                font_info.to_unicode = parse_tounicode_cmap(cmap_data)
        
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
            content_stream = decode_stream(doc, contents_obj)
    elif isinstance(contents_ref, list):
        # 여러 Content Stream 합치기
        for ref in contents_ref:
            if isinstance(ref, PDFRef):
                obj = doc.objects.get((ref.obj_num, ref.gen_num))
                if obj:
                    content_stream += decode_stream(doc, obj) + b'\n'
    
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


def extract_text(doc: PDFDocument, page_num: int = 0) -> str:
    """
    특정 페이지에서 텍스트 추출
    
    Args:
        doc: PDFDocument 객체
        page_num: 페이지 번호 (0부터 시작)
    
    Returns:
        str: 추출된 텍스트 (읽기 순서로 정렬)
    """
    items = extract_text_with_positions(doc, page_num)
    
    if not items:
        return ""
    
    # 이미 Top-Down으로 정규화되어 있으므로 Y 오름차순
    sorted_items = sorted(items, key=lambda t: (t.y, t.x))
    
    # 같은 줄의 글자들을 합치기
    lines = []
    current_line = []
    current_y = None
    
    for item in sorted_items:
        if current_y is None or abs(item.y - current_y) > 5:
            # 새로운 줄
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
        
        # 글자 간격 분석해서 단어 구분
        words = []
        current_word = ""
        prev_x = None
        prev_width = None
        
        for item in line:
            text = item.text
            if not text:
                continue
            
            # 예상 글자 너비 (폰트 크기 기반 추정)
            char_width = item.font_size * 0.6  # 대략적인 추정
            
            if prev_x is not None:
                gap = item.x - prev_x
                # 갭이 글자 너비의 1.5배 이상이면 단어 구분
                threshold = char_width * 1.5
                
                if gap > threshold:
                    if current_word.strip():
                        words.append(current_word)
                    current_word = text
                else:
                    current_word += text
            else:
                current_word = text
            
            prev_x = item.x + char_width  # 다음 글자 시작 예상 위치
        
        if current_word.strip():
            words.append(current_word)
        
        line_text = ' '.join(words)
        if line_text.strip():
            result.append(line_text)
    
    return '\n'.join(result)


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
