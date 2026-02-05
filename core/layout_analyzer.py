"""
PDF Layout Analyzer - Stage 5

PDF 페이지의 레이아웃 구조 분석

기능:
1. 텍스트 블록 감지 (단락, 제목 등)
2. 읽기 순서 결정
3. 컬럼 레이아웃 감지
4. 헤더/푸터 감지
"""

from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Any
from collections import defaultdict
from enum import Enum


class BlockType(Enum):
    """텍스트 블록 유형"""
    TITLE = "title"
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    LIST_ITEM = "list_item"
    HEADER = "header"
    FOOTER = "footer"
    CAPTION = "caption"
    TABLE = "table"
    UNKNOWN = "unknown"


@dataclass
class TextBlock:
    """텍스트 블록"""
    block_type: BlockType
    text: str
    x: float
    y: float
    width: float
    height: float
    font_size: float = 12.0
    font_name: str = ""
    items: List[Any] = field(default_factory=list)  # 원본 TextItem들
    
    @property
    def bbox(self) -> Tuple[float, float, float, float]:
        """바운딩 박스 (x1, y1, x2, y2)"""
        return (self.x, self.y, self.x + self.width, self.y + self.height)
    
    @property
    def center(self) -> Tuple[float, float]:
        """중심점"""
        return (self.x + self.width / 2, self.y + self.height / 2)


@dataclass
class PageLayout:
    """페이지 레이아웃 정보"""
    width: float
    height: float
    blocks: List[TextBlock] = field(default_factory=list)
    columns: int = 1  # 컬럼 수
    has_header: bool = False
    has_footer: bool = False
    header_blocks: List[TextBlock] = field(default_factory=list)
    footer_blocks: List[TextBlock] = field(default_factory=list)
    
    def get_reading_order(self) -> List[TextBlock]:
        """읽기 순서대로 블록 반환"""
        # 헤더 -> 본문 (위→아래, 왼→오) -> 푸터
        result = list(self.header_blocks)
        
        # 본문 블록 정렬
        body_blocks = [b for b in self.blocks 
                       if b not in self.header_blocks and b not in self.footer_blocks]
        
        if self.columns > 1:
            # 다중 컬럼: 컬럼별로 정렬
            result.extend(self._sort_multi_column(body_blocks))
        else:
            # 단일 컬럼: Y 좌표 정렬
            result.extend(sorted(body_blocks, key=lambda b: (b.y, b.x)))
        
        result.extend(self.footer_blocks)
        return result
    
    def _sort_multi_column(self, blocks: List[TextBlock]) -> List[TextBlock]:
        """다중 컬럼 정렬"""
        if not blocks:
            return []
        
        # X 좌표로 컬럼 구분
        col_width = self.width / self.columns
        columns = defaultdict(list)
        
        for block in blocks:
            col_idx = int(block.x / col_width)
            columns[col_idx].append(block)
        
        # 각 컬럼 내에서 Y 정렬 후 합치기
        result = []
        for col_idx in sorted(columns.keys()):
            col_blocks = sorted(columns[col_idx], key=lambda b: b.y)
            result.extend(col_blocks)
        
        return result
    
    def to_text(self) -> str:
        """읽기 순서대로 텍스트 추출"""
        blocks = self.get_reading_order()
        return '\n\n'.join(b.text for b in blocks if b.text.strip())


def analyze_layout(text_items: list, 
                   page_width: float = 612,
                   page_height: float = 792) -> PageLayout:
    """
    텍스트 항목에서 레이아웃 분석
    
    Args:
        text_items: TextItem 리스트
        page_width: 페이지 너비
        page_height: 페이지 높이
    
    Returns:
        PageLayout: 분석된 레이아웃
    """
    if not text_items:
        return PageLayout(width=page_width, height=page_height)
    
    # 1. 텍스트 블록 생성 (근접한 항목 그룹화)
    blocks = _create_text_blocks(text_items)
    
    # 2. 블록 유형 분류
    blocks = _classify_blocks(blocks, page_width, page_height)
    
    # 3. 헤더/푸터 감지
    header_blocks, footer_blocks, body_blocks = _detect_header_footer(
        blocks, page_height
    )
    
    # 4. 컬럼 수 감지
    columns = _detect_columns(body_blocks, page_width)
    
    return PageLayout(
        width=page_width,
        height=page_height,
        blocks=blocks,
        columns=columns,
        has_header=len(header_blocks) > 0,
        has_footer=len(footer_blocks) > 0,
        header_blocks=header_blocks,
        footer_blocks=footer_blocks
    )


def _create_text_blocks(text_items: list, 
                        x_gap: float = 20,
                        y_gap: float = 15) -> List[TextBlock]:
    """근접한 텍스트 항목을 블록으로 그룹화"""
    if not text_items:
        return []
    
    # Y 좌표로 정렬
    sorted_items = sorted(text_items, key=lambda t: (t.y, t.x))
    
    blocks = []
    current_items = [sorted_items[0]]
    current_y = sorted_items[0].y
    
    for item in sorted_items[1:]:
        # Y 간격이 크면 새 블록
        if abs(item.y - current_y) > y_gap:
            if current_items:
                blocks.append(_items_to_block(current_items))
            current_items = [item]
            current_y = item.y
        else:
            current_items.append(item)
            current_y = item.y
    
    if current_items:
        blocks.append(_items_to_block(current_items))
    
    return blocks


def _items_to_block(items: list) -> TextBlock:
    """TextItem 리스트를 TextBlock으로 변환"""
    if not items:
        return TextBlock(
            block_type=BlockType.UNKNOWN,
            text="",
            x=0, y=0, width=0, height=0
        )
    
    # 바운딩 박스 계산
    min_x = min(item.x for item in items)
    min_y = min(item.y for item in items)
    max_x = max(item.x + item.font_size * len(item.text) * 0.5 for item in items)
    max_y = max(item.y + item.font_size for item in items)
    
    # 평균 폰트 크기
    avg_font_size = sum(item.font_size for item in items) / len(items)
    
    # 텍스트 결합 (같은 줄끼리)
    lines = defaultdict(list)
    for item in items:
        line_y = round(item.y / 5) * 5  # Y 좌표 그룹화
        lines[line_y].append(item)
    
    text_parts = []
    for line_y in sorted(lines.keys()):
        line_items = sorted(lines[line_y], key=lambda t: t.x)
        line_text = ' '.join(item.text for item in line_items)
        text_parts.append(line_text)
    
    return TextBlock(
        block_type=BlockType.UNKNOWN,
        text='\n'.join(text_parts),
        x=min_x,
        y=min_y,
        width=max_x - min_x,
        height=max_y - min_y,
        font_size=avg_font_size,
        font_name=items[0].font_name if items else "",
        items=items
    )


def _classify_blocks(blocks: List[TextBlock], 
                     page_width: float,
                     page_height: float) -> List[TextBlock]:
    """블록 유형 분류"""
    if not blocks:
        return blocks
    
    # 폰트 크기 분포 분석
    font_sizes = [b.font_size for b in blocks if b.font_size > 0]
    if not font_sizes:
        return blocks
    
    avg_font_size = sum(font_sizes) / len(font_sizes)
    max_font_size = max(font_sizes)
    
    for block in blocks:
        # 큰 폰트 = 제목/헤딩
        if block.font_size >= max_font_size * 0.9 and block.font_size > avg_font_size * 1.3:
            if len(block.text) < 100:
                block.block_type = BlockType.TITLE
            else:
                block.block_type = BlockType.HEADING
        
        # 불릿/번호로 시작 = 리스트
        elif _is_list_item(block.text):
            block.block_type = BlockType.LIST_ITEM
        
        # 중앙 정렬 + 짧은 텍스트 = 캡션
        elif _is_centered(block, page_width) and len(block.text) < 150:
            block.block_type = BlockType.CAPTION
        
        # 기본 = 단락
        else:
            block.block_type = BlockType.PARAGRAPH
    
    return blocks


def _is_list_item(text: str) -> bool:
    """리스트 항목인지 확인"""
    text = text.strip()
    if not text:
        return False
    
    # 불릿 기호
    bullets = ['•', '●', '○', '■', '□', '▪', '▫', '-', '*', '→', '➤', '►']
    if text[0] in bullets:
        return True
    
    # 번호 (1. 2. a. b. i. ii.)
    import re
    if re.match(r'^(\d+\.|[a-zA-Z]\.|[ivxIVX]+\.)\s', text):
        return True
    
    # 괄호 번호 (1) (a) 
    if re.match(r'^\(\d+\)|\([a-zA-Z]\)\s', text):
        return True
    
    return False


def _is_centered(block: TextBlock, page_width: float, tolerance: float = 50) -> bool:
    """중앙 정렬인지 확인"""
    block_center = block.x + block.width / 2
    page_center = page_width / 2
    return abs(block_center - page_center) < tolerance


def _detect_header_footer(blocks: List[TextBlock], 
                          page_height: float,
                          margin_ratio: float = 0.1) -> Tuple[List[TextBlock], List[TextBlock], List[TextBlock]]:
    """헤더와 푸터 감지"""
    header_margin = page_height * margin_ratio
    footer_margin = page_height * (1 - margin_ratio)
    
    header_blocks = []
    footer_blocks = []
    body_blocks = []
    
    for block in blocks:
        if block.y < header_margin:
            header_blocks.append(block)
            block.block_type = BlockType.HEADER
        elif block.y > footer_margin:
            footer_blocks.append(block)
            block.block_type = BlockType.FOOTER
        else:
            body_blocks.append(block)
    
    return header_blocks, footer_blocks, body_blocks


def _detect_columns(blocks: List[TextBlock], page_width: float) -> int:
    """컬럼 수 감지"""
    if not blocks or len(blocks) < 3:
        return 1
    
    # X 좌표 분포 분석
    x_positions = [b.x for b in blocks]
    
    # 클러스터링으로 컬럼 수 추정
    x_sorted = sorted(set(x_positions))
    
    if len(x_sorted) < 2:
        return 1
    
    # 큰 간격 찾기 (컬럼 구분선)
    gaps = []
    for i in range(1, len(x_sorted)):
        gap = x_sorted[i] - x_sorted[i-1]
        gaps.append((gap, x_sorted[i-1], x_sorted[i]))
    
    # 페이지 너비의 20% 이상 간격이 있으면 다중 컬럼
    significant_gaps = [g for g in gaps if g[0] > page_width * 0.15]
    
    return len(significant_gaps) + 1


def analyze_page_layout(doc, page_num: int = 0) -> PageLayout:
    """
    PDF 페이지 레이아웃 분석
    
    Args:
        doc: PDFDocument 객체
        page_num: 페이지 번호
    
    Returns:
        PageLayout: 분석된 레이아웃
    """
    from .. import extract_text_with_positions, get_pages
    
    # 페이지 크기 가져오기
    pages = get_pages(doc)
    if page_num >= len(pages):
        return PageLayout(width=612, height=792)
    
    page = pages[page_num]
    media_box = page.get('MediaBox', [0, 0, 612, 792])
    page_width = media_box[2] - media_box[0]
    page_height = media_box[3] - media_box[1]
    
    # 텍스트 추출
    text_items = extract_text_with_positions(doc, page_num)
    
    # 글자 단위 항목을 단어/문장 단위로 합치기
    text_items = _merge_text_items(text_items)
    
    return analyze_layout(text_items, page_width, page_height)


def _merge_text_items(items: list, gap_threshold: float = 5) -> list:
    """
    인접한 텍스트 항목을 합치기
    
    개별 글자를 단어/문장으로 합침
    """
    if not items:
        return []
    
    from dataclasses import dataclass
    
    @dataclass
    class MergedItem:
        text: str
        x: float
        y: float
        font_name: str
        font_size: float
    
    # Y 좌표로 그룹화 (같은 줄)
    from collections import defaultdict
    lines = defaultdict(list)
    
    for item in items:
        line_y = round(item.y / 3) * 3  # Y 좌표 그룹화
        lines[line_y].append(item)
    
    merged = []
    
    for line_y in sorted(lines.keys()):
        line_items = sorted(lines[line_y], key=lambda t: t.x)
        
        if not line_items:
            continue
        
        # 같은 줄의 항목들을 합치기
        current_text = line_items[0].text
        current_x = line_items[0].x
        current_font = line_items[0].font_name
        current_size = line_items[0].font_size
        prev_x = line_items[0].x
        
        for item in line_items[1:]:
            # 간격 계산
            gap = item.x - prev_x - current_size * 0.5
            
            if gap > current_size * 0.8:
                # 큰 간격 = 새 단어
                if current_text.strip():
                    merged.append(MergedItem(
                        text=current_text,
                        x=current_x,
                        y=line_y,
                        font_name=current_font,
                        font_size=current_size
                    ))
                current_text = item.text
                current_x = item.x
                current_font = item.font_name
                current_size = item.font_size
            else:
                # 작은 간격 = 같은 단어
                current_text += item.text
            
            prev_x = item.x
        
        # 마지막 단어
        if current_text.strip():
            merged.append(MergedItem(
                text=current_text,
                x=current_x,
                y=line_y,
                font_name=current_font,
                font_size=current_size
            ))
    
    return merged


# 테스트
if __name__ == '__main__':
    from dataclasses import dataclass
    
    @dataclass
    class MockTextItem:
        text: str
        x: float
        y: float
        font_name: str = "Arial"
        font_size: float = 12
    
    # 테스트 데이터: 제목 + 2단 레이아웃
    items = [
        # 제목
        MockTextItem("Document Title", 200, 50, font_size=24),
        # 왼쪽 컬럼
        MockTextItem("Left column text paragraph one.", 50, 100),
        MockTextItem("More text in the left column.", 50, 120),
        # 오른쪽 컬럼
        MockTextItem("Right column text paragraph.", 350, 100),
        MockTextItem("Additional right column content.", 350, 120),
        # 푸터
        MockTextItem("Page 1", 280, 750),
    ]
    
    layout = analyze_layout(items, 612, 792)
    
    print(f"페이지 크기: {layout.width}x{layout.height}")
    print(f"컬럼 수: {layout.columns}")
    print(f"헤더: {layout.has_header}")
    print(f"푸터: {layout.has_footer}")
    print(f"블록 수: {len(layout.blocks)}")
    
    print("\n읽기 순서:")
    for i, block in enumerate(layout.get_reading_order()):
        print(f"  {i+1}. [{block.block_type.value}] {block.text[:50]}...")
