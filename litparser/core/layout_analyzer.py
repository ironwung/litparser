"""
PDF Layout Analyzer - Stage 8 (Adaptive Column Detection)

개선:
1. 표/그림 영역 분리 후 본문 컬럼 분석
2. 영역별 적응적 컬럼 감지
3. 주요 클러스터 기반 2단 감지
4. [Stage 8.1] 한글 텍스트 너비 계산 및 공백 삽입 개선
"""

from dataclasses import dataclass, field
from typing import List, Tuple, Any, Optional
from enum import Enum
from collections import Counter

class BlockType(Enum):
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
    text: str
    x: float
    y: float
    width: float
    height: float
    block_type: BlockType = BlockType.UNKNOWN
    font_size: float = 12.0
    items: List[Any] = field(default_factory=list)
    column: int = 0  # 0=full, 1=left, 2=right
    
    @property
    def bottom(self) -> float: return self.y + self.height
    @property
    def right(self) -> float: return self.x + self.width

@dataclass
class PageLayout:
    width: float
    height: float
    blocks: List[TextBlock] = field(default_factory=list)
    num_columns: int = 1
    
    def get_reading_order(self) -> List[TextBlock]:
        return self.blocks


def analyze_layout(text_items: list, page_width: float, page_height: float) -> PageLayout:
    if not text_items:
        return PageLayout(width=page_width, height=page_height)

    boxed_items = [_to_boxed_item(item) for item in text_items]
    
    # 1. 2단 본문 영역 찾기
    col_info, y_threshold = _find_two_column_region(boxed_items, page_width, page_height)
    num_columns = col_info['num_columns']
    
    if num_columns >= 2:
        left_max = col_info['left_col_max']
        right_min = col_info['right_col_min']
        
        # 아이템 분류
        full_width = []
        left_col = []
        right_col = []
        
        for b in boxed_items:
            # 본문 영역 이전은 전체폭 (표/그림)
            if b.y0 < y_threshold:
                full_width.append(b)
            # 본문 영역
            elif b.x0 < left_max:
                left_col.append(b)
            elif b.x0 >= right_min:
                right_col.append(b)
            else:
                full_width.append(b)
        
        all_blocks = []
        
        for items, col_num in [(full_width, 0), (left_col, 1), (right_col, 2)]:
            if items:
                groups = _y_cut_groups(items)
                for group in groups:
                    block = _create_text_block(group)
                    if block:
                        block.column = col_num
                        all_blocks.append(block)
        
        all_blocks = _sort_reading_order_with_table(all_blocks, y_threshold)
    else:
        groups = _y_cut_groups(boxed_items)
        all_blocks = []
        for group in groups:
            block = _create_text_block(group)
            if block:
                all_blocks.append(block)
        all_blocks.sort(key=lambda b: b.y)
    
    merged = _merge_adjacent_blocks(all_blocks)
    final = _classify_blocks(merged, page_height)
    
    return PageLayout(width=page_width, height=page_height, blocks=final, num_columns=num_columns)


def _find_two_column_region(items: List['BoxedItem'], page_width: float, page_height: float) -> tuple:
    """
    2단 본문 영역 찾기
    - 연속적으로 2단 구조가 유지되는 영역 탐지
    - 표 영역은 일시적 2단처럼 보이지만 연속성이 낮음
    """
    result = {'num_columns': 1, 'left_col_max': page_width/2, 'right_col_min': page_width/2}
    
    if len(items) < 10:
        return result, 0
    
    # Y 값들을 20pt 단위로 샘플링
    y_values = sorted(set(int(i.y0 / 20) * 20 for i in items))
    
    # 각 Y 시작점에서 2단 조건 체크
    def check_two_column(y_start, min_left_x=45, max_left_x=60, min_gap=250):
        region_items = [i for i in items if i.y0 >= y_start]
        if len(region_items) < 10:
            return None
        
        x_counter = Counter(int(i.x0 / 10) * 10 for i in region_items)
        major = sorted([(x, c) for x, c in x_counter.items() if c >= 5], key=lambda t: t[0])
        
        if len(major) < 2:
            return None
        
        center = page_width / 2
        for i in range(len(major) - 1):
            left_x, right_x = major[i][0], major[i+1][0]
            gap = right_x - left_x
            mid = (left_x + right_x) / 2
            
            if gap >= min_gap and min_left_x <= left_x <= max_left_x and abs(mid - center) < page_width * 0.25:
                return {'left_max': left_x + 60, 'right_min': right_x - 10, 'gap': gap}
        return None
    
    # 연속적인 2단 영역 찾기
    # 최소 3개 연속 Y 포인트에서 2단이어야 함
    consecutive_count = 0
    first_y = None
    col_info = None
    
    for y_start in y_values:
        info = check_two_column(y_start)
        if info and info['gap'] >= 250:  # 본문 2단은 갭이 큼
            if first_y is None:
                first_y = y_start
                col_info = info
            consecutive_count += 1
            
            # 3개 연속이면 본문 시작으로 인정
            if consecutive_count >= 3:
                result['num_columns'] = 2
                result['left_col_max'] = col_info['left_max']
                result['right_col_min'] = col_info['right_min']
                return result, first_y
        else:
            consecutive_count = 0
            first_y = None
            col_info = None
    
    return result, 0


def _sort_reading_order_with_table(blocks: List[TextBlock], y_threshold: float) -> List[TextBlock]:
    """표/그림 영역 고려한 읽기 순서"""
    # 상단 전체폭 (표 포함)
    top_full = sorted([b for b in blocks if b.column == 0 and b.y < y_threshold], key=lambda b: b.y)
    
    # 본문 컬럼
    left = sorted([b for b in blocks if b.column == 1], key=lambda b: b.y)
    right = sorted([b for b in blocks if b.column == 2], key=lambda b: b.y)
    
    # 하단 전체폭
    bottom_full = sorted([b for b in blocks if b.column == 0 and b.y >= y_threshold], key=lambda b: b.y)
    
    return top_full + left + right + bottom_full



def _sort_reading_order_interleaved(blocks: List[TextBlock], page_height: float) -> List[TextBlock]:
    """
    인터리빙 읽기 순서:
    전체폭 요소는 Y 위치에 따라 왼쪽/오른쪽 컬럼 사이에 배치
    """
    if not blocks:
        return []
    
    full = [b for b in blocks if b.column == 0]
    left = [b for b in blocks if b.column == 1]
    right = [b for b in blocks if b.column == 2]
    
    # 모든 블록을 Y로 정렬
    all_sorted = sorted(blocks, key=lambda b: b.y)
    
    result = []
    processed = set()
    
    for b in all_sorted:
        if id(b) in processed:
            continue
        
        if b.column == 0:
            # 전체폭 블록
            result.append(b)
            processed.add(id(b))
        else:
            # 컬럼 블록: 같은 Y 영역의 왼쪽 → 오른쪽 순서로
            y_threshold = b.font_size * 3
            
            # 같은 Y 영역의 왼쪽 컬럼 블록들
            same_y_left = [lb for lb in left if id(lb) not in processed 
                          and abs(lb.y - b.y) < y_threshold]
            same_y_left.sort(key=lambda x: x.y)
            
            # 같은 Y 영역의 오른쪽 컬럼 블록들  
            same_y_right = [rb for rb in right if id(rb) not in processed
                           and abs(rb.y - b.y) < y_threshold]
            same_y_right.sort(key=lambda x: x.y)
            
            # 왼쪽 먼저, 오른쪽 나중
            for lb in same_y_left:
                result.append(lb)
                processed.add(id(lb))
            for rb in same_y_right:
                result.append(rb)
                processed.add(id(rb))
    
    return result


def _y_cut_groups(items: List['BoxedItem']) -> List[List['BoxedItem']]:
    """Y축 갭으로 아이템 그룹 분할"""
    if not items:
        return []
    if len(items) < 2:
        return [items]
    
    sorted_items = sorted(items, key=lambda i: i.y0)
    
    # Y 구간 병합
    intervals = []
    curr_start, curr_end = sorted_items[0].y0, sorted_items[0].y1
    
    for item in sorted_items[1:]:
        if item.y0 <= curr_end + 4:
            curr_end = max(curr_end, item.y1)
        else:
            intervals.append((curr_start, curr_end))
            curr_start, curr_end = item.y0, item.y1
    intervals.append((curr_start, curr_end))
    
    # 각 구간의 아이템 수집
    groups = []
    for start, end in intervals:
        group = [item for item in items if item.y0 >= start - 2 and item.y0 <= end + 2]
        if group:
            groups.append(group)
    
    return groups


def analyze_page_layout(doc, page_num: int = 0) -> PageLayout:
    from .. import extract_text_with_positions, get_pages
    pages = get_pages(doc)
    if page_num >= len(pages):
        return PageLayout(0, 0)
    
    page = pages[page_num]
    media_box = page.get('MediaBox', [0, 0, 595, 842])
    try:
        w = float(media_box[2]) - float(media_box[0])
        h = float(media_box[3]) - float(media_box[1])
    except:
        w, h = 595.0, 842.0
    
    text_items = extract_text_with_positions(doc, page_num)
    return analyze_layout(text_items, w, h)


class BoxedItem:
    def __init__(self, item, x, y, w, h):
        self.item = item
        self.x0, self.y0 = x, y
        self.x1, self.y1 = x + w, y + h
        self.cx, self.cy = x + w/2, y + h/2


def _to_boxed_item(item) -> BoxedItem:
    fs = getattr(item, 'font_size', 10) or 10
    tlen = len(item.text) if item.text else 1
    # [수정] 한글/전각 문자 고려하여 너비 계수 0.4 → 0.55
    # 영문 반각: ~0.4, 한글 전각: ~0.5~0.6, 혼합 평균: ~0.55
    w = max(1.0, tlen * fs * 0.55)
    h = fs * 0.7
    return BoxedItem(item, item.x, item.y - h, w, h)


def _create_text_block(boxed_items: List[BoxedItem]) -> Optional[TextBlock]:
    if not boxed_items:
        return None
    
    boxed_items.sort(key=lambda b: (int(b.y0), b.x0))
    
    lines = []
    current_line = []
    last_y = boxed_items[0].y0
    
    for b in boxed_items:
        fs = getattr(b.item, 'font_size', 10) or 10
        if abs(b.y0 - last_y) > fs * 0.6:
            if current_line:
                lines.append(_merge_line(current_line))
            current_line = [b]
            last_y = b.y0
        else:
            current_line.append(b)
    if current_line:
        lines.append(_merge_line(current_line))
    
    text = _clean_text("\n".join(lines))
    if not text.strip():
        return None
    
    x0 = min(b.x0 for b in boxed_items)
    y0 = min(b.y0 for b in boxed_items)
    x1 = max(b.x1 for b in boxed_items)
    y1 = max(b.y1 for b in boxed_items)
    fs = sum(getattr(b.item, 'font_size', 10) or 10 for b in boxed_items) / len(boxed_items)
    
    return TextBlock(text=text, x=x0, y=y0, width=x1-x0, height=y1-y0, font_size=fs, items=[b.item for b in boxed_items])


def _merge_line(line: List[BoxedItem]) -> str:
    if not line:
        return ""
    line.sort(key=lambda b: b.x0)
    result = []
    prev_x = None
    for b in line:
        if prev_x is not None:
            gap = b.x0 - prev_x
            # [수정] gap threshold 조정
            # 기존: >15 → 3칸, >1 → 1칸 (너무 민감)
            # 수정: >20 → 공백, >5 → 공백, ≤5 → 연결 (너비 계수 보정과 함께)
            if gap > 20:
                result.append("  ")
            elif gap > 5:
                result.append(" ")
            # gap ≤ 5: 공백 없이 연결
        result.append(b.item.text or "")
        prev_x = b.x1
    return "".join(result)


def _clean_text(text: str) -> str:
    text = text.replace('\x00f\x00', 'fi').replace('\x00l\x00', 'fl').replace('\x00', '')
    return ''.join(c for c in text if c in '\n\t' or ord(c) >= 32)


def _merge_adjacent_blocks(blocks: List[TextBlock]) -> List[TextBlock]:
    if len(blocks) <= 1:
        return blocks
    
    merged = []
    curr = blocks[0]
    
    for nxt in blocks[1:]:
        gap = nxt.y - curr.bottom
        same_col = curr.column == nxt.column
        aligned = abs(nxt.x - curr.x) < 20
        same_font = abs(nxt.font_size - curr.font_size) < 4
        
        if same_col and aligned and same_font and -5 < gap < curr.font_size * 2:
            sep = "\n\n" if gap > curr.font_size * 1.2 else "\n"
            curr.text += sep + nxt.text
            curr.width = max(curr.right, nxt.right) - curr.x
            curr.height = nxt.bottom - curr.y
            curr.items.extend(nxt.items)
        else:
            merged.append(curr)
            curr = nxt
    merged.append(curr)
    return merged


def _classify_blocks(blocks: List[TextBlock], page_height: float) -> List[TextBlock]:
    if not blocks:
        return []
    avg_fs = sum(b.font_size for b in blocks) / len(blocks)
    
    for b in blocks:
        cy = b.y + b.height / 2
        if cy < page_height * 0.08:
            b.block_type = BlockType.HEADER
        elif cy > page_height * 0.92:
            b.block_type = BlockType.FOOTER
        elif b.font_size > avg_fs * 1.3:
            b.block_type = BlockType.TITLE if b.font_size > avg_fs * 2 else BlockType.HEADING
        elif b.text.lstrip().startswith(('•', '-', '1.', '(1)')):
            b.block_type = BlockType.LIST_ITEM
        else:
            b.block_type = BlockType.PARAGRAPH
    return blocks