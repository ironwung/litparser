"""
PDF Layout Analyzer - Stage 9.1 (Robust Column Detection)

개선:
1. [Stage 9] X좌표 범위 확대 및 갭 기반 2단 감지 개선
2. [Stage 9.1] x0 기반 아이템 분류 (추정 너비 x1 사용 안 함)
3. 표/그림 영역 분리 후 본문 컬럼 분석
4. 영역별 적응적 컬럼 감지
5. 주요 클러스터 기반 2단 감지
6. [Stage 8.1] 한글 텍스트 너비 계산 및 공백 삽입 개선

Stage 9.1 변경사항:
- 아이템 분류를 x0(실제 시작 좌표)만 사용하도록 변경
  → _to_boxed_item의 추정 너비(x1)가 부정확하여 잘못된 분류 발생 방지
- x0 < left_col_max → LEFT
- x0 >= right_col_min → RIGHT
- 나머지 (left_col_max <= x0 < right_col_min) → FULL_WIDTH
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
        
        # [Stage 9.1] x0 기반 아이템 분류
        # x0는 PDF에서 직접 추출한 정확한 시작 좌표
        # x1은 추정 너비로 계산되어 부정확하므로 사용하지 않음
        full_width = []
        left_col = []
        right_col = []
        
        for b in boxed_items:
            if b.x0 < left_max:
                left_col.append(b)
            elif b.x0 >= right_min:
                right_col.append(b)
            else:
                # left_max <= x0 < right_min: 갭 영역 (제목, 중앙 텍스트)
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
        
        all_blocks = _sort_reading_order_with_interleave(all_blocks)
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
    [Stage 9] 개선된 2단 본문 영역 찾기
    """
    result = {'num_columns': 1, 'left_col_max': page_width/2, 'right_col_min': page_width/2}
    
    if len(items) < 10:
        return result, 0
    
    y_values = sorted(set(int(i.y0 / 20) * 20 for i in items))
    
    def check_two_column(y_start, min_gap=150):
        region_items = [i for i in items if i.y0 >= y_start]
        if len(region_items) < 10:
            return None
        
        x_counter = Counter(int(i.x0 / 10) * 10 for i in region_items)
        major = sorted([(x, c) for x, c in x_counter.items() if c >= 3], key=lambda t: t[0])
        
        if len(major) < 2:
            return None
        
        center = page_width / 2
        best = None
        best_gap = 0
        
        for i in range(len(major) - 1):
            left_x, left_count = major[i]
            right_x, right_count = major[i + 1]
            gap = right_x - left_x
            gap_mid = (left_x + right_x) / 2
            
            if (gap >= min_gap and 
                page_width * 0.2 < gap_mid < page_width * 0.8 and
                gap > best_gap):
                
                left_count_total = sum(1 for i in region_items if i.x0 < gap_mid)
                right_count_total = sum(1 for i in region_items if i.x0 >= gap_mid)
                
                if left_count_total >= 5 and right_count_total >= 5:
                    best_gap = gap
                    best = {
                        'left_max': left_x + 60,
                        'right_min': right_x - 10,
                        'gap': gap,
                        'gap_center': gap_mid
                    }
        
        return best
    
    consecutive_count = 0
    first_y = None
    col_info = None
    
    for y_start in y_values:
        info = check_two_column(y_start)
        if info and info['gap'] >= 150:
            if first_y is None:
                first_y = y_start
                col_info = info
            consecutive_count += 1
            
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


def _sort_reading_order_with_interleave(blocks: List[TextBlock]) -> List[TextBlock]:
    """전체폭 블록의 Y 위치에 따라 인터리빙"""
    if not blocks:
        return []
    
    full = [b for b in blocks if b.column == 0]
    left = sorted([b for b in blocks if b.column == 1], key=lambda b: b.y)
    right = sorted([b for b in blocks if b.column == 2], key=lambda b: b.y)
    
    full_sorted = sorted(full, key=lambda b: b.y)
    
    if not full_sorted:
        return left + right
    
    result = []
    left_idx = 0
    right_idx = 0
    
    for fi, fb in enumerate(full_sorted):
        while left_idx < len(left) and left[left_idx].y < fb.y:
            result.append(left[left_idx])
            left_idx += 1
        while right_idx < len(right) and right[right_idx].y < fb.y:
            result.append(right[right_idx])
            right_idx += 1
        result.append(fb)
    
    while left_idx < len(left):
        result.append(left[left_idx])
        left_idx += 1
    while right_idx < len(right):
        result.append(right[right_idx])
        right_idx += 1
    
    return result


def _sort_reading_order_with_table(blocks: List[TextBlock], y_threshold: float) -> List[TextBlock]:
    """표/그림 영역 고려한 읽기 순서 (하위 호환)"""
    top_full = sorted([b for b in blocks if b.column == 0 and b.y < y_threshold], key=lambda b: b.y)
    left = sorted([b for b in blocks if b.column == 1], key=lambda b: b.y)
    right = sorted([b for b in blocks if b.column == 2], key=lambda b: b.y)
    bottom_full = sorted([b for b in blocks if b.column == 0 and b.y >= y_threshold], key=lambda b: b.y)
    return top_full + left + right + bottom_full


def _y_cut_groups(items: List['BoxedItem']) -> List[List['BoxedItem']]:
    """Y축 갭으로 아이템 그룹 분할"""
    if not items:
        return []
    if len(items) < 2:
        return [items]
    
    sorted_items = sorted(items, key=lambda i: i.y0)
    
    intervals = []
    curr_start, curr_end = sorted_items[0].y0, sorted_items[0].y1
    
    for item in sorted_items[1:]:
        if item.y0 <= curr_end + 4:
            curr_end = max(curr_end, item.y1)
        else:
            intervals.append((curr_start, curr_end))
            curr_start, curr_end = item.y0, item.y1
    intervals.append((curr_start, curr_end))
    
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
    prev_text = ""
    for b in line:
        text = b.item.text or ""
        if prev_x is not None:
            gap = b.x0 - prev_x
            prev_ends_space = prev_text.endswith(' ')
            curr_starts_space = text.startswith(' ')
            
            if prev_ends_space or curr_starts_space:
                pass
            elif gap > 15:
                result.append(" ")
            elif gap > 5:
                result.append(" ")
        result.append(text)
        prev_x = b.x1
        prev_text = text
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