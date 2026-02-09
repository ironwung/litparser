"""
PDF Layout Analyzer - Stage 10 (Improved Column Detection)

개선:
1. X 시작좌표 히스토그램 기반 컬럼 감지
2. 중앙 정렬 제목 영역 감지
3. 올바른 읽기 순서
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
    font_name: str = ""
    items: List[Any] = field(default_factory=list)
    column: int = 0
    
    @property
    def bottom(self) -> float:
        return self.y + self.height
    
    @property
    def right(self) -> float:
        return self.x + self.width


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
    col_info = _detect_columns_improved(boxed_items, page_width)
    num_columns = col_info['num_columns']
    
    if num_columns >= 2:
        left_max = col_info['left_col_max']
        right_min = col_info['right_col_min']
        
        full_width_items = []
        left_items = []
        right_items = []
        
        for b in boxed_items:
            # 왼쪽 컬럼: X < left_max + 여유
            if b.x0 < left_max + 20:
                left_items.append(b)
            # 오른쪽 컬럼: X > right_min - 여유
            elif b.x0 > right_min - 20:
                right_items.append(b)
            # 중앙 (제목 등)
            else:
                full_width_items.append(b)
        
        all_blocks = []
        
        # 전체폭 (제목)
        if full_width_items:
            for group in _recursive_xy_cut(full_width_items, (0, 0, page_width, page_height)):
                block = _create_text_block(group)
                if block:
                    block.column = 0
                    all_blocks.append(block)
        
        # 왼쪽 컬럼
        if left_items:
            for group in _recursive_xy_cut(left_items, (0, 0, right_min, page_height)):
                block = _create_text_block(group)
                if block:
                    block.column = 1
                    all_blocks.append(block)
        
        # 오른쪽 컬럼
        if right_items:
            for group in _recursive_xy_cut(right_items, (right_min, 0, page_width, page_height)):
                block = _create_text_block(group)
                if block:
                    block.column = 2
                    all_blocks.append(block)
        
        # 정렬: 제목(Y순) → 왼쪽(Y순) → 오른쪽(Y순)
        all_blocks = _sort_reading_order(all_blocks)
    else:
        block_groups = _recursive_xy_cut(boxed_items, (0, 0, page_width, page_height))
        all_blocks = [_create_text_block(g) for g in block_groups]
        all_blocks = [b for b in all_blocks if b]
        all_blocks.sort(key=lambda b: (b.y, b.x))
    
    merged = _merge_adjacent_blocks(all_blocks)
    final = _classify_blocks(merged, page_height)
    
    return PageLayout(width=page_width, height=page_height, blocks=final, num_columns=num_columns)


def _detect_columns_improved(items: List['BoxedItem'], page_width: float) -> dict:
    """
    X 시작좌표의 두 주요 클러스터를 찾아 컬럼 감지
    """
    result = {'num_columns': 1, 'left_col_max': page_width / 2, 'right_col_min': page_width / 2}
    
    if len(items) < 10:
        return result
    
    # X 시작좌표 수집
    x_starts = [item.x0 for item in items]
    
    # 주요 클러스터 찾기 (10pt 단위)
    x_rounded = [int(x / 10) * 10 for x in x_starts]
    x_counter = Counter(x_rounded)
    
    # 빈도 >= 10인 주요 클러스터만 (제목 등 소수 아이템 제외)
    major_clusters = sorted([(x, count) for x, count in x_counter.items() if count >= 10], key=lambda t: t[0])
    
    if len(major_clusters) < 2:
        # 폴백: 빈도 >= 3인 클러스터
        freq_x = sorted([x for x, count in x_counter.items() if count >= 3])
        if len(freq_x) < 2:
            return result
        major_clusters = [(x, x_counter[x]) for x in freq_x]
    
    # 주요 클러스터 간 갭 찾기 (최소 150pt)
    best_gap = 0
    best_idx = -1
    
    for i in range(len(major_clusters) - 1):
        x1, c1 = major_clusters[i]
        x2, c2 = major_clusters[i + 1]
        gap = x2 - x1
        if gap > best_gap and gap >= 150:
            best_gap = gap
            best_idx = i
    
    if best_idx < 0:
        return result
    
    # 왼쪽/오른쪽 컬럼 경계
    left_x = major_clusters[best_idx][0]
    right_x = major_clusters[best_idx + 1][0]
    
    left_col_max = left_x + 50  # 왼쪽 컬럼 최대 X
    right_col_min = right_x - 10  # 오른쪽 컬럼 최소 X
    
    # 양쪽 컬럼에 충분한 아이템이 있는지 확인
    left_count = sum(1 for x in x_starts if x < left_col_max)
    right_count = sum(1 for x in x_starts if x >= right_col_min)
    
    if left_count >= 5 and right_count >= 5:
        result['num_columns'] = 2
        result['left_col_max'] = left_col_max
        result['right_col_min'] = right_col_min
    
    return result


def _sort_reading_order(blocks: List[TextBlock]) -> List[TextBlock]:
    """정렬: 제목 → 왼쪽 컬럼 → 오른쪽 컬럼 (각각 Y순)"""
    full = sorted([b for b in blocks if b.column == 0], key=lambda b: b.y)
    left = sorted([b for b in blocks if b.column == 1], key=lambda b: b.y)
    right = sorted([b for b in blocks if b.column == 2], key=lambda b: b.y)
    return full + left + right


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
    w = max(1.0, tlen * fs * 0.5)
    h = fs * 0.8
    return BoxedItem(item, item.x, item.y - h, w, h)


def _recursive_xy_cut(items, bounds, depth=0):
    if not items or len(items) <= 1 or depth > 20:
        return [items] if items else []
    
    y_split = _find_y_split(items)
    if y_split:
        upper = [i for i in items if i.cy < y_split]
        lower = [i for i in items if i.cy >= y_split]
        if upper and lower:
            return _recursive_xy_cut(upper, bounds, depth+1) + _recursive_xy_cut(lower, bounds, depth+1)
    return [items]


def _find_y_split(items):
    if len(items) < 2:
        return None
    sorted_items = sorted(items, key=lambda i: i.y0)
    best = None
    best_gap = 6.0
    for i in range(len(sorted_items) - 1):
        gap = sorted_items[i+1].y0 - sorted_items[i].y1
        if gap > best_gap:
            best_gap = gap
            best = (sorted_items[i].y1 + sorted_items[i+1].y0) / 2
    return best


def _create_text_block(boxed_items):
    if not boxed_items:
        return None
    boxed_items.sort(key=lambda b: (int(b.y0/5)*5, b.x0))
    
    lines, curr_line, last_y = [], [], boxed_items[0].y0
    for b in boxed_items:
        if abs(b.y0 - last_y) > 5:
            if curr_line:
                lines.append(_merge_line(curr_line))
            curr_line = [b]
            last_y = b.y0
        else:
            curr_line.append(b)
    if curr_line:
        lines.append(_merge_line(curr_line))
    
    text = _clean_text("\n".join(lines))
    if not text.strip():
        return None
    
    x0 = min(b.x0 for b in boxed_items)
    y0 = min(b.y0 for b in boxed_items)
    x1 = max(b.x1 for b in boxed_items)
    y1 = max(b.y1 for b in boxed_items)
    fs = sum(getattr(b.item, 'font_size', 10) or 10 for b in boxed_items) / len(boxed_items)
    
    return TextBlock(text=text, x=x0, y=y0, width=x1-x0, height=y1-y0, font_size=fs, items=[b.item for b in boxed_items])


def _clean_text(text):
    text = text.replace('\x00f\x00', 'fi').replace('\x00l\x00', 'fl').replace('\x00', '')
    return ''.join(c for c in text if c in '\n\t' or ord(c) >= 32)


def _merge_line(line):
    if not line:
        return ""
    line.sort(key=lambda b: b.x0)
    result = []
    prev = None
    for b in line:
        t = b.item.text or ""
        if prev is not None:
            gap = b.x0 - prev
            if gap > 20: result.append("   ")
            elif gap > 3: result.append(" ")
        result.append(t)
        prev = b.x1
    return "".join(result)


def _merge_adjacent_blocks(blocks):
    if len(blocks) <= 1:
        return blocks
    merged = []
    curr = blocks[0]
    for nxt in blocks[1:]:
        if (curr.column == nxt.column and abs(nxt.x - curr.x) < 20 and 
            abs(nxt.font_size - curr.font_size) < 3 and -5 < nxt.y - curr.bottom < curr.font_size * 2):
            sep = "\n" if nxt.y - curr.bottom < curr.font_size else "\n\n"
            curr.text += sep + nxt.text
            curr.width = max(curr.right, nxt.right) - curr.x
            curr.height = nxt.bottom - curr.y
            curr.items.extend(nxt.items)
        else:
            merged.append(curr)
            curr = nxt
    merged.append(curr)
    return merged


def _classify_blocks(blocks, page_height):
    if not blocks:
        return []
    avg_fs = sum(b.font_size for b in blocks) / len(blocks)
    for b in blocks:
        cy = b.y + b.height / 2
        if cy < page_height * 0.08:
            b.block_type = BlockType.HEADER
        elif cy > page_height * 0.92:
            b.block_type = BlockType.FOOTER
        elif b.font_size > avg_fs * 1.5:
            b.block_type = BlockType.TITLE
        elif b.font_size > avg_fs * 1.2:
            b.block_type = BlockType.HEADING
        elif b.text.lstrip().startswith(('•', '-', '·', '*', '1.', '2.')):
            b.block_type = BlockType.LIST_ITEM
        else:
            b.block_type = BlockType.PARAGRAPH
    return blocks
