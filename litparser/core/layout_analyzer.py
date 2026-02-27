"""
PDF Layout Analyzer - Stage 9.2 (Improved Column Detection)

개선:
1. [Stage 9] X좌표 범위 확대 및 갭 기반 2단 감지 개선
2. [Stage 9.1] x0 기반 아이템 분류 (추정 너비 x1 사용 안 함)
3. [Stage 9.2] 히스토그램 빈 영역(empty zone) 기반 2단 감지
   → 인접 클러스터 비교 대신 x0 분포의 빈 구간을 찾아 2단 경계 결정
   → 읽기 순서: 같은 Y 대역에서 left 전체 → right 전체
4. 표/그림 영역 분리 후 본문 컬럼 분석
5. 영역별 적응적 컬럼 감지
6. [Stage 8.1] 한글 텍스트 너비 계산 및 공백 삽입 개선

Stage 9.2 변경사항:
- _find_two_column_region: 히스토그램 빈 영역 방식으로 완전 재작성
  → x0 좌표를 5pt 단위 히스토그램으로 만들고, 페이지 중앙 부근의
     빈 구간(아이템이 거의 없는 영역)을 찾아 2단 경계로 사용
- _sort_reading_order_with_interleave: Y 대역별 left→right 순서 보장
  → 같은 Y 대역의 left 블록들을 모두 출력한 후 right 블록 출력
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
        full_width = []
        left_col = []
        right_col = []
        
        for b in boxed_items:
            if b.x0 < left_max:
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
    [Stage 9.2] 히스토그램 빈 영역 기반 2단 감지
    
    방법:
    1. 본문 영역 아이템의 x0 좌표를 5pt 단위 히스토그램으로 만듦
    2. 페이지 중앙 부근(20%~80%)에서 아이템이 거의 없는 빈 구간(gap)을 찾음
    3. 가장 넓은 빈 구간의 좌우를 2단 경계로 사용
    
    이전 방식(인접 클러스터 비교)의 문제:
    - 같은 단 내의 들여쓰기/수식 등이 별도 클러스터로 잡혀
      실제 2단 갭을 감지하지 못하는 경우가 발생
    """
    result = {'num_columns': 1, 'left_col_max': page_width / 2, 'right_col_min': page_width / 2}
    
    if len(items) < 10:
        return result, 0
    
    # 본문 영역 아이템만 사용 (상단/하단 여백 제외)
    body_items = [i for i in items if page_height * 0.08 < i.y0 < page_height * 0.92]
    if len(body_items) < 10:
        return result, 0
    
    # x0 히스토그램 (5pt 단위 bin)
    bin_size = 5
    num_bins = int(page_width / bin_size) + 1
    histogram = [0] * num_bins
    
    for item in body_items:
        bin_idx = int(item.x0 / bin_size)
        if 0 <= bin_idx < num_bins:
            histogram[bin_idx] += 1
    
    # x0 분포의 두 주요 클러스터 사이 빈 영역 찾기
    # 검색 범위: 마진(10%) 안쪽 전체
    search_start = int(page_width * 0.1 / bin_size)
    search_end = int(page_width * 0.9 / bin_size)
    
    # 빈 구간 = 히스토그램 값이 threshold 이하인 연속 구간
    # threshold를 0으로 시작하고, 못 찾으면 점진적으로 올림
    best_gaps = []
    
    for threshold in [0, 1, 2, 3]:
        gaps = []
        gap_start = None
        
        for i in range(search_start, search_end + 1):
            if histogram[i] <= threshold:
                if gap_start is None:
                    gap_start = i
            else:
                if gap_start is not None:
                    gap_width = (i - gap_start) * bin_size
                    gap_center = ((gap_start + i) / 2) * bin_size
                    # 갭이 페이지 중앙 부근(25%~75%)에 있어야 함
                    if (gap_width >= 15 and 
                        page_width * 0.25 < gap_center < page_width * 0.75):
                        gaps.append((gap_start, i, gap_width))
                    gap_start = None
        
        if gap_start is not None:
            gap_width = (search_end - gap_start) * bin_size
            gap_center = ((gap_start + search_end) / 2) * bin_size
            if (gap_width >= 15 and 
                page_width * 0.25 < gap_center < page_width * 0.75):
                gaps.append((gap_start, search_end, gap_width))
        
        if gaps:
            best_gaps = gaps
            break  # 가장 엄격한 threshold에서 찾은 갭 사용
    
    if not best_gaps:
        return result, 0
    
    # 갭 후보 중 좌우 아이템 균형이 맞는 것 필터링 후 가장 넓은 것 선택
    min_ratio = 0.15
    total = len(body_items)
    
    valid_gaps = []
    for gap in best_gaps:
        g_start_x = gap[0] * bin_size
        g_end_x = gap[1] * bin_size
        g_center = (g_start_x + g_end_x) / 2
        
        left_count = sum(1 for i in body_items if i.x0 < g_center)
        right_count = sum(1 for i in body_items if i.x0 >= g_center)
        
        if left_count >= total * min_ratio and right_count >= total * min_ratio:
            valid_gaps.append(gap)
    
    if not valid_gaps:
        return result, 0
    
    best_gap = max(valid_gaps, key=lambda g: g[2])
    gap_start_x = best_gap[0] * bin_size
    gap_end_x = best_gap[1] * bin_size
    gap_center = (gap_start_x + gap_end_x) / 2
    
    # ─── 추가 검증: 라인별 교차 검증 (line-crossing validation) ───
    # 글자 단위 아이템 감지: 한글 PDF는 글자별로 아이템이 분리되어
    # x0가 페이지 전체에 분산됨 → 가짜 gap 발생 가능
    avg_text_len = sum(len(getattr(i.item, 'text', '') or '') for i in body_items) / max(len(body_items), 1)
    short_items = sum(1 for i in body_items if len(getattr(i.item, 'text', '') or '') <= 2)
    is_char_level = avg_text_len < 5 and short_items > len(body_items) * 0.4
    
    from collections import defaultdict
    y_groups = defaultdict(list)
    for item in body_items:
        y_key = round(item.y0 / 5) * 5
        y_groups[y_key].append(item)
    
    crossing_lines = 0
    total_lines = 0
    for y_key, group in y_groups.items():
        if len(group) < 2:
            continue
        total_lines += 1
        left_in_gap = [i for i in group if i.x0 < gap_center - 10]
        right_in_gap = [i for i in group if i.x0 > gap_center + 10]
        if left_in_gap and right_in_gap:
            if is_char_level:
                # 글자 단위: 같은 줄에 양쪽 아이템 존재 = crossing
                crossing_lines += 1
            else:
                # 단어/구 단위: inner_gap이 작을 때만 crossing (연속 텍스트)
                max_left_x = max(i.x0 for i in left_in_gap)
                min_right_x = min(i.x0 for i in right_in_gap)
                if min_right_x - max_left_x < 50:
                    crossing_lines += 1
    
    # 50% 이상의 줄이 crossing이면 → 1단 문서
    if total_lines > 5 and crossing_lines > total_lines * 0.5:
        return result, 0
    
    # 2단으로 판정
    # left_col_max: 갭 시작점 (왼쪽 단의 아이템 x0은 이보다 작음)
    # right_col_min: 갭 끝점 (오른쪽 단의 아이템 x0은 이보다 크거나 같음)
    result['num_columns'] = 2
    result['left_col_max'] = gap_start_x
    result['right_col_min'] = gap_end_x
    
    # y_threshold: 2단이 시작되는 Y 좌표 (첫 번째 right 아이템 기준)
    right_items = [i for i in body_items if i.x0 >= gap_end_x]
    y_threshold = min(i.y0 for i in right_items) if right_items else 0
    
    return result, y_threshold


def _sort_reading_order_with_interleave(blocks: List[TextBlock]) -> List[TextBlock]:
    """
    [Stage 9.2] Y 대역별 left→right 읽기 순서
    
    2단 레이아웃에서 올바른 읽기 순서:
    1. full-width 블록을 Y 기준 분리점으로 사용
    2. 각 분리 구간에서: left 블록 전체(Y순) → right 블록 전체(Y순)
    3. full-width 블록이 없으면: left 전체 → right 전체
    """
    if not blocks:
        return []
    
    full = sorted([b for b in blocks if b.column == 0], key=lambda b: b.y)
    left = sorted([b for b in blocks if b.column == 1], key=lambda b: b.y)
    right = sorted([b for b in blocks if b.column == 2], key=lambda b: b.y)
    
    if not full:
        # full-width 블록이 없으면 left 전체 → right 전체
        return left + right
    
    # full-width 블록의 Y 위치를 분리점으로 사용
    result = []
    
    # 각 full-width 블록 사이의 구간에서 left→right 순서로 배치
    boundaries = [float('-inf')] + [fb.y for fb in full] + [float('inf')]
    
    for i in range(len(boundaries) - 1):
        y_min = boundaries[i]
        y_max = boundaries[i + 1]
        
        # 이 구간이 full-width 블록의 Y 위치인 경우
        if y_min != float('-inf') and i - 1 < len(full):
            fb = full[i - 1]
            # full-width 블록의 bottom을 기준으로 구간 시작
            y_min = fb.bottom
        
        if i > 0 and i - 1 < len(full):
            result.append(full[i - 1])
        
        # 이 Y 구간의 left/right 블록
        section_left = [b for b in left if b.y >= y_min and b.y < y_max]
        section_right = [b for b in right if b.y >= y_min and b.y < y_max]
        
        # left 전체 → right 전체
        result.extend(section_left)
        result.extend(section_right)
    
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