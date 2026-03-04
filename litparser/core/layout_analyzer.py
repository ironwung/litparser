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
    
    # 1. 2단 본문 영역 찾기 (라인 시작점 기반)
    col_info, y_threshold = _find_two_column_region(boxed_items, page_width, page_height)
    num_columns = col_info['num_columns']
    
    if num_columns >= 2:
        gap_left = col_info['gap_left']
        gap_right = col_info['gap_right']
        
        # 라인 단위로 좌/우 분류
        from collections import defaultdict, Counter
        y_groups = defaultdict(list)
        for b in boxed_items:
            y_key = round(b.y0 / 3) * 3
            y_groups[y_key].append(b)
        
        # 우 컬럼 주 시작점 계산 (gap 오른쪽의 최빈 라인 시작 x0)
        right_line_starts = []
        for group in y_groups.values():
            sx = min(b.x0 for b in group)
            if sx >= gap_right:
                right_line_starts.append(sx)
        if right_line_starts:
            right_start_bins = Counter(round(x / 10) * 10 for x in right_line_starts)
            main_right_start = right_start_bins.most_common(1)[0][0]
        else:
            main_right_start = gap_right
        # 우 컬럼 판정 임계값: 우 주 시작점에서 gap 폭의 30% 이내
        right_threshold = main_right_start - (gap_right - gap_left) * 0.3
        
        left_col = []
        right_col = []
        full_width = []
        
        for y_key, group in y_groups.items():
            line_start_x = min(b.x0 for b in group)
            line_end_x = max(b.x0 for b in group)
            
            if line_start_x < gap_left and line_end_x > gap_right:
                # 양쪽에 걸침 → 갭 영역에서 분리점 찾기
                items_sorted = sorted(group, key=lambda b: b.x0)
                best_split = -1
                best_gap_size = 0
                for i in range(len(items_sorted) - 1):
                    curr_x = items_sorted[i].x0
                    next_x = items_sorted[i + 1].x0
                    inner_gap = next_x - curr_x
                    if (inner_gap > best_gap_size and 
                        curr_x < gap_right and next_x > gap_left and
                        inner_gap >= (gap_right - gap_left) * 0.3):
                        best_split = i
                        best_gap_size = inner_gap
                
                if best_split >= 0:
                    left_col.extend(items_sorted[:best_split + 1])
                    right_col.extend(items_sorted[best_split + 1:])
                else:
                    full_width.extend(group)
            elif line_start_x < gap_left:
                # 좌 컬럼 영역에서 시작
                left_col.extend(group)
            elif line_start_x >= right_threshold:
                # 우 컬럼 주 시작점 근처에서 시작
                right_col.extend(group)
            else:
                # gap 내부에서 시작하지만 우 컬럼 주 시작점과 먼 경우
                # → 아이템 간 큰 갭이 있으면 분리, 없으면 full-width
                if len(group) >= 2:
                    items_sorted = sorted(group, key=lambda b: b.x0)
                    best_split = -1
                    best_gap_size = 0
                    for i in range(len(items_sorted) - 1):
                        curr_x = items_sorted[i].x0
                        next_x = items_sorted[i + 1].x0
                        inner_gap = next_x - curr_x
                        if (inner_gap > best_gap_size and
                            inner_gap >= (gap_right - gap_left) * 0.3):
                            best_split = i
                            best_gap_size = inner_gap
                    if best_split >= 0:
                        left_col.extend(items_sorted[:best_split + 1])
                        right_col.extend(items_sorted[best_split + 1:])
                    else:
                        full_width.extend(group)
                else:
                    full_width.extend(group)
        
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
    [Stage 10] 라인 시작점 기반 2단 감지
    
    방법:
    1. 같은 Y의 아이템을 라인으로 묶음
    2. 각 라인의 시작 x0(최소 x)을 수집
    3. 라인 시작 x0 히스토그램에서 두 클러스터 사이 갭을 찾음
    
    이전 방식(개별 아이템 x0 히스토그램)의 문제:
    - 좌 컬럼의 수식/들여쓰기 아이템이 페이지 중앙까지 퍼져서
      실제 갭을 찾지 못하거나 좁은 갭만 찾음
    - 이로 인해 좌 컬럼 아이템이 우측으로 잘못 분류
    """
    result = {'num_columns': 1, 'gap_left': page_width / 2, 'gap_right': page_width / 2}
    
    if len(items) < 10:
        return result, 0
    
    # 본문 영역 아이템만 사용 (상단/하단 여백 제외)
    body_items = [i for i in items if page_height * 0.08 < i.y0 < page_height * 0.92]
    if len(body_items) < 10:
        return result, 0
    
    # 라인 구성: 같은 Y(3pt tolerance)의 아이템을 하나의 라인으로
    from collections import defaultdict
    y_groups = defaultdict(list)
    for b in body_items:
        y_key = round(b.y0 / 3) * 3
        y_groups[y_key].append(b)
    
    # 각 라인의 시작 x0 수집
    line_starts = []
    for y_key, group in y_groups.items():
        if len(group) >= 1:
            first_x = min(b.x0 for b in group)
            line_starts.append(first_x)
    
    if len(line_starts) < 6:
        return result, 0
    
    # 라인 시작 x0 히스토그램 (10pt 단위 bin)
    bin_size = 10
    num_bins = int(page_width / bin_size) + 1
    histogram = [0] * num_bins
    
    for x in line_starts:
        bin_idx = int(x / bin_size)
        if 0 <= bin_idx < num_bins:
            histogram[bin_idx] += 1
    
    # 페이지 중앙 부근(15%~85%)에서 빈 구간(gap) 찾기
    # 노이즈 처리: bin 값이 주요 클러스터 최대값의 5% 미만이면 빈 것으로 간주
    max_bin = max(histogram) if histogram else 1
    noise_threshold = max(1, int(max_bin * 0.05))
    
    search_start = int(page_width * 0.15 / bin_size)
    search_end = int(page_width * 0.85 / bin_size)
    
    # 빈 구간 = 히스토그램 값이 noise_threshold 이하인 연속 구간
    gaps = []
    gap_start = None
    
    for i in range(search_start, search_end + 1):
        if histogram[i] <= noise_threshold:
            if gap_start is None:
                gap_start = i
        else:
            if gap_start is not None:
                gap_width = (i - gap_start) * bin_size
                if gap_width >= 30:  # 최소 30pt 갭
                    gaps.append((gap_start, i, gap_width))
                gap_start = None
    
    if gap_start is not None:
        i = search_end
        gap_width = (i - gap_start) * bin_size
        if gap_width >= 30:
            gaps.append((gap_start, i, gap_width))
    
    # 追加: 個別アイテムx0ヒストグラムからもギャップを収集
    # 同じY行に左右アイテムが並ぶ場合(日本語文書等)、
    # ラインスタートでは正確なギャップが見つからない
    item_histogram = [0] * num_bins
    for item in body_items:
        bi = int(item.x0 / bin_size)
        if 0 <= bi < num_bins:
            item_histogram[bi] += 1
    
    item_max_bin = max(item_histogram) if item_histogram else 1
    item_noise = max(1, int(item_max_bin * 0.05))
    
    item_gap_start = None
    for i in range(search_start, search_end + 1):
        if item_histogram[i] <= item_noise:
            if item_gap_start is None:
                item_gap_start = i
        else:
            if item_gap_start is not None:
                gw = (i - item_gap_start) * bin_size
                if gw >= 30:
                    gaps.append((item_gap_start, i, gw))
                item_gap_start = None
    if item_gap_start is not None:
        gw = (search_end - item_gap_start) * bin_size
        if gw >= 30:
            gaps.append((item_gap_start, search_end, gw))
    
    if not gaps:
        # 히스토그램 빈 구간이 없으면 클러스터 피크 기반으로 시도
        pass
    
    # === 추가: 클러스터 피크 기반 갭 감지 ===
    # 히스토그램에서 빈 구간이 아닌, 가장 밀도 높은 2개 피크 사이를 갭으로 설정
    # Figure 캡션/각주가 갭 영역을 오염시켜도 주력 피크는 정확함
    item_bins = Counter(round(b.x0 / bin_size) * bin_size for b in body_items)
    if item_bins:
        sorted_peaks = sorted(item_bins.items(), key=lambda x: -x[1])
        peak1_x = sorted_peaks[0][0]
        peak1_cnt = sorted_peaks[0][1]
        # 2번째 피크: 1번째와 최소 gap_min_width 이상 떨어진
        gap_min_width = page_width * 0.15
        for px, cnt in sorted_peaks[1:]:
            if abs(px - peak1_x) >= gap_min_width and cnt >= peak1_cnt * 0.15:
                left_peak = min(peak1_x, px)
                right_peak = max(peak1_x, px)
                # 두 피크 사이를 갭으로
                # 좌측 끝: left_peak + margin, 우측 시작: right_peak - margin
                margin = bin_size * 2
                cluster_gap_left = left_peak + margin
                cluster_gap_right = right_peak - margin
                if cluster_gap_right > cluster_gap_left:
                    gw = cluster_gap_right - cluster_gap_left
                    gaps.append((int(cluster_gap_left / bin_size), 
                                int(cluster_gap_right / bin_size), gw))
                break
    
    if not gaps:
        return result, 0
    
    # 갭 후보 중 좌우 균형이 맞는 것 선택
    # 라인 시작점과 아이템 x0 모두에서 검증
    min_ratio = 0.2
    total_items = len(body_items)
    
    valid_gaps = []
    for gap in gaps:
        g_left = gap[0] * bin_size
        g_right = gap[1] * bin_size
        gap_w = g_right - g_left
        
        # 갭 폭이 페이지 폭의 15% 미만이면 2단이 아님 (표/양식의 좁은 갭 제외)
        if gap_w < page_width * 0.15:
            continue
        
        # アイテムx0基準でバランスチェック（ラインスタートでは検出できないケース対応）
        left_items = sum(1 for b in body_items if b.x0 < g_left)
        right_items = sum(1 for b in body_items if b.x0 >= g_right)
        
        if (left_items >= total_items * min_ratio and 
            right_items >= total_items * min_ratio):
            valid_gaps.append(gap)
    
    if not valid_gaps:
        return result, 0
    
    # 가장 넓은 갭 선택
    best_gap = max(valid_gaps, key=lambda g: g[2])
    gap_left_x = best_gap[0] * bin_size
    gap_right_x = best_gap[1] * bin_size
    
    # 라인별 교차 검증: 같은 Y에서 좌측 라인과 우측 라인이 뒤섞이면 1단
    crossing_lines = 0
    total_checked = 0
    for y_key, group in y_groups.items():
        if len(group) < 2:
            continue
        total_checked += 1
        has_left = any(b.x0 < gap_left_x for b in group)
        has_right = any(b.x0 > gap_right_x for b in group)
        if has_left and has_right:
            # 좌/우 사이에 실제 갭이 있는지 확인
            sorted_x = sorted(b.x0 for b in group)
            max_internal_gap = max(sorted_x[i+1] - sorted_x[i] for i in range(len(sorted_x)-1))
            if max_internal_gap < (gap_right_x - gap_left_x) * 0.5:
                crossing_lines += 1
    
    if total_checked > 5 and crossing_lines > total_checked * 0.5:
        return result, 0
    
    # 2단 확정
    result['num_columns'] = 2
    result['gap_left'] = gap_left_x
    result['gap_right'] = gap_right_x
    
    # y_threshold: 우측 컬럼이 시작되는 Y
    right_starts = [min(b.y0 for b in group) 
                    for y_key, group in y_groups.items() 
                    if min(b.x0 for b in group) >= gap_right_x]
    y_threshold = min(right_starts) if right_starts else 0
    
    return result, y_threshold


def _sort_reading_order_with_interleave(blocks: List[TextBlock]) -> List[TextBlock]:
    """
    [Stage 10] Y 대역별 left→right 읽기 순서
    
    2단 레이아웃에서 올바른 읽기 순서:
    1. 의미있는 full-width 블록(캡션, 제목 등)을 Y 기준 분리점으로 사용
    2. 각 분리 구간에서: left 블록 전체(Y순) → right 블록 전체(Y순)
    3. 짧은 full-width 블록(페이지 번호, 헤더)은 가까운 컬럼에 편입
    """
    if not blocks:
        return []
    
    left = sorted([b for b in blocks if b.column == 1], key=lambda b: b.y)
    right = sorted([b for b in blocks if b.column == 2], key=lambda b: b.y)
    
    # full-width 블록을 의미있는 것과 사소한 것으로 분류
    # 의미있는 full-width: 텍스트가 50자 이상이거나 Figure/Table 캡션
    significant_full = []
    minor_full = []
    for b in blocks:
        if b.column != 0:
            continue
        text = b.text.strip()
        if not text:
            continue
        is_significant = (len(text) > 50 or 
                         text.lower().startswith('figure') or
                         text.lower().startswith('table'))
        if is_significant:
            significant_full.append(b)
        else:
            minor_full.append(b)
    
    significant_full.sort(key=lambda b: b.y)
    
    # 사소한 full-width 블록은 위치에 따라 left 또는 right에 편입
    for b in minor_full:
        # 가장 가까운 left/right 블록과 비교
        if left and (not right or abs(b.y - left[0].y) < abs(b.y - right[0].y)):
            b.column = 1
            left.append(b)
        elif right:
            b.column = 2
            right.append(b)
        else:
            significant_full.append(b)
    
    left.sort(key=lambda b: b.y)
    right.sort(key=lambda b: b.y)
    significant_full.sort(key=lambda b: b.y)
    
    if not significant_full:
        return left + right
    
    # significant full-width 블록의 Y 위치를 분리점으로 사용
    result = []
    boundaries = [float('-inf')] + [fb.y for fb in significant_full] + [float('inf')]
    
    for i in range(len(boundaries) - 1):
        y_min = boundaries[i]
        y_max = boundaries[i + 1]
        
        if y_min != float('-inf') and i - 1 < len(significant_full):
            fb = significant_full[i - 1]
            y_min = fb.bottom
        
        if i > 0 and i - 1 < len(significant_full):
            result.append(significant_full[i - 1])
        
        section_left = [b for b in left if b.y >= y_min and b.y < y_max]
        section_right = [b for b in right if b.y >= y_min and b.y < y_max]
        
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