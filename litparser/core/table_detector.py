"""
PDF Table Detector - Stage 5 (Visual Projection 방식)

텍스트 위치 기반 테이블 감지:
1. Y좌표로 행(Row) 분리
2. X좌표 클러스터링으로 열(Column) 감지
3. 쪼개진 테이블 병합 (Stitching)
4. 유효성 검증

Based on Gemini's Visual-Stitch approach with improvements.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Set
from collections import defaultdict
import re as regex


@dataclass
class TableCell:
    """테이블 셀"""
    row: int
    col: int
    text: str
    x: float
    y: float
    width: float = 0
    height: float = 0


@dataclass 
class Table:
    """감지된 테이블"""
    cells: List[TableCell] = field(default_factory=list)
    rows: int = 0
    cols: int = 0
    x: float = 0
    y: float = 0
    width: float = 0
    height: float = 0
    caption: str = ""
    method: str = ""
    page: int = 0
    
    def to_list(self) -> List[List[str]]:
        """2D 리스트로 변환"""
        if not self.cells:
            return []
        
        max_row = max((c.row for c in self.cells), default=0)
        max_col = max((c.col for c in self.cells), default=0)
        
        grid = [['' for _ in range(max_col + 1)] for _ in range(max_row + 1)]
        
        for cell in sorted(self.cells, key=lambda c: (c.row, c.col)):
            r = min(cell.row, max_row)
            c = min(cell.col, max_col)
            text = regex.sub(r'\s+', ' ', cell.text.strip())
            if grid[r][c]:
                grid[r][c] += ' ' + text
            else:
                grid[r][c] = text
        
        return grid
    
    def to_markdown(self) -> str:
        """마크다운 테이블로 변환"""
        grid = self.to_list()
        if not grid:
            return ""
        
        lines = []
        if self.caption:
            lines.append(f"**{self.caption}**\n")
        
        # 헤더
        lines.append('| ' + ' | '.join(grid[0]) + ' |')
        lines.append('| ' + ' | '.join(['---'] * len(grid[0])) + ' |')
        
        # 본문
        for row in grid[1:]:
            lines.append('| ' + ' | '.join(row) + ' |')
        
        return '\n'.join(lines)
    
    def to_csv(self) -> str:
        """CSV 형식으로 변환"""
        grid = self.to_list()
        lines = []
        for row in grid:
            escaped = []
            for cell in row:
                if ',' in cell or '"' in cell or '\n' in cell:
                    escaped.append('"' + cell.replace('"', '""') + '"')
                else:
                    escaped.append(cell)
            lines.append(','.join(escaped))
        return '\n'.join(lines)


def detect_tables(text_items: list, 
                  min_rows: int = 2,
                  min_cols: int = 2,
                  debug: bool = False,
                  **kwargs) -> List[Table]:
    """
    텍스트 위치 기반 테이블 감지
    
    Args:
        text_items: TextItem 리스트
        min_rows: 최소 행 수
        min_cols: 최소 열 수
        debug: 디버그 출력
    
    Returns:
        List[Table]: 감지된 테이블 목록
    """
    if not text_items:
        return []
    
    # 1. 텍스트를 블록으로 그룹화
    blocks = _group_text_into_blocks(text_items)
    
    if debug:
        print(f"[Debug] 텍스트 블록: {len(blocks)}개")
    
    # 2. 각 블록에서 테이블 구조 분석
    raw_tables = []
    for block in blocks:
        table = _analyze_block_for_table(block, min_cols)
        if table:
            raw_tables.append(table)
    
    if debug:
        print(f"[Debug] 초기 테이블 후보: {len(raw_tables)}개")
    
    # 3. 쪼개진 테이블 병합
    merged_tables = _stitch_tables(raw_tables, debug)
    
    # 4. 유효성 검증
    final_tables = []
    for t in merged_tables:
        if _is_valid_table(t, min_rows, min_cols):
            t.method = "visual-projection"
            final_tables.append(t)
    
    # Y좌표 내림차순 정렬 (위에서 아래)
    final_tables.sort(key=lambda t: t.y, reverse=True)
    
    return final_tables


def _group_text_into_blocks(text_items: list, gap_threshold: float = 25) -> List[List]:
    """
    텍스트를 수직 간격 기준으로 블록 분리
    
    큰 Y 간격이 있으면 별도 블록으로 분리
    """
    if not text_items:
        return []
    
    # Y 오름차순 정렬 (작은 Y = 위쪽, PDF 표준 좌표계)
    sorted_items = sorted(text_items, key=lambda t: t.y)
    
    blocks = []
    current_block = [sorted_items[0]]
    last_y = sorted_items[0].y
    
    for item in sorted_items[1:]:
        # 수직 간격이 크면 새 블록
        if item.y - last_y > gap_threshold:
            if current_block:
                blocks.append(current_block)
            current_block = []
        
        current_block.append(item)
        last_y = item.y
    
    if current_block:
        blocks.append(current_block)
    
    return blocks


def _analyze_block_for_table(block: list, min_cols: int = 2) -> Optional[Table]:
    """
    블록 내 텍스트 배치를 분석하여 테이블 구조 추출
    
    Visual Projection 방식:
    - Y좌표로 행 구분
    - X좌표 세그먼트로 열 구분
    """
    if len(block) < 3:  # 최소 3개 텍스트 필요
        return None
    
    # 1. 행(Row) 구분: Y좌표 그룹화 (5pt 단위)
    rows = defaultdict(list)
    for t in block:
        row_key = int(t.y / 5)
        rows[row_key].append(t)
    
    # Y좌표 방향 자동 감지: 
    # - 일반 PDF: Y가 작을수록 위쪽 (오름차순 정렬)
    # - 일부 PDF: Y가 클수록 위쪽 (내림차순 정렬)
    # 여기서는 항상 오름차순(작은 Y가 첫 행)으로 정렬
    sorted_row_keys = sorted(rows.keys())  # 오름차순 (작은 Y = 위쪽)
    
    if len(sorted_row_keys) < 2:
        return None
    
    # 연속 문장/불릿 문단 사전 필터링
    if len(sorted_row_keys) <= 4:
        # 각 행의 텍스트 합치기
        row_texts = {}
        row_starts = {}
        for rk in sorted_row_keys:
            items = sorted(rows[rk], key=lambda t: t.x)
            non_space = [t for t in items if t.text.strip()]
            if non_space:
                row_starts[rk] = min(t.x for t in non_space)
                row_texts[rk] = ' '.join(t.text for t in non_space).strip()
            else:
                row_texts[rk] = ''
        
        if row_texts:
            all_texts = list(row_texts.values())
            
            # 패턴 1: 불릿/넘버링 문단 (◦, ➀, ➁, □, -, · 등으로 시작)
            bullet_pattern = regex.compile(r'^[◦○●◆■□➀➁➂➃\-·•※▶►▷△▲]')
            first_text = all_texts[0] if all_texts else ''
            if bullet_pattern.match(first_text):
                # 첫 행이 불릿으로 시작하면 문단일 가능성 높음
                # 나머지 행도 같은 불릿이거나 이어지는 텍스트인지 확인
                is_paragraph = True
                for txt in all_texts[1:]:
                    if not txt:
                        continue
                    # 다음 행이 불릿으로 시작하거나, 이어지는 텍스트
                    if bullet_pattern.match(txt):
                        continue  # 같은 수준의 불릿
                    # 조사/접속사/일반 단어로 시작하면 이어지는 문장
                    if not regex.match(r'^[A-Z가-힣]', txt):
                        is_paragraph = False
                        break
                if is_paragraph:
                    return None
            
            # 패턴 2: 같은 X 시작점에서 시작하는 연속 문장
            if row_starts:
                starts = list(row_starts.values())
                if max(starts) - min(starts) < 15:
                    all_sentence = True
                    for txt in all_texts:
                        if not txt or len(txt) < 3:
                            continue
                        # '-'로 시작하는 목록 또는 마침표로 끝나는 문장
                        if txt.startswith('-') or txt.startswith('·'):
                            continue
                        if txt[-1] in '.。,，':
                            continue
                        all_sentence = False
                    if all_sentence:
                        return None
    
    # 2. 열(Column) 감지: X좌표 세그먼트 병합
    x_segments = []
    for t in block:
        # 텍스트 너비 추정
        char_width = t.font_size * 0.6
        width = max(10, len(t.text) * char_width)
        x_segments.append((t.x, t.x + width))
    
    x_segments.sort()
    
    # 겹치거나 가까운 세그먼트 병합
    columns = []
    if x_segments:
        curr_start, curr_end = x_segments[0]
        for seg_start, seg_end in x_segments[1:]:
            if seg_start < curr_end + 20:  # 20pt 이내면 병합
                curr_end = max(curr_end, seg_end)
            else:
                columns.append((curr_start, curr_end))
                curr_start, curr_end = seg_start, seg_end
        columns.append((curr_start, curr_end))
    
    if len(columns) < min_cols:
        return None
    
    # 3. 셀 생성
    cells = []
    for row_idx, row_key in enumerate(sorted_row_keys):
        row_items = rows[row_key]
        
        for t in row_items:
            # 텍스트 중심 X 좌표
            t_center = t.x + len(t.text) * t.font_size * 0.3
            
            # 가장 가까운 열 찾기
            best_col = -1
            best_dist = float('inf')
            for col_idx, (col_start, col_end) in enumerate(columns):
                # 열 범위 내에 있는지
                if col_start - 25 <= t.x <= col_end + 25:
                    dist = abs(t.x - col_start)
                    if dist < best_dist:
                        best_dist = dist
                        best_col = col_idx
            
            if best_col != -1:
                cells.append(TableCell(
                    row=row_idx,
                    col=best_col,
                    text=t.text,
                    x=t.x,
                    y=t.y,
                    width=len(t.text) * t.font_size * 0.6,
                    height=t.font_size
                ))
    
    # 4. 같은 셀의 텍스트 병합
    cell_groups = defaultdict(list)
    for c in cells:
        cell_groups[(c.row, c.col)].append(c)
    
    final_cells = []
    for (row, col), group in cell_groups.items():
        # X 순서대로 정렬 후 병합
        group.sort(key=lambda c: c.x)
        merged_text = ' '.join(c.text for c in group)
        final_cells.append(TableCell(
            row=row,
            col=col,
            text=merged_text,
            x=group[0].x,
            y=group[0].y
        ))
    
    # 바운딩 박스 계산
    if block:
        bx_min = min(t.x for t in block)
        bx_max = max(t.x + len(t.text) * t.font_size for t in block)
        by_min = min(t.y for t in block)
        by_max = max(t.y for t in block)
    else:
        bx_min = bx_max = by_min = by_max = 0
    
    return Table(
        cells=final_cells,
        rows=len(sorted_row_keys),
        cols=len(columns),
        x=bx_min,
        y=by_max,
        width=bx_max - bx_min,
        height=by_max - by_min
    )


def _stitch_tables(tables: List[Table], debug: bool = False) -> List[Table]:
    """
    수직으로 인접한 테이블 병합
    
    헤더와 본문이 분리된 경우, 연속된 테이블 등을 하나로 합침
    """
    if not tables:
        return []
    
    # Y 오름차순 정렬 (작은 Y = 위쪽)
    tables.sort(key=lambda t: t.y)
    
    merged = []
    current = tables[0]
    
    for next_t in tables[1:]:
        # 현재 테이블 하단과 다음 테이블 상단 간격
        current_bottom = current.y + current.height
        gap = next_t.y - current_bottom
        
        # 수평 겹침 계산
        x1 = max(current.x, next_t.x)
        x2 = min(current.x + current.width, next_t.x + next_t.width)
        overlap = max(0, x2 - x1)
        min_width = min(current.width, next_t.width) if min(current.width, next_t.width) > 0 else 1
        
        # 병합 조건:
        # 1. 수평 70% 이상 겹침
        # 2. 수직 간격 50pt 이내 (과도한 병합 방지)
        # 3. 열 개수가 같으면 간격 완화 (70pt)
        is_overlap = overlap > (min_width * 0.7)
        is_gap_ok = -30 < gap < 50
        is_same_cols = current.cols == next_t.cols
        is_gap_ok_relaxed = -30 < gap < 70
        
        should_merge = is_overlap and (is_gap_ok or (is_same_cols and is_gap_ok_relaxed))
        
        if should_merge:
            if debug:
                print(f"[Stitch] 병합: {current.rows}x{current.cols} + {next_t.rows}x{next_t.cols}, gap={gap:.1f}")
            
            # 행 번호 오프셋
            row_offset = current.rows
            for cell in next_t.cells:
                cell.row += row_offset
            
            current.cells.extend(next_t.cells)
            current.rows += next_t.rows
            current.cols = max(current.cols, next_t.cols)
            
            # 바운딩 박스 업데이트
            new_top = min(current.y, next_t.y)
            new_bottom = max(current.y + current.height, next_t.y + next_t.height)
            new_left = min(current.x, next_t.x)
            new_right = max(current.x + current.width, next_t.x + next_t.width)
            
            current.y = new_top
            current.height = new_bottom - new_top
            current.x = new_left
            current.width = new_right - new_left
        else:
            merged.append(current)
            current = next_t
    
    merged.append(current)
    return merged


def _is_valid_table(table: Table, min_rows: int, min_cols: int) -> bool:
    """테이블 유효성 검증"""
    # 1. 최소 크기
    if table.rows < min_rows or table.cols < min_cols:
        return False
    
    # 2. 셀 채움률 (10% 이상)
    non_empty = sum(1 for c in table.cells if c.text.strip())
    total = table.rows * table.cols
    if total > 6 and non_empty < total * 0.1:
        return False
    
    # 3. 과도한 텍스트 (셀당 500자 이하)
    max_len = max((len(c.text) for c in table.cells), default=0)
    if max_len > 500:
        return False
    
    # 4. 단일 열 필터링 (가짜 테이블)
    if table.cols == 1:
        return False
    
    # 5. 소규모 테이블(2-4행)에서 문장 분할 오탐 감지
    if table.rows <= 4:
        grid = table.to_list()
        sentence_rows = 0
        single_cell_rows = 0
        non_empty_rows = 0
        
        for row in grid:
            non_empty_cells = [cell.strip() for cell in row if cell.strip()]
            if not non_empty_cells:
                continue
            non_empty_rows += 1
            
            if len(non_empty_cells) < 2:
                # 셀이 1개인 행: 앞 행에서 줄바꿈된 나머지일 가능성 높음
                single_cell_rows += 1
                continue
            
            is_split = _is_split_sentence(non_empty_cells)
            if is_split:
                sentence_rows += 1
        
        # 분할 문장 행 + 단일셀 행이 대부분이면 표가 아님
        if non_empty_rows > 0:
            bad_rows = sentence_rows + single_cell_rows
            if bad_rows >= non_empty_rows * 0.6:
                return False
    
    # 6. 빈 셀 비율이 너무 높은 테이블 (70% 이상 빈 셀)
    if total > 0:
        empty_ratio = 1 - (non_empty / total)
        if empty_ratio > 0.7 and table.rows > 3:
            return False
    
    # 7. 소규모 테이블(2-3행)에서 모든 셀이 긴 서술형 텍스트이면 표가 아님
    #    진짜 표: 헤더가 짧거나 숫자/코드/키워드 셀이 있음
    #    가짜 표: 모든 셀이 긴 한국어 문장 조각
    if table.rows <= 3 and table.cols == 2:
        grid = table.to_list()
        all_prose = True
        for row in grid:
            for cell in row:
                cell_text = cell.strip()
                if not cell_text:
                    continue
                # 짧은 셀(5자 이하)이 있으면 → 헤더/라벨일 수 있음 → 진짜 표 가능
                if len(cell_text) <= 5:
                    all_prose = False
                    break
                # 숫자만 있으면 → 데이터 셀 → 진짜 표
                if cell_text.replace(' ', '').replace(',', '').replace('.', '').isdigit():
                    all_prose = False
                    break
            if not all_prose:
                break
        if all_prose:
            # 모든 셀이 6자 이상 서술형 → 분할 문장일 가능성 높음
            return False
    
    return True


def _is_split_sentence(cells: list) -> bool:
    """
    셀 리스트가 하나의 문장이 분할된 것인지 판단
    
    예: ['입찰공고에 명시된 일정에 따른 것이라면 , 반드시 모든 입찰업체가', '참여하지 않았']
    → True (하나의 문장이 잘린 것)
    
    예: ['국토교통부 고시 제2016-943호', '국토교통부 고시 제2018-614호']
    → False (두 개의 독립된 내용)
    """
    if len(cells) < 2:
        return False
    
    first = cells[0]
    last = cells[-1]
    
    if not first or not last:
        return False
    
    # 완전한 문장 끝 패턴
    sentence_end_chars = ('.', '!', '?', '。', ')', '」', ', ')
    # 한국어 문장 종결어미
    ko_endings = ('다.', '함.', '임.', '음.', '됨.', '있다.', '한다.', '된다.',
                  '없다.', '않다.', '이다.', '하다.')
    
    first_stripped = first.rstrip()
    last_stripped = last.rstrip()
    
    # --- 패턴 1: 앞 셀이 길고 문장 끝이 아님, 뒤 셀이 짧음 ---
    first_ends_mid = (
        len(first) > 15 and
        not any(first_stripped.endswith(e) for e in sentence_end_chars)
    )
    last_is_shorter = len(last) < len(first)
    last_not_new_sentence = (
        not last[0].isupper() and not last[0].isdigit()
    )
    
    if first_ends_mid and last_is_shorter and last_not_new_sentence:
        return True
    
    # --- 패턴 2: 뒤 셀이 한국어 연결 표현으로 시작 ---
    ko_continuation_starts = (
        '도 ', '를 ', '을 ', '에 ', '의 ', '와 ', '과 ', '으로', '에서',
        '하여', '하고', '하는', '되는', '된 ', '한 ', '할 ', '다고',
        '부 ', '등 ', '이상', '이하', '위한', '관한')
    if any(last.startswith(s) for s in ko_continuation_starts):
        return True
    
    # --- 패턴 3: 괄호/인용부호가 셀 경계에서 분리 ---
    # 앞 셀이 여는 괄호/인용부호로 끝나거나, 뒤 셀이 닫는 괄호로 시작
    if first_stripped.endswith(('「', '(', '[', '「')) and not last_stripped.startswith(('「', '(', '[')):
        return True
    if last.startswith(('」', ')', ']')) or last.startswith(('」 ,', ') ,', '] ,')):
        return True
    
    # --- 패턴 4: 셀을 합쳤을 때 자연스러운 문장, 개별 셀은 불완전 ---
    if len(cells) == 2:
        last_words = last.split()
        first_words = first.split()
        # 뒤 셀이 매우 짧고 (3단어 이하), 앞 셀이 문장 종결이 아님
        if len(last_words) <= 3 and len(first_words) >= 3:
            if not any(first_stripped.endswith(e) for e in ko_endings):
                return True
    
    # --- 패턴 5: 3개 이상 셀, 중간 셀이 매우 짧음 (줄바꿈으로 쪼개진 것) ---
    if len(cells) >= 3:
        mid_cells = cells[1:-1]
        short_mids = sum(1 for c in mid_cells if len(c.split()) <= 2)
        if short_mids >= len(mid_cells) * 0.5:
            # 중간 셀 대부분이 1-2단어면 분할된 문장
            combined = ' '.join(cells)
            # 합쳤을 때 너무 짧지 않으면
            if len(combined) > 20:
                return True
    
    return False


def extract_tables_from_page(doc, page_num: int = 0, debug: bool = False, **kwargs) -> List[Table]:
    """
    페이지에서 테이블 추출
    
    1차: 그리드 라인(수평/수직선) 기반 감지 (가장 정확)
    2차: 텍스트 좌표 패턴 기반 감지 (폴백)
    """
    from .. import extract_text_with_positions, get_pages
    
    pages = get_pages(doc)
    if page_num >= len(pages):
        return []
    
    text_items = extract_text_with_positions(doc, page_num)
    if not text_items:
        return []
    
    tables = _extract_tables_internal(doc, page_num, text_items, debug, **kwargs)
    
    # 모든 표에 page 번호 설정 (1-based)
    for t in tables:
        t.page = page_num + 1
    
    return tables


def _extract_tables_internal(doc, page_num, text_items, debug=False, **kwargs) -> List[Table]:
    
    # 1차: 그리드 기반 감지 (수평선 + 수직선 모두 있는 표)
    try:
        from .._grid_table import detect_tables_by_grid
        grid_tables = detect_tables_by_grid(doc, page_num, text_items, debug=debug)
        if grid_tables:
            return grid_tables
    except Exception:
        pass
    
    # 1.5차: 수평선 기반 borderless 표 감지 (수평선은 있지만 수직선 없는 표)
    hline_tables = []
    try:
        from .. import _extract_page_lines, _get_page_dimensions
        from .._grid_table import _detect_hline_tables
        h_lines, v_lines = _extract_page_lines(doc, page_num)
        if len(h_lines) >= 3:
            page_w, page_h = _get_page_dimensions(doc, page_num)
            hline_tables = _detect_hline_tables(h_lines, text_items, page_w, page_h, debug)
    except Exception:
        pass
    
    # 1.7차: 텍스트 x좌표 정렬 기반 표 감지 (선 없는 표)
    align_tables = []
    try:
        from .._grid_table import detect_tables_by_alignment
        from .. import _get_page_dimensions
        pw, ph = _get_page_dimensions(doc, page_num)
        align_tables = detect_tables_by_alignment(text_items, debug=debug, 
                                                   page_width=pw, page_height=ph)
    except Exception:
        pass
    
    # hline과 alignment 결과 중 더 나은 것 선택
    # 기준: 더 큰 표(셀 수)를 우선
    all_found = hline_tables + align_tables
    if all_found:
        # 겹치는 표 제거 (Y+X 범위 모두 겹치면 큰 표만 유지)
        all_found.sort(key=lambda t: len(t.cells), reverse=True)
        filtered = []
        for t in all_found:
            overlap = False
            for existing in filtered:
                # Y범위 겹침
                y_ov = max(0, min(t.y + t.height, existing.y + existing.height) - max(t.y, existing.y))
                min_h = min(t.height, existing.height) if min(t.height, existing.height) > 0 else 1
                # X범위 겹침
                x_ov = max(0, min(t.x + t.width, existing.x + existing.width) - max(t.x, existing.x))
                min_w = min(t.width, existing.width) if min(t.width, existing.width) > 0 else 1
                # Y와 X 모두 30%+ 겹쳐야 같은 표
                if y_ov > min_h * 0.3 and x_ov > min_w * 0.3:
                    overlap = True
                    break
            if not overlap:
                filtered.append(t)
        if filtered:
            return filtered
    
    # 2차: 텍스트 좌표 기반 감지 (기존 로직)
    return detect_tables(text_items, debug=debug, **kwargs)


# 테스트
if __name__ == '__main__':
    print("테이블 감지 모듈 로드됨")
    groups = {}
    current_group = sorted_values[0]