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
        # 2. 수직 간격 100pt 이내
        # 3. 또는 열 개수가 같음
        is_overlap = overlap > (min_width * 0.7)
        is_gap_ok = -30 < gap < 100
        is_same_cols = current.cols == next_t.cols
        
        should_merge = is_overlap and (is_gap_ok or is_same_cols)
        
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
    
    return True


def extract_tables_from_page(doc, page_num: int = 0, debug: bool = False, **kwargs) -> List[Table]:
    """
    페이지에서 테이블 추출
    """
    from .. import extract_text_with_positions, get_pages
    
    pages = get_pages(doc)
    if page_num >= len(pages):
        return []
    
    # 텍스트 추출 (이미 Top-Down으로 정규화됨)
    text_items = extract_text_with_positions(doc, page_num)
    
    if not text_items:
        return []
    
    return detect_tables(text_items, debug=debug, **kwargs)


# 테스트
if __name__ == '__main__':
    print("테이블 감지 모듈 로드됨")
    groups = {}
    current_group = sorted_values[0]
    
