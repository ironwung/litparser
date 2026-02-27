"""
Grid-based Table Detection

PDF 수평/수직 라인(그래픽 선)으로 표 셀 영역을 구성하고,
각 셀에 텍스트를 매핑하여 표를 추출합니다.

텍스트 좌표 기반 감지보다 훨씬 정확합니다.
"""

from typing import List, Tuple, Dict
from collections import defaultdict
from .core.table_detector import Table, TableCell


def detect_tables_by_grid(doc, page_num: int, text_items: list, 
                          debug: bool = False) -> List[Table]:
    """
    PDF 그래픽 라인으로 표를 감지합니다.
    
    Args:
        doc: PDFDocument
        page_num: 페이지 번호
        text_items: TextItem 리스트 (top-down 좌표)
        debug: 디버그 출력
    
    Returns:
        List[Table]: 감지된 표 목록 (없으면 빈 리스트)
    """
    from . import _extract_page_lines, _get_page_dimensions
    
    h_lines, v_lines = _extract_page_lines(doc, page_num)
    page_w, page_h = _get_page_dimensions(doc, page_num)
    
    if len(v_lines) < 3 or len(h_lines) < 3:
        return []
    
    # 전체 폭 수평선 제외 (페이지 경계)
    h_real = [(x0, y, x1) for x0, y, x1 in h_lines
              if not (x0 < 1 and x1 > page_w - 2)]
    
    if len(h_real) < 3:
        return []
    
    # 수직선: 페이지 테두리(높이 40%+)와 내부 라인 분리
    border_height_threshold = page_h * 0.4
    v_border = [(x, y0, y1) for x, y0, y1 in v_lines
                if (y1 - y0) >= border_height_threshold]
    v_inner = [(x, y0, y1) for x, y0, y1 in v_lines
               if (y1 - y0) < border_height_threshold]
    
    # 내부 수직선의 X 클러스터
    inner_col_xs = _cluster_values([x for x, _, _ in v_inner], tolerance=4)
    # 전체(테두리 포함) 수직선 X 클러스터  
    all_col_xs = _cluster_values([x for x, _, _ in v_lines], tolerance=4)
    
    # 내부 수직선이 4개 이상 있으면 → 진짜 표 (내부 라인으로 그리드 구성)
    # 내부 수직선이 적으면 → 전체 라인으로 시도하되 검증 강화
    if len(inner_col_xs) >= 4:
        col_xs = inner_col_xs
    elif len(all_col_xs) >= 3:
        col_xs = all_col_xs
    else:
        return []
    
    # 2. 수평선 Y좌표 클러스터링 (PDF 좌표)
    row_ys_pdf = _cluster_values([y for _, y, _ in h_real], tolerance=4)
    
    # 최소 그리드: 열 3+, 행 4+
    if len(col_xs) < 3 or len(row_ys_pdf) < 4:
        return []
    
    # PDF Y → top-down Y 변환
    row_ys_td = sorted([page_h - y for y in row_ys_pdf])
    col_xs = sorted(col_xs)
    
    num_rows = len(row_ys_td) - 1
    num_cols = len(col_xs) - 1
    
    if num_rows < 2 or num_cols < 2:
        return []
    
    if debug:
        print(f"[Grid] {num_rows} rows x {num_cols} cols")
        print(f"  col_xs: {[round(x,1) for x in col_xs]}")
        print(f"  row_ys_td: {[round(y,1) for y in row_ys_td]}")
    
    # 3. 각 셀에 텍스트 매핑
    cells = []
    for r in range(num_rows):
        y_top = row_ys_td[r]
        y_bot = row_ys_td[r + 1]
        
        for c in range(num_cols):
            x_left = col_xs[c]
            x_right = col_xs[c + 1]
            
            # 셀 너비가 너무 작으면 스킵 (5pt 미만)
            if x_right - x_left < 5:
                continue
            
            # 셀 영역에 속하는 텍스트 수집
            cell_texts = []
            for item in text_items:
                # 아이템 중심이 셀 내에 있는지
                if (x_left - 3 <= item.x <= x_right + 3 and
                    y_top - 3 <= item.y <= y_bot + 3):
                    cell_texts.append(item)
            
            # Y→X 순으로 정렬 후 텍스트 합치기
            cell_texts.sort(key=lambda t: (t.y, t.x))
            text = _merge_cell_texts(cell_texts)
            
            cells.append(TableCell(
                row=r, col=c, text=text,
                x=x_left, y=y_top,
                width=x_right - x_left,
                height=y_bot - y_top
            ))
    
    if not cells:
        return []
    
    # 4. 유효성 검증
    table = Table(
        cells=cells,
        rows=num_rows,
        cols=num_cols,
        x=col_xs[0],
        y=row_ys_td[0],
        width=col_xs[-1] - col_xs[0],
        height=row_ys_td[-1] - row_ys_td[0],
        method='grid'
    )
    
    non_empty = sum(1 for c in cells if c.text.strip())
    total = len(cells)
    
    # 기본 조건
    if total == 0 or num_rows < 2 or num_cols < 2:
        return []
    
    fill_rate = non_empty / total if total > 0 else 0
    
    # 내부 수직선이 적은 경우(테두리만 사용): 높은 채움률 요구
    if len(inner_col_xs) < 4:
        if fill_rate < 0.15:
            return []
        if num_rows < 5:
            return []
    else:
        if fill_rate < 0.05:
            return []
    
    # 대조표/해설 페이지 필터: 열이 2~4개이고 셀 텍스트가 대부분 긴 문장이면
    # 데이터 테이블이 아니라 대조표 형태 → 그리드 표로 반환하지 않음
    # (이미 _extract_text_table_columns에서 컬럼 분리 처리됨)
    if num_cols <= 4:
        grid = table.to_list()
        long_prose_cells = 0
        total_nonempty = 0
        for row in grid:
            for cell_text in row:
                ct = cell_text.strip()
                if not ct:
                    continue
                total_nonempty += 1
                # 20자 이상이면 긴 문장형 셀
                if len(ct) > 20:
                    long_prose_cells += 1
        if total_nonempty > 0 and long_prose_cells > total_nonempty * 0.25:
            return []
    
    if debug:
        print(f"[Grid] Valid table: {num_rows}x{num_cols}, "
              f"fill={fill_rate:.1%}, inner_cols={len(inner_col_xs)}")
    
    return [table]


def _cluster_values(values: List[float], tolerance: float = 4) -> List[float]:
    """
    근접한 값들을 클러스터링해서 대표값 반환.
    
    예: [59.5, 59.8, 60.1, 124.5, 124.8] → [59.8, 124.7]
    """
    if not values:
        return []
    
    sorted_vals = sorted(values)
    clusters = []
    current_cluster = [sorted_vals[0]]
    
    for v in sorted_vals[1:]:
        if v - current_cluster[-1] <= tolerance:
            current_cluster.append(v)
        else:
            clusters.append(current_cluster)
            current_cluster = [v]
    clusters.append(current_cluster)
    
    # 대표값: 중앙값
    result = []
    for cluster in clusters:
        # 최소 2개 이상의 라인이 모인 클러스터만 사용 (노이즈 제거)
        if len(cluster) >= 2:
            result.append(sum(cluster) / len(cluster))
    
    return sorted(result)


def _merge_cell_texts(items: list) -> str:
    """셀 내 TextItem들을 줄 단위로 합쳐서 문자열 반환."""
    if not items:
        return ""
    
    # Y 기준 줄 분리 (tolerance 5pt)
    lines = []
    current_line = []
    current_y = None
    
    for item in items:
        if current_y is None or abs(item.y - current_y) > 5:
            if current_line:
                lines.append(current_line)
            current_line = [item]
            current_y = item.y
        else:
            current_line.append(item)
    if current_line:
        lines.append(current_line)
    
    # 각 줄에서 X순 정렬 후 텍스트 합치기
    result_lines = []
    for line in lines:
        line.sort(key=lambda t: t.x)
        parts = []
        for item in line:
            text = item.text
            if text:
                parts.append(text)
        if parts:
            result_lines.append(' '.join(parts))
    
    return '\n'.join(result_lines)
