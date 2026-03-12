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
    한 페이지에 여러 독립 표가 있을 수 있으므로, 수직선 Y범위로 표 영역을 분리합니다.
    """
    from . import _extract_page_lines, _get_page_dimensions
    
    h_lines, v_lines = _extract_page_lines(doc, page_num)
    page_w, page_h = _get_page_dimensions(doc, page_num)
    
    if len(v_lines) < 3 or len(h_lines) < 3:
        if len(h_lines) >= 3:
            return _detect_hline_tables(h_lines, text_items, page_w, page_h, debug)
        return []
    
    # h_lines 형식: (y, x_min, x_max) in PDF coordinates
    # 페이지 전체 폭 수평선 제외 (페이지 경계)
    h_real = [(y, x0, x1) for y, x0, x1 in h_lines
              if not (x0 < 1 and x1 > page_w - 2)]
    
    border_height_threshold = page_h * 0.4
    v_inner = [(x, y0, y1) for x, y0, y1 in v_lines
               if (y1 - y0) < border_height_threshold]
    
    # 페이지 경계 수직선 제거 (x < 2 또는 x > page_w - 2)
    v_inner = [(x, y0, y1) for x, y0, y1 in v_inner 
               if 2 < x < page_w - 2]
    
    if not v_inner or not h_real:
        if len(h_real) >= 3:
            return _detect_hline_tables(h_lines, text_items, page_w, page_h, debug)
        return []
    
    # 수직선을 Y범위가 겹치는 것끼리 그룹화하여 독립 표 영역 분리
    table_regions = _group_lines_into_regions(v_inner, h_real, page_h)
    
    if debug:
        print(f"[Grid] {len(table_regions)} table regions found")
    
    # 각 영역별로 표 구성
    all_tables = []
    for region in table_regions:
        r_v_lines = region['v_lines']
        r_h_lines = region['h_lines']
        
        table = _build_grid_table_from_region(
            r_v_lines, r_h_lines, text_items, page_w, page_h, debug)
        if table:
            all_tables.append(table)
    
    return all_tables


def _group_lines_into_regions(v_inner, h_real, page_h):
    """
    수직선을 Y범위가 겹치는 것끼리 그룹화하여 독립 표 영역을 분리.
    각 영역에 해당하는 수평선도 매핑.
    """
    if not v_inner:
        return []
    
    # 수직선을 Y범위 기준으로 클러스터링
    # Y범위가 겹치면 같은 그룹
    groups = []
    for vl in v_inner:
        x, y0, y1 = vl
        y_min_td = page_h - y1  # top-down
        y_max_td = page_h - y0
        
        merged = False
        for g in groups:
            # Y범위가 겹치는지 확인
            overlap = max(0, min(y_max_td, g['y_max']) - max(y_min_td, g['y_min']))
            if overlap > 0:
                g['v_lines'].append(vl)
                g['y_min'] = min(g['y_min'], y_min_td)
                g['y_max'] = max(g['y_max'], y_max_td)
                merged = True
                break
        
        if not merged:
            groups.append({
                'v_lines': [vl],
                'y_min': y_min_td,
                'y_max': y_max_td
            })
    
    # 그룹 병합 (Y범위가 겹치는 그룹끼리)
    changed = True
    while changed:
        changed = False
        new_groups = []
        used = set()
        for i, g1 in enumerate(groups):
            if i in used:
                continue
            for j, g2 in enumerate(groups):
                if j <= i or j in used:
                    continue
                overlap = max(0, min(g1['y_max'], g2['y_max']) - max(g1['y_min'], g2['y_min']))
                if overlap > 0:
                    g1['v_lines'].extend(g2['v_lines'])
                    g1['y_min'] = min(g1['y_min'], g2['y_min'])
                    g1['y_max'] = max(g1['y_max'], g2['y_max'])
                    used.add(j)
                    changed = True
            new_groups.append(g1)
        groups = new_groups
    
    # 각 그룹에 해당하는 수평선 매핑
    for g in groups:
        y_min_pdf = page_h - g['y_max']
        y_max_pdf = page_h - g['y_min']
        margin = 5
        g['h_lines'] = [(y, x0, x1) for y, x0, x1 in h_real
                        if y_min_pdf - margin <= y <= y_max_pdf + margin]
    
    # 수직선이 3개+이고 수평선이 3개+ 인 그룹만
    regions = [g for g in groups 
               if len(g['v_lines']) >= 3 and len(g['h_lines']) >= 3]
    
    return regions


def _build_grid_table_from_region(v_lines, h_lines, text_items, 
                                   page_w, page_h, debug=False):
    """한 표 영역에서 그리드 테이블을 구성."""
    # 열 좌표 클러스터링
    col_xs = _cluster_values([x for x, _, _ in v_lines], tolerance=4)
    
    # 행 좌표 클러스터링 (PDF 좌표 → top-down)
    # h_lines 형식: (y, x_min, x_max)
    # 행 좌표: 클러스터링 후 체인 방지 (최대 범위 8pt)
    raw_ys = [y for y, _, _ in h_lines]
    row_ys_pdf = _cluster_values_bounded(raw_ys, tolerance=4, max_span=8)
    
    if len(col_xs) < 3 or len(row_ys_pdf) < 3:
        return None
    
    row_ys_td = sorted([page_h - y for y in row_ys_pdf])
    col_xs = sorted(col_xs)
    
    # 마지막 수평선 아래에 텍스트가 있으면 추가 행 경계 (표가 아래쪽 수평선 없이 끝나는 경우)
    # 수직선의 최대 Y(top-down)를 하한으로 사용
    v_max_td = max(page_h - y0 for x, y0, y1 in v_lines)
    last_hline_td = row_ys_td[-1] if row_ys_td else 0
    
    # 마지막 수평선 아래 텍스트 검색 (수평선 Y + 3pt 이후만, 최대 50pt)
    import re as _re_grid
    below_items = [it for it in text_items
                   if last_hline_td + 3 < it.y < last_hline_td + 50 
                   and col_xs[0] - 5 <= it.x <= col_xs[-1] + 5
                   and not _re_grid.match(r'^-\s*\d+\s*-$', it.text.strip())]
    
    if below_items:
        from collections import defaultdict
        y_groups_below = defaultdict(list)
        for it in below_items:
            y_key = round(it.y / 5) * 5
            y_groups_below[y_key].append(it)
        
        prev_y = last_hline_td
        for yk in sorted(y_groups_below.keys()):
            if yk - prev_y > 30:
                break
            row_text = ' '.join(it.text.strip() for it in y_groups_below[yk]).strip()
            if _re_grid.match(r'^-\s*\d+\s*-$', row_text):
                break
            row_ys_td.append(yk)
            prev_y = yk
        
        if len(row_ys_td) > len(row_ys_pdf) + 1:
            row_ys_td.append(prev_y + 15)
    
    # 근접한 행 경계 병합 (15pt 이내)
    if len(row_ys_td) > 2:
        merged = [row_ys_td[0]]
        for y in row_ys_td[1:]:
            if y - merged[-1] > 15:
                merged.append(y)
        row_ys_td = merged
    
    # 첫 수평선 위에 텍스트가 있으면 추가 행 경계
    v_min_td = min(page_h - y1 for x, y0, y1 in v_lines)
    first_hline_td = row_ys_td[0] if row_ys_td else page_h
    if v_min_td < first_hline_td - 5:
        has_text_above = any(
            v_min_td - 5 < it.y < first_hline_td and col_xs[0] - 5 <= it.x <= col_xs[-1] + 5
            for it in text_items
        )
        if has_text_above:
            row_ys_td.insert(0, v_min_td - 2)
    
    num_rows = len(row_ys_td) - 1
    num_cols = len(col_xs) - 1
    
    if num_rows < 1 or num_cols < 1:
        return None
    
    if debug:
        print(f"[Grid] {num_rows} rows x {num_cols} cols")
        print(f"  col_xs: {[round(x,1) for x in col_xs]}")
        print(f"  row_ys_td: {[round(y,1) for y in row_ys_td]}")
    
    # 셀에 텍스트 매핑
    cells = []
    for r in range(num_rows):
        y_top = row_ys_td[r]
        y_bot = row_ys_td[r + 1]
        
        for c in range(num_cols):
            x_left = col_xs[c]
            x_right = col_xs[c + 1]
            
            if x_right - x_left < 5:
                continue
            
            cell_texts = []
            for item in text_items:
                if (x_left - 3 <= item.x <= x_right + 3 and
                    y_top - 3 <= item.y <= y_bot + 3):
                    cell_texts.append(item)
            
            cell_texts.sort(key=lambda t: (t.y, t.x))
            text = _merge_cell_texts(cell_texts)
            
            cells.append(TableCell(
                row=r, col=c, text=text,
                x=x_left, y=y_top,
                width=x_right - x_left,
                height=y_bot - y_top
            ))
    
    if not cells:
        return None
    
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
    
    # 후처리 1: 첫 행이 수평선 바깥의 본문이면 제거
    # 첫 수평선 Y(top-down)보다 위에 있는 텍스트만으로 구성된 첫 행 제거
    if num_rows >= 2 and len(row_ys_pdf) >= 2:
        first_hline_td = row_ys_td[0]  # 첫 수평선 = 첫 행 경계
        first_row_cells = [c for c in cells if c.row == 0]
        first_row_nonempty = [c for c in first_row_cells if c.text.strip()]
        
        if first_row_nonempty:
            # 첫 행의 텍스트가 모두 수평선 위에만 있는지 검사
            # text_items에서 첫 행 영역 내의 아이템 Y를 확인
            first_row_items = [it for it in text_items
                              if first_hline_td - 5 <= it.y <= row_ys_td[1] + 3
                              and col_xs[0] - 3 <= it.x <= col_xs[-1] + 3]
            above_line_items = [it for it in first_row_items if it.y < first_hline_td]
            below_line_items = [it for it in first_row_items if it.y >= first_hline_td]
            
            # 수평선 위 아이템만 있고 아래 아이템이 없으면 → 본문, 제거
            if above_line_items and not below_line_items:
                cells = [c for c in cells if c.row > 0]
                for c in cells:
                    c.row -= 1
                num_rows -= 1
    
    # 후처리 2: 마지막 행이 이전 행과 텍스트 포함(중복) 관계이면 제거
    if num_rows >= 2:
        last_row_cells = [c for c in cells if c.row == num_rows - 1]
        prev_row_cells = [c for c in cells if c.row == num_rows - 2]
        
        last_texts = set()
        for c in last_row_cells:
            for word in c.text.strip().split():
                if word:
                    last_texts.add(word)
        
        prev_texts = set()
        for c in prev_row_cells:
            for word in c.text.strip().split():
                if word:
                    prev_texts.add(word)
        
        # 마지막 행의 단어 80%+ 가 이전 행에 이미 포함되면 중복
        if last_texts and prev_texts:
            overlap = len(last_texts & prev_texts) / len(last_texts)
            if overlap >= 0.8:
                cells = [c for c in cells if c.row < num_rows - 1]
                num_rows -= 1
    
    # 후처리 결과 반영
    table.cells = cells
    table.rows = num_rows
    
    # 유효성 검증
    non_empty = sum(1 for c in cells if c.text.strip())
    total = len(cells)
    fill_rate = non_empty / total if total > 0 else 0
    
    if fill_rate < 0.10:
        return None
    if num_rows < 2 or num_cols < 2:
        return None
    
    if debug:
        print(f"[Grid] Valid table: {num_rows}x{num_cols}, fill={fill_rate:.1%}")
    
    return table


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


def _cluster_values_bounded(values: List[float], tolerance: float = 4, 
                             max_span: float = 8) -> List[float]:
    """
    클러스터링 + 최대 범위 제한.
    체인 연결(424→425→426→...→462)을 방지하기 위해 
    클러스터 내 첫 값~마지막 값 차이가 max_span을 넘으면 분할.
    """
    if not values:
        return []
    
    sorted_vals = sorted(values)
    clusters = []
    current_cluster = [sorted_vals[0]]
    
    for v in sorted_vals[1:]:
        # 이전 값과 tolerance 이내이고 클러스터 범위가 max_span 이내
        if (v - current_cluster[-1] <= tolerance and 
            v - current_cluster[0] <= max_span):
            current_cluster.append(v)
        else:
            clusters.append(current_cluster)
            current_cluster = [v]
    clusters.append(current_cluster)
    
    result = []
    for cluster in clusters:
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


def _detect_hline_tables(h_lines: list, text_items: list, 
                          page_w: float, page_h: float,
                          debug: bool = False) -> List[Table]:
    """
    수평선 + 텍스트 x좌표 정렬 기반 borderless 표 감지.
    
    수직선(v_lines)이 없지만 수평선(h_lines)이 있는 표를 감지합니다.
    열 구분은 텍스트의 x좌표 정렬 패턴으로 추론합니다.
    
    동작:
    1. 수평선을 x범위가 비슷한 그룹으로 묶어 표 영역 후보 생성
    2. 각 후보 영역 내 텍스트의 x좌표 클러스터로 열 감지
    3. 행은 수평선 Y좌표로 구분
    4. 유효성 검증
    """
    if not h_lines or not text_items:
        return []
    
    # h_lines: (y_pdf, x_start, x_end)
    # top-down Y로 변환하고 폭이 의미 있는 선만 필터
    hlines_td = []
    for y_pdf, x0, x1 in h_lines:
        y_td = page_h - y_pdf
        w = x1 - x0
        if w >= 30:  # 30pt 이상
            hlines_td.append((y_td, x0, x1, w))
    
    if len(hlines_td) < 3:
        return []
    
    # 1. 수평선을 x범위가 비슷한 그룹으로 분류
    # 같은 표의 수평선은 x_start~x_end가 유사함
    line_groups = _group_hlines_by_xrange(hlines_td)
    
    if debug:
        print(f"[HLine] {len(hlines_td)} hlines -> {len(line_groups)} groups")
    
    tables = []
    for group in line_groups:
        if len(group) < 3:  # 최소 3개 수평선 (헤더 + 구분 + 본문)
            continue
        
        table = _build_table_from_hlines(group, text_items, page_w, page_h, debug)
        if table:
            tables.append(table)
    
    # 겹치는 표 제거 (Y범위가 50% 이상 겹치면 큰 표만 유지)
    if len(tables) > 1:
        tables.sort(key=lambda t: len(t.cells), reverse=True)
        filtered = []
        for t in tables:
            overlap = False
            for existing in filtered:
                # Y범위 겹침 계산
                y_overlap = max(0, min(t.y + t.height, existing.y + existing.height) - max(t.y, existing.y))
                min_h = min(t.height, existing.height)
                if min_h > 0 and y_overlap > min_h * 0.5:
                    overlap = True
                    break
            if not overlap:
                filtered.append(t)
        tables = filtered
    
    return tables


def _group_hlines_by_xrange(hlines_td: list, x_tol: float = 30) -> List[List]:
    """
    x범위가 비슷한 수평선끼리 그룹화.
    같은 표의 수평선은 x_start와 x_end가 유사함.
    """
    if not hlines_td:
        return []
    
    # (y_td, x0, x1, w) 를 x0 기준으로 클러스터링
    sorted_lines = sorted(hlines_td, key=lambda l: (l[1], l[2]))
    
    groups = []
    current_group = [sorted_lines[0]]
    
    for line in sorted_lines[1:]:
        ref = current_group[0]
        # x_start와 x_end가 모두 tol 이내면 같은 그룹
        if abs(line[1] - ref[1]) < x_tol and abs(line[2] - ref[2]) < x_tol:
            current_group.append(line)
        else:
            groups.append(current_group)
            current_group = [line]
    groups.append(current_group)
    
    return groups


def _build_table_from_hlines(hlines: list, text_items: list,
                              page_w: float, page_h: float,
                              debug: bool = False):
    """
    수평선 그룹 + 텍스트 x정렬로 표 구성.
    
    수평선 사이 각 행에서 텍스트가 일정한 x좌표에 정렬되어 있는지 검증.
    정렬된 행만 표로 인정.
    """
    # 수평선 Y좌표로 행 경계 설정
    row_ys = sorted(set(round(l[0], 1) for l in hlines))
    
    if len(row_ys) < 3:
        return None
    
    # 표 영역의 x범위
    table_x0 = min(l[1] for l in hlines) 
    table_x1 = max(l[2] for l in hlines)
    
    # 수평선 사이 각 행에서 텍스트 수집 & 정렬 검증
    # 데이터 행: 여러 개의 짧은 텍스트가 서로 다른 x좌표에 배치
    # 본문 행: 긴 텍스트가 왼쪽 정렬
    valid_rows = []
    all_row_items = []
    
    for r in range(len(row_ys) - 1):
        y_top = row_ys[r]
        y_bot = row_ys[r + 1]
        
        # 행 높이가 너무 크면 (60pt+) 본문 영역일 가능성
        if y_bot - y_top > 60:
            continue
        
        # 행 내 텍스트 수집
        row_items = [it for it in text_items
                    if table_x0 - 5 <= it.x <= table_x1 + 5
                    and y_top - 2 <= it.y <= y_bot - 2]
        
        if not row_items:
            continue
        
        # 긴 텍스트(50자+)가 있으면 본문 행 → 스킵
        has_long_text = any(len(it.text.strip()) > 50 for it in row_items)
        if has_long_text:
            continue
        
        # 행 내 x좌표 다양성 검증
        # 데이터 행: x좌표가 2개 이상의 클러스터에 분포 (열 구분)
        x_positions = sorted(set(round(it.x / 10) * 10 for it in row_items))
        
        if len(x_positions) >= 2:
            # x 범위가 표 폭의 30% 이상 → 열 구조 있음
            x_spread = max(x_positions) - min(x_positions)
            table_width = table_x1 - table_x0
            if x_spread >= table_width * 0.3:
                valid_rows.append(r)
                all_row_items.extend(row_items)
    
    if len(valid_rows) < 2:
        return None
    
    # 연속된 유효 행 그룹 찾기 (중간에 본문 행이 끼면 분리)
    # 인접한 수평선 인덱스가 연속인 유효 행만 하나의 표로
    consecutive_groups = []
    current_group = [valid_rows[0]]
    for r in valid_rows[1:]:
        if r == current_group[-1] + 1:
            current_group.append(r)
        else:
            consecutive_groups.append(current_group)
            current_group = [r]
    consecutive_groups.append(current_group)
    
    # 가장 긴 연속 그룹 선택
    best_group = max(consecutive_groups, key=len)
    if len(best_group) < 2:
        return None
    
    # 유효한 행 내부를 Y좌표 기반으로 세분화 (sub-rows)
    sub_row_ys = set()
    for r in best_group:
        y_top = row_ys[r]
        y_bot = row_ys[r + 1]
        sub_row_ys.add(y_top)
        
        row_items = [it for it in text_items
                    if table_x0 - 5 <= it.x <= table_x1 + 5
                    and y_top - 2 <= it.y <= y_bot - 2]
        if row_items:
            text_ys = sorted(set(round(it.y / 5) * 5 for it in row_items))
            for ty in text_ys:
                if ty > y_top + 2:
                    sub_row_ys.add(ty)
    
    sub_row_ys.add(row_ys[max(best_group) + 1])
    valid_row_ys = sorted(sub_row_ys)
    
    if len(valid_row_ys) < 3:
        return None
    
    table_y_top = min(valid_row_ys)
    table_y_bot = max(valid_row_ys)
    
    # 열 경계 추론 (유효한 행의 텍스트만 사용)
    col_xs = _find_column_boundaries_from_text(all_row_items, table_x0, table_x1)
    
    if len(col_xs) < 2:
        return None
    
    num_rows = len(valid_row_ys) - 1
    num_cols = len(col_xs) - 1
    
    if num_rows < 2 or num_cols < 1:
        return None
    
    if debug:
        print(f"[HLine] table: {num_rows}x{num_cols} "
              f"y={table_y_top:.0f}~{table_y_bot:.0f} "
              f"x={table_x0:.0f}~{table_x1:.0f}")
        print(f"  col_xs: {[round(x,1) for x in col_xs]}")
        print(f"  row_ys: {[round(y,1) for y in valid_row_ys]}")
    
    # 셀에 텍스트 매핑
    cells = []
    for r in range(num_rows):
        y_top = valid_row_ys[r]
        y_bot = valid_row_ys[r + 1]
        
        for c in range(num_cols):
            x_left = col_xs[c]
            x_right = col_xs[c + 1]
            
            if x_right - x_left < 5:
                continue
            
            cell_texts = [it for it in text_items
                         if x_left - 5 <= it.x <= x_right + 5
                         and y_top - 2 <= it.y <= y_bot - 2]
            
            cell_texts.sort(key=lambda t: (t.y, t.x))
            text = _merge_cell_texts(cell_texts)
            
            cells.append(TableCell(
                row=r, col=c, text=text,
                x=x_left, y=y_top,
                width=x_right - x_left,
                height=y_bot - y_top
            ))
    
    if not cells:
        return None
    
    table = Table(
        cells=cells,
        rows=num_rows,
        cols=num_cols,
        x=table_x0,
        y=table_y_top,
        width=table_x1 - table_x0,
        height=table_y_bot - table_y_top,
        method='hline'
    )
    
    # 유효성 검증
    non_empty = sum(1 for c in cells if c.text.strip())
    total = len(cells)
    fill_rate = non_empty / total if total > 0 else 0
    
    if fill_rate < 0.20:
        return None
    
    return table


def _find_column_boundaries_from_text(items: list, table_x0: float, 
                                       table_x1: float) -> List[float]:
    """
    텍스트 x좌표 정렬 패턴에서 열 경계를 추론.
    
    방법: x좌표를 클러스터링하여 큰 갭(30pt+)을 열 경계로 사용.
    """
    from collections import Counter
    
    if not items:
        return []
    
    # x좌표를 5pt 단위로 binning
    bin_size = 5
    x_bins = Counter(round(it.x / bin_size) * bin_size for it in items)
    
    if not x_bins:
        return []
    
    # 빈도 2 이상인 x좌표만 (노이즈 제거)
    significant_xs = sorted(x for x, cnt in x_bins.items() if cnt >= 2)
    
    if len(significant_xs) < 2:
        return []
    
    # 인접한 x좌표를 클러스터로 묶기 (30pt 이내 → 같은 열)
    clusters = []
    current = [significant_xs[0]]
    
    for x in significant_xs[1:]:
        if x - current[-1] <= 30:
            current.append(x)
        else:
            clusters.append(current)
            current = [x]
    clusters.append(current)
    
    if len(clusters) < 2:
        return []
    
    # 각 클러스터의 시작점 = 열 왼쪽 경계
    boundaries = [min(c) - 5 for c in clusters]
    boundaries.append(table_x1 + 5)
    
    return sorted(boundaries)


def detect_tables_by_alignment(text_items: list, debug: bool = False,
                                page_width: float = 0, page_height: float = 0) -> List[Table]:
    """
    텍스트 x좌표 정렬 패턴으로 borderless 표를 감지.
    
    2단 레이아웃이면 좌/우 각각에서 독립적으로 표를 찾음.
    """
    if not text_items or len(text_items) < 8:
        return []
    
    # 2단 분리: 페이지 중앙 30~70% 범위에서 가장 큰 x좌표 갭 찾기
    split_x = None
    if page_width > 0:
        from collections import Counter
        x_bins = Counter(round(it.x / 10) * 10 for it in text_items)
        sorted_xs = sorted(x_bins.keys())
        
        center_min = page_width * 0.25
        center_max = page_width * 0.75
        max_gap = 0
        for i in range(len(sorted_xs) - 1):
            gap_start = sorted_xs[i]
            gap_end = sorted_xs[i + 1]
            gap = gap_end - gap_start
            gap_mid = (gap_start + gap_end) / 2
            # 갭이 페이지 중앙 근처(25~75%)에 있고 30pt+ 이면
            if gap >= 30 and center_min <= gap_mid <= center_max and gap > max_gap:
                max_gap = gap
                split_x = (gap_start + gap_end) / 2
    
    if split_x and max_gap >= 30:
        left_items = [it for it in text_items if it.x < split_x]
        right_items = [it for it in text_items if it.x >= split_x]
        
        # 양쪽 모두 최소 아이템이 있어야 분리
        if len(left_items) >= 10 and len(right_items) >= 10:
            if debug:
                print(f"[Align] split at x={split_x:.0f} (gap={max_gap:.0f}) L={len(left_items)} R={len(right_items)}")
            
            tables = []
            for side_items, label in [(left_items, 'L'), (right_items, 'R')]:
                side_tables = _detect_aligned_tables_in_region(side_items, debug, label)
                tables.extend(side_tables)
            
            # 분리 후 결과가 있으면 반환, 없으면 전체로 시도
            if tables:
                return tables
    
    # 분리 안 하거나 분리 후 결과 없으면 전체로
    return _detect_aligned_tables_in_region(text_items, debug, '')


def _detect_aligned_tables_in_region(text_items: list, debug: bool, 
                                      label: str = '') -> List[Table]:
    """한 영역(좌 또는 우 컬럼) 내에서 텍스트 정렬 기반 표 감지."""
    if not text_items or len(text_items) < 6:
        return []
    
    # 1. 행 구성 (Y 5pt tolerance)
    y_groups = defaultdict(list)
    for it in text_items:
        y_key = round(it.y / 5) * 5
        y_groups[y_key].append(it)
    
    sorted_y_keys = sorted(y_groups.keys())
    if len(sorted_y_keys) < 4:
        return []
    
    # 2. 각 행의 데이터/비데이터 분류
    row_info = {}
    for yk in sorted_y_keys:
        items = y_groups[yk]
        
        # x좌표 클러스터 (30pt+ 떨어진 그룹)
        x_sorted = sorted(it.x for it in items)
        clusters = []
        if x_sorted:
            cur = [x_sorted[0]]
            for x in x_sorted[1:]:
                if x - cur[-1] > 30:
                    clusters.append(cur)
                    cur = [x]
                else:
                    cur.append(x)
            clusters.append(cur)
        
        max_text_len = max((len(it.text.strip()) for it in items), default=0)
        is_data = (len(clusters) >= 2 and max_text_len <= 50)
        
        row_info[yk] = {
            'is_data': is_data,
            'x_clusters': [min(c) for c in clusters] if clusters else [],
            'items': items
        }
    
    # 3. 연속 데이터 행 그룹 찾기 (중간에 1-2행 헤더 허용)
    data_regions = []
    current_region = []
    gap_count = 0
    
    for yk in sorted_y_keys:
        info = row_info[yk]
        if info['is_data']:
            if gap_count <= 2:
                current_region.append(yk)
            else:
                if len(current_region) >= 4:
                    data_regions.append(current_region)
                current_region = [yk]
            gap_count = 0
        else:
            gap_count += 1
            if gap_count > 2 and current_region:
                if len(current_region) >= 4:
                    data_regions.append(current_region)
                current_region = []
    
    if current_region and len(current_region) >= 4:
        data_regions.append(current_region)
    
    if not data_regions:
        return []
    
    if debug:
        print(f"[Align{label}] {len(data_regions)} data regions")
    
    # 4. 각 영역에서 표 구성
    tables = []
    for region_ys in data_regions:
        table = _build_table_from_alignment(region_ys, row_info, y_groups, sorted_y_keys, debug)
        if table:
            tables.append(table)
    
    return tables


def _build_table_from_alignment(region_ys: list, row_info: dict, 
                                 y_groups: dict, all_y_keys: list,
                                 debug: bool = False):
    """
    데이터 행 영역에서 표를 구성.
    
    영역 내의 모든 행(데이터 + 섹션 헤더)을 포함하되,
    열 구조는 데이터 행의 x좌표 정렬로 결정.
    """
    if len(region_ys) < 4:
        return None
    
    y_min = min(region_ys)
    y_max = max(region_ys)
    
    # 영역 내의 모든 행 (텍스트가 있는 y_key만, 데이터 + 비데이터 포함)
    all_region_ys = [yk for yk in all_y_keys if y_min <= yk <= y_max and yk in y_groups]
    
    # 데이터 행의 x 클러스터에서 열 경계 추론
    all_data_items = []
    for yk in region_ys:
        all_data_items.extend(row_info[yk]['items'])
    
    if not all_data_items:
        return None
    
    # x좌표 히스토그램 (10pt 빈) - 데이터 행만
    from collections import Counter
    x_bins = Counter(round(it.x / 10) * 10 for it in all_data_items)
    
    # 빈도 3+ 인 x좌표만
    significant_xs = sorted(x for x, cnt in x_bins.items() if cnt >= 3)
    
    if len(significant_xs) < 2:
        return None
    
    # 클러스터링 (30pt tolerance)
    clusters = []
    cur = [significant_xs[0]]
    for x in significant_xs[1:]:
        if x - cur[-1] <= 30:
            cur.append(x)
        else:
            clusters.append(cur)
            cur = [x]
    clusters.append(cur)
    
    if len(clusters) < 2:
        return None
    
    # 열 경계: 각 클러스터의 시작점
    col_boundaries = [min(c) - 5 for c in clusters]
    # 마지막 열의 끝: 가장 오른쪽 텍스트 + 여유
    max_x = max(it.x for it in all_data_items)
    col_boundaries.append(max_x + 50)
    
    num_cols = len(col_boundaries) - 1
    
    if num_cols < 2:
        return None
    
    # 행 경계: all_region_ys의 각 Y
    row_boundaries = sorted(all_region_ys)
    # 마지막 행의 하단: 다음 Y키 또는 +15
    next_y_after = None
    for yk in all_y_keys:
        if yk > y_max:
            next_y_after = yk
            break
    row_boundaries.append(next_y_after if next_y_after else y_max + 15)
    
    num_rows = len(row_boundaries) - 1
    
    if num_rows < 3:
        return None
    
    if debug:
        print(f"[Align] table: {num_rows}x{num_cols} "
              f"y={y_min:.0f}~{y_max:.0f}")
        print(f"  cols: {[round(x) for x in col_boundaries]}")
    
    # 셀에 텍스트 매핑
    cells = []
    for r in range(num_rows):
        y_top = row_boundaries[r]
        y_bot = row_boundaries[r + 1]
        
        # 이 행의 모든 아이템
        row_items = y_groups.get(row_boundaries[r], [])
        
        for c in range(num_cols):
            x_left = col_boundaries[c]
            x_right = col_boundaries[c + 1]
            
            # 셀 내 아이템
            cell_items = [it for it in row_items
                         if x_left - 5 <= it.x <= x_right + 5]
            
            cell_items.sort(key=lambda t: t.x)
            text = ' '.join(it.text for it in cell_items).strip()
            
            cells.append(TableCell(
                row=r, col=c, text=text,
                x=x_left, y=y_top,
                width=x_right - x_left,
                height=y_bot - y_top
            ))
    
    if not cells:
        return None
    
    # x 범위 계산
    x_min = col_boundaries[0]
    x_max = col_boundaries[-1]
    
    table = Table(
        cells=cells,
        rows=num_rows,
        cols=num_cols,
        x=x_min,
        y=y_min,
        width=x_max - x_min,
        height=y_max - y_min,
        method='alignment'
    )
    
    # 유효성 검증
    non_empty = sum(1 for c in cells if c.text.strip())
    total = len(cells)
    fill_rate = non_empty / total if total > 0 else 0
    
    if fill_rate < 0.20:
        return None
    
    # 2열 표는 본문 들여쓰기와 구분 불가 → 추가 검증
    if num_cols <= 2:
        # 짧은 셀(15자 이하) 또는 숫자 셀이 30%+ 있어야 표
        short_or_num = 0
        for c in cells:
            ct = c.text.strip()
            if not ct:
                continue
            if len(ct) <= 15:
                short_or_num += 1
            elif ct.replace(' ', '').replace(',', '').replace('.', '').replace('%', '').replace('-', '').isdigit():
                short_or_num += 1
        if non_empty > 0 and short_or_num < non_empty * 0.3:
            return None
    
    # 긴 텍스트(30자+) 셀 비율이 높으면 본문
    long_cells = sum(1 for c in cells if len(c.text.strip()) > 30)
    if non_empty > 0 and long_cells > non_empty * 0.5:
        return None
    
    # 수식 영역 필터: 단일 변수(a-z 한 글자)가 비정상적으로 많으면 수식
    import re as _re
    all_cell_text = ' '.join(c.text for c in cells if c.text.strip())
    single_vars = len(_re.findall(r'(?<![a-zA-Z])[a-z](?![a-zA-Z])', all_cell_text))
    if non_empty > 0 and single_vars > non_empty * 1.0:
        return None
    
    # 문장 조각 필터: 첫 2행(캡션 제외)이 문장의 일부이면 FP
    grid = table.to_list()
    if len(grid) >= 2:
        header_cells = []
        for row in grid[:3]:  # 첫 3행 검사
            for cell in row:
                ct = cell.strip()
                if ct:
                    # "Table N." 캡션은 제외
                    import re as _re2
                    if _re2.match(r'^Table\s+\d', ct):
                        continue
                    header_cells.append(ct)
            if len(header_cells) >= 3:
                break
        if header_cells:
            fragment_count = 0
            for ct in header_cells:
                is_fragment = False
                last_word = ct.rstrip().split()[-1] if ct.rstrip().split() else ''
                # 20자+ 이고 마지막 단어가 소문자이면서 문장 종결이 아닌 일반 단어
                # 단, 약어(3자 이하)나 대문자 포함 단어는 제외
                if len(ct) > 20 and last_word.islower() and len(last_word) > 3:
                    is_fragment = True
                elif ct[0].islower() and len(ct) > 10:
                    is_fragment = True
                elif ct.rstrip().endswith((',', ':')):
                    is_fragment = True
                elif ct[0] in ')]}' and len(ct) > 3:
                    is_fragment = True
                if is_fragment:
                    fragment_count += 1
            # 첫 비빈 셀이 소문자로 시작하면 거의 확실히 본문
            first_cell = header_cells[0] if header_cells else ''
            if first_cell and first_cell[0].islower():
                return None
            # 첫 셀이 섹션 번호(N.N.)로 시작하면 본문
            if first_cell and _re2.match(r'^\d+\.\d+', first_cell):
                return None
            if len(header_cells) > 0 and fragment_count >= len(header_cells) * 0.5:
                return None
    
    if debug:
        print(f"[Align] valid: {num_rows}x{num_cols} fill={fill_rate:.0%}")
    
    return table
