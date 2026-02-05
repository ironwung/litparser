"""
PDF StructTree Parser - Tagged PDF 구조 파싱

Tagged PDF에서 문서 구조(테이블, 헤딩, 리스트 등)를 추출

지원하는 구조 타입:
- Document: 문서 루트
- H1-H6: 헤딩
- P: 단락
- L, LI: 리스트
- Table, TR, TH, TD: 테이블
- Figure: 이미지/그림
- Link: 링크
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Tuple
from enum import Enum


class StructType(Enum):
    """구조 요소 타입"""
    DOCUMENT = "Document"
    PART = "Part"
    SECT = "Sect"
    DIV = "Div"
    
    # 헤딩
    H1 = "H1"
    H2 = "H2"
    H3 = "H3"
    H4 = "H4"
    H5 = "H5"
    H6 = "H6"
    
    # 블록
    P = "P"
    L = "L"          # List
    LI = "LI"        # List Item
    
    # 테이블
    TABLE = "Table"
    TR = "TR"        # Table Row
    TH = "TH"        # Table Header Cell
    TD = "TD"        # Table Data Cell
    THEAD = "THead"
    TBODY = "TBody"
    
    # 기타
    FIGURE = "Figure"
    LINK = "Link"
    SPAN = "Span"
    NONSTRUCT = "NonStruct"
    
    UNKNOWN = "Unknown"
    
    @classmethod
    def from_string(cls, s: str) -> 'StructType':
        for t in cls:
            if t.value == s:
                return t
        return cls.UNKNOWN


@dataclass
class StructElement:
    """구조 요소"""
    type: StructType
    children: List['StructElement'] = field(default_factory=list)
    text: str = ""
    mcid: Optional[int] = None  # Marked Content ID
    page: Optional[int] = None
    attributes: Dict[str, Any] = field(default_factory=dict)
    
    def get_text(self) -> str:
        """모든 하위 텍스트 수집"""
        if self.text:
            return self.text
        
        texts = []
        for child in self.children:
            child_text = child.get_text()
            if child_text:
                texts.append(child_text)
        
        return ' '.join(texts)


@dataclass
class StructTable:
    """추출된 테이블"""
    rows: List[List[str]] = field(default_factory=list)
    headers: List[str] = field(default_factory=list)
    num_rows: int = 0
    num_cols: int = 0
    
    def to_markdown(self) -> str:
        """마크다운 테이블로 변환"""
        if not self.rows and not self.headers:
            return ""
        
        lines = []
        
        # 헤더
        if self.headers:
            lines.append('| ' + ' | '.join(self.headers) + ' |')
            lines.append('| ' + ' | '.join(['---'] * len(self.headers)) + ' |')
            data_rows = self.rows
        elif self.rows:
            lines.append('| ' + ' | '.join(str(c) for c in self.rows[0]) + ' |')
            lines.append('| ' + ' | '.join(['---'] * len(self.rows[0])) + ' |')
            data_rows = self.rows[1:]
        else:
            return ""
        
        for row in data_rows:
            lines.append('| ' + ' | '.join(str(c) for c in row) + ' |')
        
        return '\n'.join(lines)
    
    def to_list(self) -> List[List[str]]:
        """2D 리스트로 반환"""
        if self.headers:
            return [self.headers] + self.rows
        return self.rows


class StructTreeParser:
    """StructTree 파서"""
    
    def __init__(self, doc):
        """
        Args:
            doc: PDFDocument 객체
        """
        self.doc = doc
        self.mcid_to_text: Dict[Tuple[int, int], str] = {}  # (page, mcid) -> text
        self._page_contents_parsed = set()
    
    def _resolve(self, ref) -> Any:
        """참조 해결"""
        from .parser import PDFRef
        if isinstance(ref, PDFRef):
            return self.doc.objects.get((ref.obj_num, ref.gen_num))
        return ref
    
    def is_tagged(self) -> bool:
        """Tagged PDF인지 확인"""
        root_ref = self.doc.trailer.get('Root')
        catalog = self._resolve(root_ref)
        
        if not catalog:
            return False
        
        mark_info = catalog.get('MarkInfo', {})
        if isinstance(mark_info, dict) and mark_info.get('Marked'):
            return True
        
        struct_tree_ref = catalog.get('StructTreeRoot')
        return struct_tree_ref is not None
    
    def get_struct_tree_root(self) -> Optional[Dict]:
        """StructTreeRoot 객체 반환"""
        root_ref = self.doc.trailer.get('Root')
        catalog = self._resolve(root_ref)
        
        if not catalog:
            return None
        
        struct_tree_ref = catalog.get('StructTreeRoot')
        return self._resolve(struct_tree_ref)
    
    def _build_mcid_map_for_page(self, page_num: int):
        """특정 페이지의 MCID -> 텍스트 매핑 구축"""
        if page_num in self._page_contents_parsed:
            return
        
        from .parser import PDFRef
        from .. import get_pages, decode_stream, extract_text_with_positions
        
        pages = get_pages(self.doc)
        if page_num >= len(pages):
            return
        
        page = pages[page_num]
        
        # Content Stream 가져오기
        contents_ref = page.get('Contents')
        content_data = b''
        
        if isinstance(contents_ref, PDFRef):
            contents_obj = self.doc.objects.get((contents_ref.obj_num, contents_ref.gen_num))
            if contents_obj:
                content_data = decode_stream(self.doc, contents_obj)
        elif isinstance(contents_ref, list):
            for ref in contents_ref:
                if isinstance(ref, PDFRef):
                    obj = self.doc.objects.get((ref.obj_num, ref.gen_num))
                    if obj:
                        content_data += decode_stream(self.doc, obj) + b'\n'
        
        if not content_data:
            self._page_contents_parsed.add(page_num)
            return
        
        # 기존 텍스트 추출 사용
        text_items = extract_text_with_positions(self.doc, page_num)
        
        # MCID 범위 파싱 (BDC ~ EMC 사이의 위치)
        self._parse_mcid_ranges(content_data, page_num, text_items)
        self._page_contents_parsed.add(page_num)
    
    def _parse_mcid_ranges(self, content: bytes, page_num: int, text_items: list):
        """MCID 범위와 텍스트 아이템 매칭"""
        import re
        
        try:
            text = content.decode('latin-1', errors='replace')
        except:
            return
        
        # BDC ~ EMC 범위 파싱
        mcid_ranges = []  # [(mcid, start_pos, end_pos)]
        
        bdc_pattern = r'/\w+\s*<<[^>]*?/MCID\s+(\d+)[^>]*>>\s*BDC'
        
        for match in re.finditer(bdc_pattern, text):
            mcid = int(match.group(1))
            start_pos = match.end()
            
            # 매칭되는 EMC 찾기
            emc_pos = text.find('EMC', start_pos)
            if emc_pos > 0:
                mcid_ranges.append((mcid, start_pos, emc_pos))
        
        # 각 MCID 범위 내의 Td/Tm 위치 추적하여 텍스트 매칭
        # 간단한 휴리스틱: MCID 순서와 텍스트 아이템 Y 좌표 기반 매칭
        
        # MCID별 텍스트 수집 (Content Stream 내 Tj/TJ에서 직접 추출하는 대신
        # 이미 추출된 text_items를 순차적으로 할당)
        
        if not mcid_ranges or not text_items:
            return
        
        # 텍스트 아이템을 Y, X 순으로 정렬
        sorted_items = sorted(text_items, key=lambda t: (t.y, t.x))
        
        # MCID 범위 수와 텍스트 그룹 수가 대략 맞아야 의미 있음
        # 간단히 Y 좌표 변화를 기준으로 그룹화
        groups = []
        current_group = []
        prev_y = None
        
        for item in sorted_items:
            if prev_y is not None and abs(item.y - prev_y) > 15:
                if current_group:
                    groups.append(current_group)
                current_group = [item]
            else:
                current_group.append(item)
            prev_y = item.y
        
        if current_group:
            groups.append(current_group)
        
        # MCID와 그룹 매칭 (순서 기반)
        for i, (mcid, _, _) in enumerate(mcid_ranges):
            if i < len(groups):
                group_text = ''.join(item.text for item in groups[i])
                self.mcid_to_text[(page_num, mcid)] = group_text
    
    def _parse_marked_content(self, content: bytes, page_num: int):
        """Content Stream에서 Marked Content 파싱"""
        import re
        
        try:
            text = content.decode('latin-1', errors='replace')
        except:
            return
        
        # BDC (Begin Marked Content) 패턴
        # /Tag <<...MCID X...>> BDC ... EMC
        bdc_pattern = r'/(\w+)\s*<<([^>]*)>>\s*BDC'
        
        current_pos = 0
        current_mcid = None
        
        while current_pos < len(text):
            # BDC 찾기
            bdc_match = re.search(bdc_pattern, text[current_pos:])
            if not bdc_match:
                break
            
            bdc_start = current_pos + bdc_match.end()
            props = bdc_match.group(2)
            
            # MCID 추출
            mcid_match = re.search(r'/MCID\s+(\d+)', props)
            if mcid_match:
                current_mcid = int(mcid_match.group(1))
            
            # EMC 찾기
            emc_pos = text.find('EMC', bdc_start)
            if emc_pos == -1:
                break
            
            # BDC와 EMC 사이의 텍스트 추출
            if current_mcid is not None:
                content_between = text[bdc_start:emc_pos]
                extracted_text = self._extract_text_from_content(content_between)
                if extracted_text:
                    key = (page_num, current_mcid)
                    self.mcid_to_text[key] = extracted_text
            
            current_pos = emc_pos + 3
            current_mcid = None
    
    def _extract_text_from_content(self, content: str) -> str:
        """Content Stream 조각에서 텍스트 추출"""
        import re
        
        texts = []
        
        # Tj 명령어 (단일 문자열)
        tj_pattern = r'\(([^)]*)\)\s*Tj'
        for match in re.finditer(tj_pattern, content):
            texts.append(match.group(1))
        
        # TJ 명령어 (배열)
        tj_array_pattern = r'\[([^\]]+)\]\s*TJ'
        for match in re.finditer(tj_array_pattern, content):
            array_content = match.group(1)
            # 문자열 부분만 추출
            for str_match in re.finditer(r'\(([^)]*)\)', array_content):
                texts.append(str_match.group(1))
        
        return ''.join(texts)
    
    def parse(self) -> Optional[StructElement]:
        """전체 구조 트리 파싱"""
        struct_tree = self.get_struct_tree_root()
        
        if not struct_tree:
            return None
        
        # 모든 페이지의 MCID 맵 구축
        from .. import get_page_count
        page_count = get_page_count(self.doc)
        for i in range(page_count):
            self._build_mcid_map_for_page(i)
        
        # K (Kids) 파싱
        k = struct_tree.get('K')
        return self._parse_element(k)
    
    def _parse_element(self, node, depth=0, page_num=0) -> Optional[StructElement]:
        """구조 요소 재귀 파싱"""
        if depth > 50:  # 무한 재귀 방지
            return None
        
        node = self._resolve(node)
        
        if node is None:
            return None
        
        # MCID (정수)인 경우
        if isinstance(node, int):
            text = self.mcid_to_text.get((page_num, node), "")
            return StructElement(
                type=StructType.UNKNOWN,
                mcid=node,
                text=text,
                page=page_num
            )
        
        if not isinstance(node, dict):
            return None
        
        # 페이지 번호 추출
        pg = node.get('Pg')
        if pg:
            pg_obj = self._resolve(pg)
            # 페이지 객체에서 번호 찾기 (간접적)
            # 일단은 현재 page_num 유지
        
        # 구조 타입
        s_type = node.get('S', 'Unknown')
        struct_type = StructType.from_string(s_type)
        
        # 요소 생성
        element = StructElement(
            type=struct_type,
            attributes={},
            page=page_num
        )
        
        # 속성 복사
        for key in ['Lang', 'Alt', 'ActualText', 'Title']:
            if key in node:
                element.attributes[key] = node[key]
        
        # ActualText가 있으면 텍스트로 사용
        if 'ActualText' in node:
            element.text = str(node['ActualText'])
        
        # 자식 파싱
        kids = node.get('K')
        
        if kids is None:
            pass
        elif isinstance(kids, int):
            # MCID
            element.mcid = kids
            element.text = self.mcid_to_text.get((page_num, kids), "")
        elif isinstance(kids, list):
            for kid in kids:
                child = self._parse_element(kid, depth + 1, page_num)
                if child:
                    element.children.append(child)
        else:
            # 단일 자식
            child = self._parse_element(kids, depth + 1, page_num)
            if child:
                element.children.append(child)
        
        return element
    
    def find_tables_with_text(self, page_num: int = None) -> List[StructTable]:
        """
        테이블 찾기 (텍스트 포함)
        
        StructTree에서 테이블 구조를 찾고,
        텍스트는 위치 기반으로 매핑
        """
        from .. import extract_text_with_positions, get_page_count
        
        root = self.parse()
        if not root:
            return []
        
        tables = []
        
        # 페이지 범위
        if page_num is not None:
            pages = [page_num]
        else:
            pages = range(get_page_count(self.doc))
        
        for pg in pages:
            # 해당 페이지의 텍스트 아이템
            text_items = extract_text_with_positions(self.doc, pg)
            if not text_items:
                continue
            
            # 텍스트 항목 합치기
            merged_items = self._merge_text_items(text_items)
            
            # 이 페이지의 테이블 구조 찾기 (벡터 기반)
            page_tables = self._find_tables_from_vectors_on_page(pg, merged_items)
            tables.extend(page_tables)
        
        return tables
    
    def _merge_text_items(self, items: list) -> list:
        """인접한 텍스트 항목 합치기"""
        if not items:
            return []
        
        from collections import defaultdict
        
        # Y 좌표로 그룹화
        lines = defaultdict(list)
        for item in items:
            line_y = round(item.y / 5) * 5
            lines[line_y].append(item)
        
        merged = []
        
        for line_y in sorted(lines.keys()):
            line_items = sorted(lines[line_y], key=lambda t: t.x)
            
            if not line_items:
                continue
            
            # 같은 줄 합치기
            current_text = line_items[0].text
            current_x = line_items[0].x
            current_size = getattr(line_items[0], 'font_size', 12)
            prev_x = line_items[0].x
            
            for item in line_items[1:]:
                item_size = getattr(item, 'font_size', 12)
                gap = item.x - prev_x - current_size * 0.5
                
                if gap > current_size * 0.8:
                    # 새 단어
                    if current_text.strip():
                        merged.append({
                            'text': current_text,
                            'x': current_x,
                            'y': line_y,
                            'font_size': current_size
                        })
                    current_text = item.text
                    current_x = item.x
                    current_size = item_size
                else:
                    current_text += item.text
                
                prev_x = item.x
            
            if current_text.strip():
                merged.append({
                    'text': current_text,
                    'x': current_x,
                    'y': line_y,
                    'font_size': current_size
                })
        
        return merged
    
    def _find_tables_from_vectors_on_page(self, page_num: int, text_items: list) -> List[StructTable]:
        """페이지에서 벡터 기반 테이블 찾기"""
        from .parser import PDFRef
        from .. import get_pages, decode_stream
        
        pages = get_pages(self.doc)
        if page_num >= len(pages):
            return []
        
        page = pages[page_num]
        
        # Content Stream 가져오기
        contents_ref = page.get('Contents')
        content_data = b''
        
        if isinstance(contents_ref, PDFRef):
            contents_obj = self.doc.objects.get((contents_ref.obj_num, contents_ref.gen_num))
            if contents_obj:
                content_data = decode_stream(self.doc, contents_obj)
        elif isinstance(contents_ref, list):
            for ref in contents_ref:
                if isinstance(ref, PDFRef):
                    obj = self.doc.objects.get((ref.obj_num, ref.gen_num))
                    if obj:
                        content_data += decode_stream(self.doc, obj) + b'\n'
        
        if not content_data:
            return []
        
        # 테이블 태그가 있는지 확인
        try:
            content_str = content_data.decode('latin-1', errors='replace')
        except:
            return []
        
        if '/Table' not in content_str and '/TH' not in content_str:
            return []
        
        # 사각형 추출
        import re
        rect_pattern = r'([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+re'
        rectangles = []
        
        for match in re.finditer(rect_pattern, content_str):
            try:
                x = float(match.group(1))
                y = float(match.group(2))
                w = float(match.group(3))
                h = float(match.group(4))
                if abs(w) > 10 and abs(h) > 10:
                    rectangles.append({'x': x, 'y': y, 'w': abs(w), 'h': abs(h)})
            except:
                continue
        
        if len(rectangles) < 4:
            return []
        
        # 그리드 구성
        tables = self._build_tables_from_rectangles(rectangles, text_items)
        return tables
    
    def _build_tables_from_rectangles(self, rectangles: list, text_items: list) -> List[StructTable]:
        """사각형과 텍스트로 테이블 구성"""
        if not rectangles:
            return []
        
        # X, Y 좌표 클러스터링
        x_coords = sorted(set(r['x'] for r in rectangles))
        y_coords = sorted(set(r['y'] for r in rectangles))
        
        def cluster(coords, tol=10):
            if not coords:
                return []
            clusters = []
            current = [coords[0]]
            for c in coords[1:]:
                if c - current[-1] <= tol:
                    current.append(c)
                else:
                    clusters.append(sum(current) / len(current))
                    current = [c]
            if current:
                clusters.append(sum(current) / len(current))
            return clusters
        
        x_clusters = cluster(x_coords)
        y_clusters = cluster(y_coords)
        
        if len(x_clusters) < 2 or len(y_clusters) < 2:
            return []
        
        # 셀 그리드 생성
        num_cols = len(x_clusters)
        num_rows = len(y_clusters)
        
        rows = []
        for row_idx in range(num_rows):
            row = []
            cell_y = y_clusters[row_idx]
            cell_h = y_clusters[row_idx + 1] - cell_y if row_idx + 1 < num_rows else 50
            
            for col_idx in range(num_cols):
                cell_x = x_clusters[col_idx]
                cell_w = x_clusters[col_idx + 1] - cell_x if col_idx + 1 < num_cols else 100
                
                # 이 셀에 속하는 텍스트 찾기
                cell_texts = []
                for item in text_items:
                    ix, iy = item['x'], item['y']
                    # 셀 범위 내에 있는지 확인
                    if (cell_x - 20 <= ix <= cell_x + cell_w + 20 and
                        cell_y - 20 <= iy <= cell_y + cell_h + 20):
                        cell_texts.append(item['text'])
                
                row.append(' '.join(cell_texts))
            
            rows.append(row)
        
        # 빈 테이블 제외
        non_empty = sum(1 for row in rows for cell in row if cell.strip())
        if non_empty < 3:
            return []
        
        return [StructTable(
            rows=rows[1:] if len(rows) > 1 else rows,
            headers=rows[0] if rows else [],
            num_rows=len(rows) - 1 if len(rows) > 1 else len(rows),
            num_cols=num_cols
        )]
    
    def _find_tables_recursive(self, element: StructElement, tables: List[StructTable]):
        """재귀적으로 테이블 찾기"""
        if element.type == StructType.TABLE:
            table = self._extract_table(element)
            if table and (table.num_rows > 0 or table.headers):
                tables.append(table)
        
        for child in element.children:
            self._find_tables_recursive(child, tables)
    
    def find_tables(self) -> List[StructTable]:
        """모든 테이블 찾기 (하이브리드 방식)"""
        # 먼저 벡터+텍스트 기반 시도
        tables = self.find_tables_with_text()
        if tables:
            return tables
        
        # fallback: StructTree 기반
        root = self.parse()
        if not root:
            return []
        
        tables = []
        self._find_tables_recursive(root, tables)
        return tables
    
    def _extract_table(self, table_element: StructElement) -> Optional[StructTable]:
        """테이블 요소에서 데이터 추출"""
        rows = []
        headers = []
        
        # TR 직접 자식 또는 THead/TBody 안의 TR 찾기
        def collect_rows(elem):
            result = []
            for child in elem.children:
                if child.type == StructType.TR:
                    result.append(child)
                elif child.type in [StructType.THEAD, StructType.TBODY]:
                    result.extend(collect_rows(child))
            return result
        
        tr_elements = collect_rows(table_element)
        
        for tr in tr_elements:
            row_data = []
            is_header_row = False
            
            for cell in tr.children:
                if cell.type in [StructType.TH, StructType.TD]:
                    cell_text = cell.get_text().strip()
                    row_data.append(cell_text)
                    
                    if cell.type == StructType.TH:
                        is_header_row = True
            
            if row_data:
                if is_header_row and not headers:
                    headers = row_data
                else:
                    rows.append(row_data)
        
        if not rows and not headers:
            return None
        
        # 열 수 정규화
        num_cols = max(
            len(headers) if headers else 0,
            max((len(r) for r in rows), default=0)
        )
        
        if num_cols == 0:
            return None
        
        if headers:
            headers = headers + [''] * (num_cols - len(headers))
        
        normalized_rows = []
        for row in rows:
            normalized_rows.append(row + [''] * (num_cols - len(row)))
        
        return StructTable(
            rows=normalized_rows,
            headers=headers,
            num_rows=len(normalized_rows),
            num_cols=num_cols
        )
    
    def get_document_outline(self) -> List[Tuple[int, str]]:
        """문서 개요 (헤딩 목록) 추출"""
        root = self.parse()
        if not root:
            return []
        
        outline = []
        self._collect_headings(root, outline)
        return outline
    
    def _collect_headings(self, element: StructElement, outline: List[Tuple[int, str]]):
        """헤딩 수집"""
        heading_types = {
            StructType.H1: 1,
            StructType.H2: 2,
            StructType.H3: 3,
            StructType.H4: 4,
            StructType.H5: 5,
            StructType.H6: 6,
        }
        
        if element.type in heading_types:
            level = heading_types[element.type]
            text = element.get_text().strip()
            if text:
                outline.append((level, text))
        
        for child in element.children:
            self._collect_headings(child, outline)
    
    def get_structure_stats(self) -> Dict[str, int]:
        """구조 타입 통계"""
        root = self.parse()
        if not root:
            return {}
        
        stats = {}
        self._count_types(root, stats)
        return stats
    
    def _count_types(self, element: StructElement, stats: Dict[str, int]):
        """타입별 개수 세기"""
        type_name = element.type.value
        stats[type_name] = stats.get(type_name, 0) + 1
        
        for child in element.children:
            self._count_types(child, stats)


def extract_tables_from_struct_tree(doc) -> List[StructTable]:
    """
    Tagged PDF에서 테이블 추출
    
    Args:
        doc: PDFDocument 객체
    
    Returns:
        List[StructTable]: 추출된 테이블 목록
    """
    parser = StructTreeParser(doc)
    
    if not parser.is_tagged():
        return []
    
    return parser.find_tables()


def get_document_structure(doc) -> Optional[StructElement]:
    """
    문서 구조 트리 반환
    
    Args:
        doc: PDFDocument 객체
    
    Returns:
        StructElement: 문서 구조 루트 또는 None
    """
    parser = StructTreeParser(doc)
    return parser.parse()


def is_tagged_pdf(doc) -> bool:
    """Tagged PDF인지 확인"""
    parser = StructTreeParser(doc)
    return parser.is_tagged()
