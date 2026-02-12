"""
HWP Parser - 한글 5.0+ 바이너리 포맷 (.hwp)

구조: OLE2 Container -> BodyText Storage -> Section 스트림 (zlib 압축) -> HWP Tag 레코드

순수 Python 구현 - 외부 라이브러리 없음
"""

import struct
import zlib
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Union

from ..core.ole_parser import OLE2Reader, is_ole2_file


# HWP 태그 ID (HWPTAG_BEGIN = 0x010 = 16)
HWPTAG_BEGIN = 0x010
HWPTAG_PARA_HEADER = HWPTAG_BEGIN + 50   # 66
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51     # 67
HWPTAG_CTRL_HEADER = HWPTAG_BEGIN + 55   # 71
HWPTAG_LIST_HEADER = HWPTAG_BEGIN + 56   # 72
HWPTAG_TABLE = HWPTAG_BEGIN + 61         # 77

# HWP 제어 문자별 추가 데이터 길이 (바이트)
CTRL_EXTRA_BYTES = {
    0: 0,   # NULL
    1: 0,   # reserved
    2: 12,  # section/column define
    3: 12,  # field start
    4: 0,   # field end
    5: 12,  # reserved
    6: 12,  # reserved
    7: 12,  # reserved
    8: 12,  # title mark
    9: 2,   # tab
    10: 0,  # line break
    11: 12, # drawing object
    12: 12, # reserved
    13: 0,  # paragraph end
    14: 12, # reserved
    15: 12, # hidden comment
    16: 12, # header/footer
    17: 12, # footnote/endnote
    18: 12, # auto number
    19: 12, # reserved
    20: 12, # reserved
    21: 12, # page control
    22: 12, # bookmark
    23: 12, # dutmal/glossary
    24: 2,  # form object
    25: 12, # hyphen
    26: 12, # reserved
    27: 12, # reserved
    28: 12, # reserved
    29: 12, # reserved
    30: 0,  # non-break space
    31: 12, # reserved
}


@dataclass
class HwpParagraph:
    """HWP 문단"""
    text: str = ""
    is_heading: bool = False
    heading_level: int = 0


@dataclass 
class HwpTable:
    """HWP 테이블"""
    rows: List[List[str]] = field(default_factory=list)
    
    def to_markdown(self) -> str:
        if not self.rows:
            return ""
        
        max_cols = max(len(r) for r in self.rows) if self.rows else 0
        if max_cols == 0:
            return ""
        
        lines = []
        header = self.rows[0] + [''] * (max_cols - len(self.rows[0]))
        lines.append("| " + " | ".join(str(c) for c in header) + " |")
        lines.append("| " + " | ".join("---" for _ in range(max_cols)) + " |")
        for row in self.rows[1:]:
            row = row + [''] * (max_cols - len(row))
            lines.append("| " + " | ".join(str(c) for c in row) + " |")
        
        return "\n".join(lines)


@dataclass
class HwpDocument:
    """파싱된 HWP 문서"""
    paragraphs: List[HwpParagraph] = field(default_factory=list)
    tables: List[HwpTable] = field(default_factory=list)
    images: List[dict] = field(default_factory=list)
    
    title: str = ""
    author: str = ""
    created: str = ""
    
    def get_text(self) -> str:
        """전체 텍스트"""
        lines = []
        for p in self.paragraphs:
            if p.text.strip():
                if p.is_heading and p.heading_level:
                    prefix = '#' * p.heading_level + ' '
                    lines.append(prefix + p.text)
                else:
                    lines.append(p.text)
        return '\n\n'.join(lines)
    
    def get_headings(self) -> List[Tuple[int, str]]:
        """헤딩 목록"""
        return [(p.heading_level, p.text) for p in self.paragraphs if p.is_heading]


# 태그 레코드
@dataclass
class TagRecord:
    tag_id: int
    level: int
    data: bytes


def _decode_text(data: bytes) -> str:
    """HWP PARA_TEXT 디코딩 - 제어 문자 처리"""
    text_chars = []
    i = 0
    
    while i < len(data) - 1:
        char_code = struct.unpack('<H', data[i:i+2])[0]
        
        if char_code == 0:  # NULL
            i += 2
            continue
        elif char_code < 32:  # 제어 문자
            extra = CTRL_EXTRA_BYTES.get(char_code, 12)
            i += 2 + extra
            
            if char_code in (10, 13):  # line break, paragraph end
                text_chars.append('\n')
            elif char_code == 9:  # tab
                text_chars.append('\t')
        else:
            text_chars.append(chr(char_code))
            i += 2
    
    return ''.join(text_chars).strip()


def _parse_tag_records(data: bytes) -> List[TagRecord]:
    """바이너리 데이터에서 태그 레코드 목록 추출"""
    records = []
    pos = 0
    size = len(data)
    
    while pos + 4 <= size:
        header = struct.unpack('<I', data[pos:pos+4])[0]
        pos += 4
        
        tag_id = header & 0x3FF
        level = (header >> 10) & 0x3FF
        rec_len = (header >> 20) & 0xFFF
        
        if rec_len == 0xFFF:
            if pos + 4 > size:
                break
            rec_len = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
        
        if pos + rec_len > size:
            break
        
        rec_data = data[pos:pos + rec_len]
        pos += rec_len
        
        records.append(TagRecord(tag_id=tag_id, level=level, data=rec_data))
    
    return records


def _extract_tables(records: List[TagRecord]) -> List[HwpTable]:
    """태그 레코드에서 테이블 추출"""
    tables = []
    
    # TABLE 태그 위치 찾기
    table_indices = [i for i, r in enumerate(records) if r.tag_id == HWPTAG_TABLE]
    
    for ti in table_indices:
        rec = records[ti]
        tbl_level = rec.level
        
        # TABLE 레코드에서 행/열 수 추출 (offset 4, 6)
        if len(rec.data) < 8:
            continue
        n_rows = struct.unpack('<H', rec.data[4:6])[0]
        n_cols = struct.unpack('<H', rec.data[6:8])[0]
        
        if n_rows == 0 or n_cols == 0 or n_rows > 500 or n_cols > 100:
            continue
        
        # 셀 텍스트 수집: LIST_HEADER(same level) = 셀 구분자
        cell_texts = []
        current_cell = []
        
        for j in range(ti + 1, len(records)):
            r2 = records[j]
            
            # 종료 조건: 다음 TABLE 태그
            if r2.tag_id == HWPTAG_TABLE:
                break
            # 종료 조건: TABLE보다 낮은 레벨의 CTRL_HEADER (테이블 범위 벗어남)
            if r2.tag_id == HWPTAG_CTRL_HEADER and r2.level <= tbl_level - 1:
                break
            # 종료 조건: TABLE보다 낮은 레벨
            if r2.level < tbl_level:
                break
            
            # LIST_HEADER(same level) = 새 셀 시작
            if r2.tag_id == HWPTAG_LIST_HEADER and r2.level == tbl_level:
                if current_cell:
                    cell_texts.append('\n'.join(current_cell))
                current_cell = []
            # PARA_TEXT = 셀 내 텍스트
            elif r2.tag_id == HWPTAG_PARA_TEXT:
                text = _decode_text(r2.data)
                if text.strip():
                    current_cell.append(text.strip())
        
        # 마지막 셀
        if current_cell:
            cell_texts.append('\n'.join(current_cell))
        
        # 행/열로 재구성
        table = HwpTable()
        for r in range(n_rows):
            row = []
            for c in range(n_cols):
                idx = r * n_cols + c
                row.append(cell_texts[idx] if idx < len(cell_texts) else "")
            table.rows.append(row)
        
        if table.rows:
            tables.append(table)
    
    return tables


def _extract_paragraphs_and_tables(records: List[TagRecord]) -> Tuple[List[HwpParagraph], List[HwpTable]]:
    """
    태그 레코드에서 문단과 테이블을 문서 순서대로 추출
    
    테이블 내부의 PARA_TEXT는 테이블로만 추출하고,
    테이블 외부의 PARA_TEXT는 일반 문단으로 추출
    """
    paragraphs = []
    tables = []
    
    # 먼저 테이블 범위 (시작~끝 record index) 파악
    table_ranges = []  # (start_idx, end_idx, HwpTable)
    table_indices = [i for i, r in enumerate(records) if r.tag_id == HWPTAG_TABLE]
    
    for ti in table_indices:
        rec = records[ti]
        tbl_level = rec.level
        
        if len(rec.data) < 8:
            continue
        n_rows = struct.unpack('<H', rec.data[4:6])[0]
        n_cols = struct.unpack('<H', rec.data[6:8])[0]
        
        if n_rows == 0 or n_cols == 0 or n_rows > 500 or n_cols > 100:
            continue
        
        # 테이블 끝 위치와 셀 텍스트 수집
        cell_texts = []
        current_cell = []
        end_idx = ti
        
        for j in range(ti + 1, len(records)):
            r2 = records[j]
            
            if r2.tag_id == HWPTAG_TABLE:
                break
            if r2.tag_id == HWPTAG_CTRL_HEADER and r2.level <= tbl_level - 1:
                break
            if r2.level < tbl_level:
                break
            
            end_idx = j
            
            if r2.tag_id == HWPTAG_LIST_HEADER and r2.level == tbl_level:
                if current_cell:
                    cell_texts.append('\n'.join(current_cell))
                current_cell = []
            elif r2.tag_id == HWPTAG_PARA_TEXT:
                text = _decode_text(r2.data)
                if text.strip():
                    current_cell.append(text.strip())
        
        if current_cell:
            cell_texts.append('\n'.join(current_cell))
        
        # 행/열 재구성
        table = HwpTable()
        for r in range(n_rows):
            row = []
            for c in range(n_cols):
                idx = r * n_cols + c
                row.append(cell_texts[idx] if idx < len(cell_texts) else "")
            table.rows.append(row)
        
        if table.rows:
            table_ranges.append((ti, end_idx, table))
    
    # 테이블 범위에 속하는 record index 집합
    in_table = set()
    for start, end, _ in table_ranges:
        for i in range(start, end + 1):
            in_table.add(i)
    
    # 문서 순서대로 문단/테이블 출력
    added_tables = set()
    for i, rec in enumerate(records):
        # 테이블 시작이면 테이블 추가
        for start, end, table in table_ranges:
            if i == start and start not in added_tables:
                added_tables.add(start)
                tables.append(table)
                # 테이블 내용을 문단으로도 추가 (텍스트 추출용)
                for row in table.rows:
                    texts = [c.replace('\n', ' ').strip() for c in row if c.strip()]
                    if texts:
                        combined = ' '.join(texts)
                        paragraphs.append(HwpParagraph(text=combined))
                break
        
        # 테이블 범위 내 레코드는 스킵 (이미 테이블로 처리됨)
        if i in in_table:
            continue
        
        # 일반 PARA_TEXT
        if rec.tag_id == HWPTAG_PARA_TEXT and rec.data:
            text = _decode_text(rec.data)
            if text.strip():
                paragraphs.append(HwpParagraph(text=text))
    
    return paragraphs, tables


class HwpParser:
    """HWP 파일 파서"""
    
    def __init__(self, filepath_or_bytes: Union[str, bytes]):
        if isinstance(filepath_or_bytes, str):
            with open(filepath_or_bytes, 'rb') as f:
                data = f.read()
        else:
            data = filepath_or_bytes
        
        if not is_ole2_file(data):
            raise ValueError("OLE2 형식이 아님 - HWP 파일이 아닐 수 있음")
        
        self.ole = OLE2Reader(data)
        self.doc = HwpDocument()
        self.is_compressed = True
    
    def parse(self) -> HwpDocument:
        """HWP 파일 파싱"""
        if not self._validate_header():
            raise ValueError("유효한 HWP 5.0 파일이 아님")
        
        self._parse_body_text()
        self._parse_metadata()
        
        return self.doc
    
    def _validate_header(self) -> bool:
        """FileHeader 검증"""
        header = self.ole.get_stream("FileHeader")
        if not header or len(header) < 256:
            return False
        
        signature = header[:32].decode('utf-8', errors='ignore')
        if "HWP Document File" not in signature:
            return False
        
        if len(header) >= 40:
            properties = struct.unpack('<I', header[36:40])[0]
            self.is_compressed = bool(properties & 0x01)
            is_encrypted = bool(properties & 0x02)
            
            if is_encrypted:
                raise ValueError("암호화된 HWP 파일은 지원하지 않음")
        
        return True
    
    def _parse_body_text(self):
        """본문 텍스트 파싱"""
        section_idx = 0
        
        while True:
            stream_name = f"Section{section_idx}"
            stream_data = self.ole.get_stream(stream_name)
            
            if not stream_data:
                break
            
            self._parse_section(stream_data)
            section_idx += 1
        
        if section_idx == 0:
            for entry in self.ole.entries:
                if entry.name.startswith("Section") and entry.entry_type == 2:
                    stream_data = self.ole._read_stream_data(entry)
                    if stream_data:
                        self._parse_section(stream_data)
    
    def _parse_section(self, data: bytes):
        """섹션 파싱 - 문단 + 테이블"""
        if self.is_compressed:
            try:
                data = zlib.decompress(data, -15)
            except zlib.error:
                try:
                    data = zlib.decompress(data)
                except zlib.error:
                    return
        
        # 태그 레코드 파싱
        records = _parse_tag_records(data)
        
        # 문단 + 테이블 추출
        paragraphs, tables = _extract_paragraphs_and_tables(records)
        
        self.doc.paragraphs.extend(paragraphs)
        self.doc.tables.extend(tables)
    
    def _parse_metadata(self):
        """메타데이터 추출"""
        summary = self.ole.get_stream("\x05SummaryInformation")
        if summary:
            pass  # 복잡한 포맷이라 스킵


def parse_hwp(filepath_or_bytes: Union[str, bytes]) -> HwpDocument:
    """HWP 파일 파싱"""
    parser = HwpParser(filepath_or_bytes)
    return parser.parse()