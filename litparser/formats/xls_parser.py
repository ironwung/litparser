"""
XLS Parser - Microsoft Excel 97-2003 (.xls)

구조:
  OLE2 컨테이너
  ├── Workbook: 메인 스트림 (BIFF8)
  └── SummaryInformation: 메타데이터

BIFF8 (Binary Interchange File Format) 레코드:
  - BOF: 시작
  - BOUNDSHEET: 시트 정보
  - SST: 공유 문자열 테이블
  - LABELSST: 문자열 셀 (SST 참조)
  - NUMBER: 숫자 셀
  - RK: 압축 숫자 셀
  - EOF: 끝
"""

import struct
from dataclasses import dataclass, field
from typing import List, Optional, Union, Tuple, Any
from pathlib import Path

from ..core.ole_parser import OLE2Reader, is_ole2_file


# BIFF8 레코드 타입
BIFF_BOF = 0x0809
BIFF_EOF = 0x000A
BIFF_BOUNDSHEET = 0x0085
BIFF_SST = 0x00FC         # Shared String Table
BIFF_LABELSST = 0x00FD    # 문자열 셀 (SST 참조)
BIFF_NUMBER = 0x0203      # 숫자 셀
BIFF_RK = 0x027E          # 압축 숫자 셀
BIFF_MULRK = 0x00BD       # 다중 RK
BIFF_LABEL = 0x0204       # 문자열 셀 (직접)
BIFF_BLANK = 0x0201       # 빈 셀
BIFF_BOOLERR = 0x0205     # 불리언/에러
BIFF_FORMULA = 0x0006     # 수식
BIFF_CONTINUE = 0x003C    # 연속 레코드


@dataclass
class XlsCell:
    """셀"""
    row: int
    col: int
    value: Any
    
    @property
    def address(self) -> str:
        return f"{_col_to_letter(self.col)}{self.row + 1}"


@dataclass
class XlsSheet:
    """시트"""
    name: str
    index: int
    cells: dict = field(default_factory=dict)  # (row, col) -> XlsCell
    
    @property
    def rows(self) -> int:
        if not self.cells:
            return 0
        return max(c.row for c in self.cells.values()) + 1
    
    @property
    def cols(self) -> int:
        if not self.cells:
            return 0
        return max(c.col for c in self.cells.values()) + 1
    
    def get_value(self, row: int, col: int) -> Any:
        cell = self.cells.get((row, col))
        return cell.value if cell else ""
    
    def to_list(self) -> List[List[Any]]:
        """2D 리스트로 변환"""
        if not self.cells:
            return []
        
        result = []
        for r in range(self.rows):
            row_data = []
            for c in range(self.cols):
                row_data.append(self.get_value(r, c))
            result.append(row_data)
        
        return result
    
    def to_markdown(self) -> str:
        """마크다운 테이블로 변환"""
        data = self.to_list()
        if not data:
            return ""
        
        lines = []
        header = data[0]
        lines.append("| " + " | ".join(str(c) for c in header) + " |")
        lines.append("| " + " | ".join("---" for _ in header) + " |")
        
        for row in data[1:]:
            lines.append("| " + " | ".join(str(c) for c in row) + " |")
        
        return "\n".join(lines)
    
    def get_text(self) -> str:
        """탭 구분 텍스트"""
        data = self.to_list()
        lines = []
        for row in data:
            lines.append("\t".join(str(c) for c in row))
        return "\n".join(lines)


@dataclass
class XlsDocument:
    """XLS 문서"""
    sheets: List[XlsSheet] = field(default_factory=list)
    
    title: str = ""
    author: str = ""
    created: str = ""
    
    @property
    def sheet_count(self) -> int:
        return len(self.sheets)
    
    def get_sheet(self, name_or_index: Union[str, int]) -> Optional[XlsSheet]:
        if isinstance(name_or_index, int):
            if 0 <= name_or_index < len(self.sheets):
                return self.sheets[name_or_index]
        else:
            for sheet in self.sheets:
                if sheet.name == name_or_index:
                    return sheet
        return None
    
    def get_text(self) -> str:
        """전체 텍스트"""
        parts = []
        for sheet in self.sheets:
            parts.append(f"=== {sheet.name} ===")
            parts.append(sheet.get_text())
            parts.append("")
        return "\n".join(parts)


def parse_xls(filepath_or_bytes: Union[str, bytes]) -> XlsDocument:
    """
    XLS 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        XlsDocument: 파싱된 문서
    """
    if isinstance(filepath_or_bytes, (str, Path)):
        with open(filepath_or_bytes, 'rb') as f:
            data = f.read()
    else:
        data = filepath_or_bytes
    
    if not is_ole2_file(data):
        raise ValueError("유효한 XLS 파일이 아닙니다")
    
    ole = OLE2Reader(data)
    doc = XlsDocument()
    
    # Workbook 스트림
    workbook = ole.get_stream("Workbook")
    if not workbook:
        workbook = ole.get_stream("Book")  # Excel 5/95
    
    if not workbook:
        raise ValueError("Workbook 스트림을 찾을 수 없습니다")
    
    # BIFF 레코드 파싱
    sst = []  # Shared String Table
    sheet_info = []  # (name, offset)
    
    pos = 0
    size = len(workbook)
    
    while pos + 4 <= size:
        rec_type = struct.unpack('<H', workbook[pos:pos+2])[0]
        rec_len = struct.unpack('<H', workbook[pos+2:pos+4])[0]
        pos += 4
        
        if pos + rec_len > size:
            break
        
        rec_data = workbook[pos:pos+rec_len]
        
        # 시트 정보
        if rec_type == BIFF_BOUNDSHEET:
            offset = struct.unpack('<I', rec_data[0:4])[0]
            flags = rec_data[4]
            name_len = rec_data[6]
            
            # 이름 인코딩
            if rec_data[7] == 0:
                name = rec_data[8:8+name_len].decode('latin-1', errors='ignore')
            else:
                name = rec_data[8:8+name_len*2].decode('utf-16le', errors='ignore')
            
            sheet_info.append((name, offset))
        
        # 공유 문자열 테이블
        elif rec_type == BIFF_SST:
            sst = _parse_sst(rec_data, workbook, pos + rec_len)
        
        pos += rec_len
    
    # 각 시트 파싱
    for idx, (name, offset) in enumerate(sheet_info):
        sheet = XlsSheet(name=name, index=idx)
        _parse_sheet(workbook, offset, sst, sheet)
        doc.sheets.append(sheet)
    
    return doc


def _parse_sst(data: bytes, workbook: bytes, continue_pos: int) -> List[str]:
    """Shared String Table 파싱"""
    strings = []
    
    if len(data) < 8:
        return strings
    
    total_strings = struct.unpack('<I', data[0:4])[0]
    unique_strings = struct.unpack('<I', data[4:8])[0]
    
    pos = 8
    full_data = data
    
    # CONTINUE 레코드 처리
    curr_pos = continue_pos
    while curr_pos + 4 <= len(workbook):
        rec_type = struct.unpack('<H', workbook[curr_pos:curr_pos+2])[0]
        rec_len = struct.unpack('<H', workbook[curr_pos+2:curr_pos+4])[0]
        
        if rec_type != BIFF_CONTINUE:
            break
        
        full_data += workbook[curr_pos+4:curr_pos+4+rec_len]
        curr_pos += 4 + rec_len
    
    # 문자열 파싱
    while len(strings) < unique_strings and pos < len(full_data):
        try:
            s, pos = _read_unicode_string(full_data, pos)
            strings.append(s)
        except:
            break
    
    return strings


def _read_unicode_string(data: bytes, pos: int) -> Tuple[str, int]:
    """유니코드 문자열 읽기"""
    if pos + 3 > len(data):
        return "", pos + 1
    
    str_len = struct.unpack('<H', data[pos:pos+2])[0]
    flags = data[pos + 2]
    pos += 3
    
    is_unicode = flags & 0x01
    has_ext = flags & 0x04
    has_rich = flags & 0x08
    
    # 확장 정보 스킵
    if has_rich:
        rich_count = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2
    
    if has_ext:
        ext_size = struct.unpack('<I', data[pos:pos+4])[0]
        pos += 4
    
    # 문자열 읽기
    if is_unicode:
        byte_len = str_len * 2
        text = data[pos:pos+byte_len].decode('utf-16le', errors='ignore')
        pos += byte_len
    else:
        text = data[pos:pos+str_len].decode('latin-1', errors='ignore')
        pos += str_len
    
    # 확장 데이터 스킵
    if has_rich:
        pos += rich_count * 4
    if has_ext:
        pos += ext_size
    
    return text, pos


def _parse_sheet(workbook: bytes, offset: int, sst: List[str], sheet: XlsSheet):
    """시트 데이터 파싱"""
    pos = offset
    size = len(workbook)
    
    while pos + 4 <= size:
        rec_type = struct.unpack('<H', workbook[pos:pos+2])[0]
        rec_len = struct.unpack('<H', workbook[pos+2:pos+4])[0]
        pos += 4
        
        if pos + rec_len > size:
            break
        
        rec_data = workbook[pos:pos+rec_len]
        
        if rec_type == BIFF_EOF:
            break
        
        # 문자열 셀 (SST 참조)
        elif rec_type == BIFF_LABELSST:
            if len(rec_data) >= 10:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                sst_idx = struct.unpack('<I', rec_data[6:10])[0]
                
                if sst_idx < len(sst):
                    cell = XlsCell(row=row, col=col, value=sst[sst_idx])
                    sheet.cells[(row, col)] = cell
        
        # 숫자 셀
        elif rec_type == BIFF_NUMBER:
            if len(rec_data) >= 14:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                value = struct.unpack('<d', rec_data[6:14])[0]
                
                cell = XlsCell(row=row, col=col, value=value)
                sheet.cells[(row, col)] = cell
        
        # RK 셀 (압축 숫자)
        elif rec_type == BIFF_RK:
            if len(rec_data) >= 10:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                rk = struct.unpack('<I', rec_data[6:10])[0]
                value = _decode_rk(rk)
                
                cell = XlsCell(row=row, col=col, value=value)
                sheet.cells[(row, col)] = cell
        
        # 다중 RK
        elif rec_type == BIFF_MULRK:
            if len(rec_data) >= 6:
                row = struct.unpack('<H', rec_data[0:2])[0]
                first_col = struct.unpack('<H', rec_data[2:4])[0]
                
                # 각 RK 값 (6 bytes씩: xf(2) + rk(4))
                rk_pos = 4
                col = first_col
                while rk_pos + 6 <= len(rec_data) - 2:
                    rk = struct.unpack('<I', rec_data[rk_pos+2:rk_pos+6])[0]
                    value = _decode_rk(rk)
                    
                    cell = XlsCell(row=row, col=col, value=value)
                    sheet.cells[(row, col)] = cell
                    
                    col += 1
                    rk_pos += 6
        
        # 문자열 셀 (직접)
        elif rec_type == BIFF_LABEL:
            if len(rec_data) >= 8:
                row = struct.unpack('<H', rec_data[0:2])[0]
                col = struct.unpack('<H', rec_data[2:4])[0]
                str_len = struct.unpack('<H', rec_data[6:8])[0]
                
                if len(rec_data) >= 9:
                    flags = rec_data[8]
                    if flags & 0x01:
                        text = rec_data[9:9+str_len*2].decode('utf-16le', errors='ignore')
                    else:
                        text = rec_data[9:9+str_len].decode('latin-1', errors='ignore')
                    
                    cell = XlsCell(row=row, col=col, value=text)
                    sheet.cells[(row, col)] = cell
        
        pos += rec_len


def _decode_rk(rk: int) -> float:
    """RK 값 디코딩"""
    is_int = rk & 0x02
    div_100 = rk & 0x01
    
    if is_int:
        value = (rk >> 2)
        if rk & 0x80000000:  # 음수
            value = value - 0x40000000
    else:
        # IEEE 754
        rk_bytes = struct.pack('<I', rk & 0xFFFFFFFC) + b'\x00\x00\x00\x00'
        value = struct.unpack('<d', rk_bytes)[0]
    
    if div_100:
        value /= 100
    
    return value


def _col_to_letter(col: int) -> str:
    """열 번호를 문자로 (0 -> A)"""
    result = ""
    col += 1
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result
