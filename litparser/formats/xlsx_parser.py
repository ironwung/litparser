"""
XLSX Parser - 순수 Python Excel 파서

ZIP + XML 기반 OOXML 파싱
외부 라이브러리 없이 구현

구조:
    xlsx
    ├── [Content_Types].xml
    ├── xl/
    │   ├── workbook.xml        # 워크북 정보
    │   ├── sharedStrings.xml   # 공유 문자열
    │   ├── styles.xml          # 스타일
    │   └── worksheets/
    │       ├── sheet1.xml
    │       └── sheet2.xml
    └── docProps/
        └── core.xml            # 메타데이터
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any, Union
from pathlib import Path
from io import BytesIO


# 네임스페이스
NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
}


@dataclass
class Cell:
    """셀 데이터"""
    row: int
    col: int
    value: Any
    formula: str = ""
    cell_type: str = "string"  # string, number, boolean, date, formula
    
    @property
    def address(self) -> str:
        """A1 형식 주소"""
        return f"{_col_to_letter(self.col)}{self.row}"


@dataclass
class Sheet:
    """워크시트"""
    name: str
    index: int
    cells: Dict[Tuple[int, int], Cell] = field(default_factory=dict)
    
    @property
    def rows(self) -> int:
        if not self.cells:
            return 0
        return max(c.row for c in self.cells.values())
    
    @property
    def cols(self) -> int:
        if not self.cells:
            return 0
        return max(c.col for c in self.cells.values())
    
    def get_cell(self, row: int, col: int) -> Optional[Cell]:
        return self.cells.get((row, col))
    
    def get_value(self, row: int, col: int) -> Any:
        cell = self.get_cell(row, col)
        return cell.value if cell else None
    
    def to_list(self) -> List[List[Any]]:
        """2D 리스트로 변환"""
        if not self.cells:
            return []
        
        max_row = self.rows
        max_col = self.cols
        
        result = []
        for r in range(1, max_row + 1):
            row_data = []
            for c in range(1, max_col + 1):
                cell = self.get_cell(r, c)
                row_data.append(cell.value if cell else "")
            result.append(row_data)
        
        return result
    
    def to_markdown(self) -> str:
        """마크다운 테이블로 변환"""
        data = self.to_list()
        if not data:
            return ""
        
        lines = []
        
        # 헤더
        header = data[0]
        lines.append("| " + " | ".join(str(c) for c in header) + " |")
        lines.append("| " + " | ".join("---" for _ in header) + " |")
        
        # 데이터
        for row in data[1:]:
            lines.append("| " + " | ".join(str(c) for c in row) + " |")
        
        return "\n".join(lines)
    
    def get_text(self) -> str:
        """텍스트로 변환 (탭 구분)"""
        data = self.to_list()
        lines = []
        for row in data:
            lines.append("\t".join(str(c) for c in row))
        return "\n".join(lines)


@dataclass
class XlsxDocument:
    """XLSX 문서"""
    sheets: List[Sheet] = field(default_factory=list)
    title: str = ""
    author: str = ""
    created: str = ""
    
    @property
    def sheet_count(self) -> int:
        return len(self.sheets)
    
    def get_sheet(self, name_or_index: Union[str, int]) -> Optional[Sheet]:
        """시트 가져오기"""
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


def parse_xlsx(filepath_or_bytes: Union[str, bytes]) -> XlsxDocument:
    """
    XLSX 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        XlsxDocument: 파싱된 문서
    """
    if isinstance(filepath_or_bytes, str):
        with open(filepath_or_bytes, 'rb') as f:
            data = f.read()
    else:
        data = filepath_or_bytes
    
    doc = XlsxDocument()
    
    try:
        zf = zipfile.ZipFile(BytesIO(data), 'r')
    except zipfile.BadZipFile:
        raise ValueError("유효한 XLSX 파일이 아닙니다")
    
    # 메타데이터
    if 'docProps/core.xml' in zf.namelist():
        doc.title, doc.author, doc.created = _parse_core_xml(zf)
    
    # 공유 문자열
    shared_strings = []
    if 'xl/sharedStrings.xml' in zf.namelist():
        shared_strings = _parse_shared_strings(zf)
    
    # 워크북 - 시트 이름
    sheet_names = _parse_workbook(zf)
    
    # 각 시트 파싱
    for idx, name in enumerate(sheet_names):
        sheet_path = f'xl/worksheets/sheet{idx + 1}.xml'
        if sheet_path in zf.namelist():
            sheet = _parse_sheet(zf, sheet_path, name, idx, shared_strings)
            doc.sheets.append(sheet)
    
    zf.close()
    return doc


def _parse_core_xml(zf: zipfile.ZipFile) -> Tuple[str, str, str]:
    """메타데이터 파싱"""
    title, author, created = "", "", ""
    
    try:
        content = zf.read('docProps/core.xml').decode('utf-8')
        root = ET.fromstring(content)
        
        # 제목
        el = root.find('dc:title', NS)
        if el is not None and el.text:
            title = el.text
        
        # 작성자
        el = root.find('dc:creator', NS)
        if el is not None and el.text:
            author = el.text
        
        # 생성일
        el = root.find('dcterms:created', NS)
        if el is not None and el.text:
            created = el.text
    except:
        pass
    
    return title, author, created


def _parse_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    """공유 문자열 파싱"""
    strings = []
    
    try:
        content = zf.read('xl/sharedStrings.xml').decode('utf-8')
        root = ET.fromstring(content)
        
        for si in root.findall('main:si', NS):
            text_parts = []
            
            # 단순 텍스트
            t = si.find('main:t', NS)
            if t is not None and t.text:
                text_parts.append(t.text)
            
            # 리치 텍스트
            for r in si.findall('main:r', NS):
                t = r.find('main:t', NS)
                if t is not None and t.text:
                    text_parts.append(t.text)
            
            strings.append("".join(text_parts))
    except:
        pass
    
    return strings


def _parse_workbook(zf: zipfile.ZipFile) -> List[str]:
    """워크북에서 시트 이름 추출"""
    names = []
    
    try:
        content = zf.read('xl/workbook.xml').decode('utf-8')
        root = ET.fromstring(content)
        
        sheets = root.find('main:sheets', NS)
        if sheets is not None:
            for sheet in sheets.findall('main:sheet', NS):
                name = sheet.get('name', f'Sheet{len(names) + 1}')
                names.append(name)
    except:
        pass
    
    # 시트가 없으면 파일에서 직접 찾기
    if not names:
        for name in zf.namelist():
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                idx = name.replace('xl/worksheets/sheet', '').replace('.xml', '')
                try:
                    names.append(f'Sheet{int(idx)}')
                except:
                    pass
        names.sort()
    
    return names


def _parse_sheet(zf: zipfile.ZipFile, path: str, name: str, 
                 index: int, shared_strings: List[str]) -> Sheet:
    """워크시트 파싱"""
    sheet = Sheet(name=name, index=index)
    
    try:
        content = zf.read(path).decode('utf-8')
        root = ET.fromstring(content)
        
        # sheetData
        sheet_data = root.find('main:sheetData', NS)
        if sheet_data is None:
            return sheet
        
        for row_el in sheet_data.findall('main:row', NS):
            row_num = int(row_el.get('r', 0))
            
            for cell_el in row_el.findall('main:c', NS):
                cell = _parse_cell(cell_el, row_num, shared_strings)
                if cell:
                    sheet.cells[(cell.row, cell.col)] = cell
    except:
        pass
    
    return sheet


def _parse_cell(cell_el: ET.Element, row_num: int, 
                shared_strings: List[str]) -> Optional[Cell]:
    """셀 파싱"""
    ref = cell_el.get('r', '')
    if not ref:
        return None
    
    row, col = _parse_cell_ref(ref)
    if row == 0:
        row = row_num
    
    cell_type = cell_el.get('t', '')  # s=shared string, n=number, b=boolean, str=string
    
    # 값
    value = None
    formula = ""
    
    # 수식
    f_el = cell_el.find('main:f', NS)
    if f_el is not None and f_el.text:
        formula = f_el.text
    
    # 값
    v_el = cell_el.find('main:v', NS)
    if v_el is not None and v_el.text:
        raw_value = v_el.text
        
        if cell_type == 's':
            # 공유 문자열 참조
            try:
                idx = int(raw_value)
                if 0 <= idx < len(shared_strings):
                    value = shared_strings[idx]
                else:
                    value = raw_value
            except:
                value = raw_value
        elif cell_type == 'b':
            # 불리언
            value = raw_value == '1'
        elif cell_type in ['n', '']:
            # 숫자
            try:
                if '.' in raw_value:
                    value = float(raw_value)
                else:
                    value = int(raw_value)
            except:
                value = raw_value
        else:
            value = raw_value
    
    # 인라인 문자열
    is_el = cell_el.find('main:is', NS)
    if is_el is not None:
        t_el = is_el.find('main:t', NS)
        if t_el is not None and t_el.text:
            value = t_el.text
    
    if value is None:
        value = ""
    
    return Cell(
        row=row,
        col=col,
        value=value,
        formula=formula,
        cell_type=cell_type or 'string'
    )


def _parse_cell_ref(ref: str) -> Tuple[int, int]:
    """A1 형식 셀 참조 파싱"""
    col_str = ""
    row_str = ""
    
    for c in ref:
        if c.isalpha():
            col_str += c.upper()
        elif c.isdigit():
            row_str += c
    
    col = _letter_to_col(col_str) if col_str else 0
    row = int(row_str) if row_str else 0
    
    return row, col


def _letter_to_col(letters: str) -> int:
    """A -> 1, B -> 2, ..., Z -> 26, AA -> 27"""
    result = 0
    for c in letters:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result


def _col_to_letter(col: int) -> str:
    """1 -> A, 2 -> B, ..., 27 -> AA"""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result
