"""
DOC Parser - Microsoft Word 97-2003 (.doc)

구조:
  OLE2 컨테이너
  ├── WordDocument: 메인 스트림
  ├── 1Table 또는 0Table: 테이블 스트림
  ├── Data: 데이터 스트림
  └── SummaryInformation: 메타데이터

WordDocument 스트림:
  - FIB (File Information Block): 문서 정보
  - 텍스트는 clx 구조로 위치 지정
"""

import struct
from dataclasses import dataclass, field
from typing import List, Optional, Union, Tuple
from pathlib import Path

from ..core.ole_parser import OLE2Reader, is_ole2_file


@dataclass
class DocParagraph:
    """문단"""
    text: str
    style: str = ""
    is_heading: bool = False
    heading_level: int = 0


@dataclass
class DocTable:
    """테이블"""
    rows: List[List[str]] = field(default_factory=list)
    
    def to_markdown(self) -> str:
        if not self.rows:
            return ""
        
        lines = []
        header = self.rows[0]
        lines.append("| " + " | ".join(str(c) for c in header) + " |")
        lines.append("| " + " | ".join("---" for _ in header) + " |")
        
        for row in self.rows[1:]:
            lines.append("| " + " | ".join(str(c) for c in row) + " |")
        
        return "\n".join(lines)


@dataclass
class DocImage:
    """이미지"""
    filename: str
    data: bytes
    content_type: str = "image/png"


@dataclass
class DocDocument:
    """DOC 문서"""
    paragraphs: List[DocParagraph] = field(default_factory=list)
    tables: List[DocTable] = field(default_factory=list)
    images: List[DocImage] = field(default_factory=list)
    
    title: str = ""
    author: str = ""
    created: str = ""
    
    def get_text(self) -> str:
        """전체 텍스트"""
        lines = []
        for p in self.paragraphs:
            if p.text.strip():
                if p.is_heading and p.heading_level:
                    lines.append('#' * p.heading_level + ' ' + p.text)
                else:
                    lines.append(p.text)
        return '\n\n'.join(lines)
    
    def get_headings(self) -> List[Tuple[int, str]]:
        """헤딩 목록"""
        return [(p.heading_level, p.text) for p in self.paragraphs if p.is_heading]


def parse_doc(filepath_or_bytes: Union[str, bytes]) -> DocDocument:
    """
    DOC 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        DocDocument: 파싱된 문서
    """
    if isinstance(filepath_or_bytes, (str, Path)):
        with open(filepath_or_bytes, 'rb') as f:
            data = f.read()
    else:
        data = filepath_or_bytes
    
    if not is_ole2_file(data):
        raise ValueError("유효한 DOC 파일이 아닙니다")
    
    ole = OLE2Reader(data)
    doc = DocDocument()
    
    # WordDocument 스트림
    word_doc = ole.get_stream("WordDocument")
    if not word_doc:
        raise ValueError("WordDocument 스트림을 찾을 수 없습니다")
    
    # FIB (File Information Block) 파싱
    fib = _parse_fib(word_doc)
    
    # 테이블 스트림 (1Table 또는 0Table)
    table_stream = ole.get_stream("1Table")
    if not table_stream:
        table_stream = ole.get_stream("0Table")
    
    # 텍스트 추출
    text = _extract_text(word_doc, fib, table_stream)
    
    # 문단 분리
    for para_text in text.split('\r'):
        para_text = para_text.strip()
        if para_text:
            doc.paragraphs.append(DocParagraph(text=para_text))
    
    # 메타데이터
    summary = ole.get_stream("\x05SummaryInformation")
    if summary:
        doc.title, doc.author = _parse_summary(summary)
    
    return doc


def _parse_fib(word_doc: bytes) -> dict:
    """FIB (File Information Block) 파싱"""
    if len(word_doc) < 68:
        raise ValueError("WordDocument가 너무 작습니다")
    
    fib = {}
    
    # 기본 정보
    fib['wIdent'] = struct.unpack('<H', word_doc[0:2])[0]
    fib['nFib'] = struct.unpack('<H', word_doc[2:4])[0]
    
    # Word 97+ 확인
    if fib['wIdent'] != 0xA5EC:
        raise ValueError("Word 97+ 파일이 아닙니다")
    
    # 플래그
    flags = struct.unpack('<H', word_doc[10:12])[0]
    fib['fComplex'] = bool(flags & 0x0004)
    fib['fEncrypted'] = bool(flags & 0x0100)
    
    if fib['fEncrypted']:
        raise ValueError("암호화된 DOC 파일은 지원하지 않습니다")
    
    # 텍스트 위치 정보 (FibRgLw97)
    # 오프셋은 Word 버전에 따라 다를 수 있음
    base = 0x0018  # FibBase 크기
    
    # cbMac: 텍스트 + 기타 데이터 크기
    fib['cbMac'] = struct.unpack('<I', word_doc[base+4:base+8])[0]
    
    # ccpText: 메인 문서 텍스트 문자 수
    fib['ccpText'] = struct.unpack('<I', word_doc[base+76:base+80])[0]
    
    # ccpFtn: 각주 텍스트 문자 수
    fib['ccpFtn'] = struct.unpack('<I', word_doc[base+80:base+84])[0]
    
    # fcMin: 텍스트 시작 오프셋
    fib['fcMin'] = 0x200  # 보통 512 (섹터 크기)
    
    # fcMac: 텍스트 끝 오프셋
    fib['fcMac'] = fib['fcMin'] + fib['ccpText'] * 2  # Unicode
    
    return fib


def _extract_text(word_doc: bytes, fib: dict, table_stream: bytes) -> str:
    """텍스트 추출"""
    # 간단한 방법: 텍스트 영역 직접 읽기
    # 실제로는 clx, piece table 등 복잡한 구조 파싱 필요
    
    text_parts = []
    
    # Word 97+는 유니코드와 ANSI 혼합 사용
    # Piece Table을 통해 정확한 위치 파악 필요
    
    # 간단한 휴리스틱: 텍스트 영역 스캔
    fc_min = fib.get('fcMin', 0x200)
    ccp_text = fib.get('ccpText', 0)
    
    if ccp_text == 0:
        # ccpText가 0이면 전체 스캔
        return _scan_text(word_doc)
    
    # 텍스트 영역 읽기 시도
    try:
        # Unicode 시도
        text_data = word_doc[fc_min:fc_min + ccp_text * 2]
        text = text_data.decode('utf-16le', errors='ignore')
        text = _clean_text(text)
        if text:
            return text
    except:
        pass
    
    # ANSI 시도
    try:
        text_data = word_doc[fc_min:fc_min + ccp_text]
        text = text_data.decode('cp1252', errors='ignore')
        text = _clean_text(text)
        if text:
            return text
    except:
        pass
    
    # 폴백: 전체 스캔
    return _scan_text(word_doc)


def _scan_text(data: bytes) -> str:
    """전체 데이터에서 텍스트 스캔"""
    # 여러 인코딩 시도
    
    # UTF-16LE 시도 (Word 97+)
    try:
        # 텍스트로 보이는 영역 찾기
        text = data.decode('utf-16le', errors='ignore')
        text = _clean_text(text)
        if len(text) > 100:  # 충분한 텍스트
            return text
    except:
        pass
    
    # CP1252 시도 (ANSI)
    try:
        text = data.decode('cp1252', errors='ignore')
        text = _clean_text(text)
        if len(text) > 100:
            return text
    except:
        pass
    
    # CP949 시도 (한글)
    try:
        text = data.decode('cp949', errors='ignore')
        text = _clean_text(text)
        if len(text) > 100:
            return text
    except:
        pass
    
    return ""


def _clean_text(text: str) -> str:
    """텍스트 정리"""
    result = []
    
    for char in text:
        code = ord(char)
        
        # 제어 문자 처리
        if code == 0:
            continue
        elif code == 7:  # 셀 구분
            result.append('\t')
        elif code == 9:  # 탭
            result.append('\t')
        elif code == 10:  # 줄바꿈
            result.append('\n')
        elif code == 11:  # 강제 줄바꿈
            result.append('\n')
        elif code == 12:  # 페이지 나누기
            result.append('\n')
        elif code == 13:  # 문단 끝
            result.append('\r')
        elif code == 14:  # 열 나누기
            result.append('\n')
        elif code < 32:  # 기타 제어 문자
            continue
        elif code >= 32 and code < 127:  # ASCII
            result.append(char)
        elif code >= 0xAC00 and code <= 0xD7A3:  # 한글
            result.append(char)
        elif code >= 0x4E00 and code <= 0x9FFF:  # 한자
            result.append(char)
        elif code >= 0x3000 and code <= 0x303F:  # CJK 구두점
            result.append(char)
        elif code >= 0xFF00 and code <= 0xFFEF:  # 전각 문자
            result.append(char)
        elif code >= 127:  # 확장 문자
            result.append(char)
    
    return ''.join(result)


def _parse_summary(data: bytes) -> Tuple[str, str]:
    """SummaryInformation에서 메타데이터 추출"""
    # OLE Property Set 포맷 파싱
    # 간단한 구현: 문자열 스캔
    
    title = ""
    author = ""
    
    try:
        # UTF-16LE로 시도
        text = data.decode('utf-16le', errors='ignore')
        # 실제로는 Property Set 구조 파싱 필요
    except:
        pass
    
    return title, author
