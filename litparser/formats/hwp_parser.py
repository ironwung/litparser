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


# HWP 태그 ID
HWPTAG_PARA_TEXT = 67


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
        
        lines = []
        lines.append("| " + " | ".join(str(c) for c in self.rows[0]) + " |")
        lines.append("| " + " | ".join("---" for _ in self.rows[0]) + " |")
        for row in self.rows[1:]:
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
            
            paragraphs = self._parse_section(stream_data)
            self.doc.paragraphs.extend(paragraphs)
            section_idx += 1
        
        if section_idx == 0:
            for entry in self.ole.entries:
                if entry.name.startswith("Section") and entry.entry_type == 2:
                    stream_data = self.ole._read_stream_data(entry)
                    if stream_data:
                        paragraphs = self._parse_section(stream_data)
                        self.doc.paragraphs.extend(paragraphs)
    
    def _parse_section(self, data: bytes) -> List[HwpParagraph]:
        """섹션 파싱"""
        paragraphs = []
        
        if self.is_compressed:
            try:
                data = zlib.decompress(data, -15)
            except zlib.error:
                try:
                    data = zlib.decompress(data)
                except zlib.error:
                    return paragraphs
        
        pos = 0
        size = len(data)
        
        while pos + 4 <= size:
            header = struct.unpack('<I', data[pos:pos+4])[0]
            pos += 4
            
            tag_id = header & 0x3FF
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
            
            if tag_id == HWPTAG_PARA_TEXT and rec_data:
                text = self._extract_text(rec_data)
                if text.strip():
                    paragraphs.append(HwpParagraph(text=text))
        
        return paragraphs
    
    def _extract_text(self, data: bytes) -> str:
        """텍스트 디코딩 - HWP 제어 문자 처리"""
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
        
        text_chars = []
        i = 0
        
        while i < len(data) - 1:
            # 2바이트씩 읽기 (UTF-16LE)
            char_code = struct.unpack('<H', data[i:i+2])[0]
            
            if char_code == 0:  # NULL
                i += 2
                continue
            elif char_code < 32:  # 제어 문자
                extra = CTRL_EXTRA_BYTES.get(char_code, 12)
                i += 2 + extra
                
                # 줄바꿈 처리
                if char_code in (10, 13):  # line break, paragraph end
                    text_chars.append('\n')
                elif char_code == 9:  # tab
                    text_chars.append('\t')
            else:
                # 일반 문자
                text_chars.append(chr(char_code))
                i += 2
        
        return ''.join(text_chars).strip()
    
    def _parse_metadata(self):
        """메타데이터 추출"""
        summary = self.ole.get_stream("\x05SummaryInformation")
        if summary:
            pass  # 복잡한 포맷이라 스킵


def parse_hwp(filepath_or_bytes: Union[str, bytes]) -> HwpDocument:
    """HWP 파일 파싱"""
    parser = HwpParser(filepath_or_bytes)
    return parser.parse()
