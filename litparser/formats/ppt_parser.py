"""
PPT Parser - Microsoft PowerPoint 97-2003 (.ppt)

구조:
  OLE2 컨테이너
  ├── PowerPoint Document: 메인 스트림
  ├── Current User: 현재 사용자 정보
  ├── Pictures: 이미지 스트림
  └── SummaryInformation: 메타데이터

PowerPoint Document 스트림:
  - UserEditAtom: 편집 정보
  - DocumentContainer: 문서 데이터
  - SlideContainer: 슬라이드 데이터
"""

import struct
from dataclasses import dataclass, field
from typing import List, Optional, Union, Tuple
from pathlib import Path

from ..core.ole_parser import OLE2Reader, is_ole2_file


# PPT Record Types
RT_DOCUMENT = 0x03E8           # 1000
RT_SLIDE = 0x03EE              # 1006
RT_SLIDE_BASE = 0x03F0         # 1008
RT_MAIN_MASTER = 0x03F8        # 1016
RT_NOTES = 0x03F0              # 1008
RT_TEXT_HEADER = 0x0F9F        # 3999
RT_TEXT_CHARS = 0x0FA0         # 4000
RT_TEXT_BYTES = 0x0FA8         # 4008
RT_SLIDE_PERSIST_ATOM = 0x03F3 # 1011
RT_CSTRING = 0x0FBA            # 4026


@dataclass
class PptSlide:
    """슬라이드"""
    number: int
    title: str = ""
    texts: List[str] = field(default_factory=list)
    
    def get_text(self) -> str:
        parts = []
        if self.title:
            parts.append(self.title)
        parts.extend(self.texts)
        return '\n'.join(parts)


@dataclass
class PptTable:
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
class PptImage:
    """이미지"""
    filename: str
    data: bytes
    content_type: str = "image/png"


@dataclass
class PptDocument:
    """PPT 문서"""
    slides: List[PptSlide] = field(default_factory=list)
    tables: List[PptTable] = field(default_factory=list)
    images: List[PptImage] = field(default_factory=list)
    
    title: str = ""
    author: str = ""
    created: str = ""
    
    @property
    def slide_count(self) -> int:
        return len(self.slides)
    
    def get_text(self) -> str:
        """전체 텍스트"""
        parts = []
        for slide in self.slides:
            parts.append(f"=== 슬라이드 {slide.number} ===")
            parts.append(slide.get_text())
            parts.append("")
        return '\n'.join(parts)
    
    def get_headings(self) -> List[Tuple[int, str]]:
        """슬라이드 제목 목록"""
        return [(1, s.title) for s in self.slides if s.title]


def parse_ppt(filepath_or_bytes: Union[str, bytes]) -> PptDocument:
    """
    PPT 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        PptDocument: 파싱된 문서
    """
    if isinstance(filepath_or_bytes, (str, Path)):
        with open(filepath_or_bytes, 'rb') as f:
            data = f.read()
    else:
        data = filepath_or_bytes
    
    if not is_ole2_file(data):
        raise ValueError("유효한 PPT 파일이 아닙니다")
    
    ole = OLE2Reader(data)
    doc = PptDocument()
    
    # PowerPoint Document 스트림
    ppt_doc = ole.get_stream("PowerPoint Document")
    if not ppt_doc:
        raise ValueError("PowerPoint Document 스트림을 찾을 수 없습니다")
    
    # 레코드 파싱하여 텍스트 추출
    texts = _extract_all_texts(ppt_doc)
    
    # 슬라이드 구성
    current_slide = None
    slide_num = 0
    
    for text in texts:
        text = text.strip()
        if not text:
            continue
        
        # 새 슬라이드 감지 (휴리스틱)
        if current_slide is None or len(current_slide.texts) > 10:
            slide_num += 1
            current_slide = PptSlide(number=slide_num)
            doc.slides.append(current_slide)
        
        if not current_slide.title and len(text) < 100:
            current_slide.title = text
        else:
            current_slide.texts.append(text)
    
    # 슬라이드가 없으면 기본 생성
    if not doc.slides and texts:
        slide = PptSlide(number=1, texts=texts)
        doc.slides.append(slide)
    
    # 이미지 추출
    pictures = ole.get_stream("Pictures")
    if pictures:
        doc.images = _extract_pictures(pictures)
    
    return doc


def _extract_all_texts(data: bytes) -> List[str]:
    """모든 텍스트 레코드 추출"""
    texts = []
    pos = 0
    size = len(data)
    
    while pos + 8 <= size:
        # 레코드 헤더 (8 bytes)
        rec_ver = struct.unpack('<H', data[pos:pos+2])[0]
        rec_type = struct.unpack('<H', data[pos+2:pos+4])[0]
        rec_len = struct.unpack('<I', data[pos+4:pos+8])[0]
        pos += 8
        
        # 레코드 데이터
        if pos + rec_len > size:
            break
        
        rec_data = data[pos:pos+rec_len]
        
        # 텍스트 레코드 처리
        if rec_type == RT_TEXT_CHARS:
            # UTF-16LE 텍스트
            try:
                text = rec_data.decode('utf-16le', errors='ignore')
                text = _clean_text(text)
                if text:
                    texts.append(text)
            except:
                pass
        
        elif rec_type == RT_TEXT_BYTES:
            # ANSI 텍스트
            try:
                text = rec_data.decode('cp1252', errors='ignore')
                text = _clean_text(text)
                if text:
                    texts.append(text)
            except:
                pass
        
        elif rec_type == RT_CSTRING:
            # C 문자열 (UTF-16LE)
            try:
                text = rec_data.decode('utf-16le', errors='ignore').rstrip('\x00')
                text = _clean_text(text)
                if text:
                    texts.append(text)
            except:
                pass
        
        # 컨테이너 레코드는 내부 파싱
        ver_instance = rec_ver & 0x000F
        if ver_instance == 0x000F:
            # 컨테이너: 재귀 파싱
            inner_texts = _extract_all_texts(rec_data)
            texts.extend(inner_texts)
            pos += rec_len
            continue
        
        pos += rec_len
    
    return texts


def _clean_text(text: str) -> str:
    """텍스트 정리"""
    result = []
    
    for char in text:
        code = ord(char)
        
        if code == 0:
            continue
        elif code == 9:  # 탭
            result.append(' ')
        elif code == 10:  # 줄바꿈
            result.append('\n')
        elif code == 11:  # 수직 탭
            result.append('\n')
        elif code == 13:  # 캐리지 리턴
            continue
        elif code < 32:  # 기타 제어 문자
            continue
        else:
            result.append(char)
    
    return ''.join(result).strip()


def _extract_pictures(data: bytes) -> List[PptImage]:
    """Pictures 스트림에서 이미지 추출"""
    images = []
    pos = 0
    size = len(data)
    img_num = 0
    
    while pos + 8 <= size:
        # 레코드 헤더
        rec_ver = struct.unpack('<H', data[pos:pos+2])[0]
        rec_type = struct.unpack('<H', data[pos+2:pos+4])[0]
        rec_len = struct.unpack('<I', data[pos+4:pos+8])[0]
        pos += 8
        
        if pos + rec_len > size:
            break
        
        rec_data = data[pos:pos+rec_len]
        
        # 이미지 타입 (0x046A=PNG, 0x046B=JPEG 등)
        if rec_type in (0x046A, 0xF01A):  # PNG
            # 헤더 스킵 (17 bytes 정도)
            img_data = _find_image_start(rec_data, b'\x89PNG')
            if img_data:
                img_num += 1
                images.append(PptImage(
                    filename=f"image{img_num}.png",
                    data=img_data,
                    content_type="image/png"
                ))
        
        elif rec_type in (0x046B, 0xF01D, 0xF01E, 0xF01F, 0xF020):  # JPEG
            img_data = _find_image_start(rec_data, b'\xff\xd8')
            if img_data:
                img_num += 1
                images.append(PptImage(
                    filename=f"image{img_num}.jpg",
                    data=img_data,
                    content_type="image/jpeg"
                ))
        
        pos += rec_len
    
    return images


def _find_image_start(data: bytes, signature: bytes) -> Optional[bytes]:
    """이미지 시그니처 위치 찾기"""
    idx = data.find(signature)
    if idx >= 0:
        return data[idx:]
    return None
