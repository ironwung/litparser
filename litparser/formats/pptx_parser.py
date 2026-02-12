"""
PPTX Parser (Office Open XML)

.pptx 파일 파싱 - 순수 Python, 외부 라이브러리 없음
PPTX = ZIP 파일

구조:
  ppt/presentation.xml  <- 프레젠테이션 정보
  ppt/slides/slide1.xml <- 슬라이드
  ppt/media/            <- 이미지
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional
from io import BytesIO
import re


# OOXML 네임스페이스
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


@dataclass
class PptxTextBox:
    """텍스트 박스"""
    text: str
    x: int = 0
    y: int = 0
    width: int = 0
    height: int = 0
    is_title: bool = False


@dataclass
class PptxTable:
    """테이블"""
    rows: List[List[str]] = field(default_factory=list)
    
    def to_markdown(self) -> str:
        if not self.rows:
            return ""
        lines = []
        max_cols = max(len(row) for row in self.rows)
        
        # 헤더
        header = self.rows[0] + [''] * (max_cols - len(self.rows[0]))
        lines.append('| ' + ' | '.join(header) + ' |')
        lines.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
        
        # 본문
        for row in self.rows[1:]:
            row = row + [''] * (max_cols - len(row))
            lines.append('| ' + ' | '.join(row) + ' |')
        
        return '\n'.join(lines)


@dataclass
class PptxImage:
    """이미지"""
    filename: str
    data: bytes
    content_type: str = ""


@dataclass
class PptxSlide:
    """슬라이드"""
    number: int
    title: str = ""
    texts: List[PptxTextBox] = field(default_factory=list)
    tables: List[PptxTable] = field(default_factory=list)
    notes: str = ""
    
    def get_text(self) -> str:
        """슬라이드 텍스트"""
        lines = []
        if self.title:
            lines.append(f"# {self.title}")
        for tb in self.texts:
            if tb.text.strip() and tb.text != self.title:
                lines.append(tb.text)
        return '\n\n'.join(lines)


@dataclass
class PptxDocument:
    """파싱된 PPTX 문서"""
    slides: List[PptxSlide] = field(default_factory=list)
    images: List[PptxImage] = field(default_factory=list)
    
    # 메타데이터
    title: str = ""
    author: str = ""
    slide_count: int = 0
    
    def get_text(self) -> str:
        """전체 텍스트"""
        parts = []
        for slide in self.slides:
            parts.append(f"--- Slide {slide.number} ---")
            parts.append(slide.get_text())
        return '\n\n'.join(parts)
    
    def get_outline(self) -> List[str]:
        """슬라이드 제목 목록"""
        return [s.title for s in self.slides if s.title]


def parse_pptx(filepath_or_bytes) -> PptxDocument:
    """
    PPTX 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        PptxDocument
    """
    # ZIP 열기
    if isinstance(filepath_or_bytes, str):
        zf = zipfile.ZipFile(filepath_or_bytes, 'r')
    else:
        zf = zipfile.ZipFile(BytesIO(filepath_or_bytes), 'r')
    
    doc = PptxDocument()
    
    try:
        # 슬라이드 파일 목록
        slide_files = sorted([
            f for f in zf.namelist() 
            if re.match(r'ppt/slides/slide\d+\.xml$', f)
        ], key=lambda x: int(re.search(r'slide(\d+)', x).group(1)))
        
        doc.slide_count = len(slide_files)
        
        # 각 슬라이드 파싱
        for i, slide_file in enumerate(slide_files, 1):
            content = zf.read(slide_file)
            slide = _parse_slide_xml(content, i)
            
            # 노트 파싱
            notes_file = f'ppt/notesSlides/notesSlide{i}.xml'
            if notes_file in zf.namelist():
                notes_content = zf.read(notes_file)
                slide.notes = _parse_notes_xml(notes_content)
            
            doc.slides.append(slide)
        
        # 이미지 추출
        _extract_images(zf, doc)
        
        # 메타데이터
        if 'docProps/core.xml' in zf.namelist():
            core_content = zf.read('docProps/core.xml')
            _parse_core_xml(core_content, doc)
    
    finally:
        zf.close()
    
    return doc


def _parse_slide_xml(content: bytes, slide_num: int) -> PptxSlide:
    """슬라이드 XML 파싱"""
    slide = PptxSlide(number=slide_num)
    root = ET.fromstring(content)
    
    # 모든 텍스트 프레임 찾기
    for sp in root.findall('.//p:sp', NS):
        textbox = _parse_shape(sp)
        if textbox:
            slide.texts.append(textbox)
            # 첫 번째 제목으로 설정
            if textbox.is_title and not slide.title:
                slide.title = textbox.text
    
    # 테이블 찾기
    for tbl in root.findall('.//a:tbl', NS):
        table = _parse_table(tbl)
        if table:
            slide.tables.append(table)
    
    # 제목이 없으면 첫 번째 텍스트를 제목으로
    if not slide.title and slide.texts:
        slide.title = slide.texts[0].text.split('\n')[0][:50]
    
    return slide


def _parse_shape(sp) -> Optional[PptxTextBox]:
    """도형(텍스트박스) 파싱"""
    # 텍스트 추출
    texts = []
    for t in sp.findall('.//a:t', NS):
        if t.text:
            texts.append(t.text)
    
    if not texts:
        return None
    
    text = ''.join(texts)
    
    # 제목 여부 체크 (placeholder type)
    is_title = False
    nvSpPr = sp.find('.//p:nvSpPr', NS)
    if nvSpPr is not None:
        nvPr = nvSpPr.find('p:nvPr', NS)
        if nvPr is not None:
            ph = nvPr.find('p:ph', NS)
            if ph is not None:
                ph_type = ph.get('type', '')
                if ph_type in ['title', 'ctrTitle']:
                    is_title = True
    
    return PptxTextBox(text=text, is_title=is_title)


def _parse_table(tbl) -> Optional[PptxTable]:
    """테이블 파싱"""
    table = PptxTable()
    
    for tr in tbl.findall('.//a:tr', NS):
        row = []
        for tc in tr.findall('.//a:tc', NS):
            cell_texts = []
            for t in tc.findall('.//a:t', NS):
                if t.text:
                    cell_texts.append(t.text)
            row.append(' '.join(cell_texts))
        
        if row:
            table.rows.append(row)
    
    return table if table.rows else None


def _parse_notes_xml(content: bytes) -> str:
    """노트 XML 파싱"""
    root = ET.fromstring(content)
    texts = []
    
    for t in root.findall('.//a:t', NS):
        if t.text:
            texts.append(t.text)
    
    return ''.join(texts)


def _extract_images(zf: zipfile.ZipFile, doc: PptxDocument):
    """이미지 추출"""
    for name in zf.namelist():
        if name.startswith('ppt/media/'):
            data = zf.read(name)
            filename = name.split('/')[-1]
            
            ext = filename.split('.')[-1].lower()
            content_types = {
                'png': 'image/png',
                'jpg': 'image/jpeg',
                'jpeg': 'image/jpeg',
                'gif': 'image/gif',
                'bmp': 'image/bmp',
                'emf': 'image/x-emf',
                'wmf': 'image/x-wmf',
            }
            
            doc.images.append(PptxImage(
                filename=filename,
                data=data,
                content_type=content_types.get(ext, 'application/octet-stream')
            ))


def _parse_core_xml(content: bytes, doc: PptxDocument):
    """메타데이터 파싱"""
    dc_ns = 'http://purl.org/dc/elements/1.1/'
    
    root = ET.fromstring(content)
    
    title = root.find(f'{{{dc_ns}}}title')
    if title is not None and title.text:
        doc.title = title.text
    
    creator = root.find(f'{{{dc_ns}}}creator')
    if creator is not None and creator.text:
        doc.author = creator.text


def extract_text(doc: PptxDocument) -> str:
    """순수 텍스트 추출"""
    return doc.get_text()


def extract_tables(doc: PptxDocument) -> List[PptxTable]:
    """모든 테이블 추출"""
    tables = []
    for slide in doc.slides:
        tables.extend(slide.tables)
    return tables
