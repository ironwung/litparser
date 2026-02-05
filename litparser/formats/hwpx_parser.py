"""
HWPX Parser (Open Word Processor XML - 한글 개방형 문서)

.hwpx 파일 파싱 - 순수 Python, 외부 라이브러리 없음
HWPX = ZIP 파일 (OWPML 표준)

구조:
  Contents/content.hpf  <- 본문 구조
  Contents/section0.xml <- 본문 내용
  Contents/header.xml   <- 헤더
  META-INF/manifest.xml <- 매니페스트
  Preview/PrvImage.png  <- 미리보기
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional
from io import BytesIO
import re


# HWPX 네임스페이스
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'opf': 'http://www.idpf.org/2007/opf',
    'odf': 'urn:oasis:names:tc:opendocument:xmlns:container',
}


@dataclass
class HwpxParagraph:
    """문단"""
    text: str
    style: str = ""
    is_heading: bool = False
    heading_level: int = 0


@dataclass
class HwpxTable:
    """테이블"""
    rows: List[List[str]] = field(default_factory=list)
    
    def to_markdown(self) -> str:
        if not self.rows:
            return ""
        lines = []
        max_cols = max(len(row) for row in self.rows) if self.rows else 0
        
        if max_cols == 0:
            return ""
        
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
class HwpxImage:
    """이미지"""
    filename: str
    data: bytes
    content_type: str = ""


@dataclass
class HwpxDocument:
    """파싱된 HWPX 문서"""
    paragraphs: List[HwpxParagraph] = field(default_factory=list)
    tables: List[HwpxTable] = field(default_factory=list)
    images: List[HwpxImage] = field(default_factory=list)
    
    # 메타데이터
    title: str = ""
    author: str = ""
    created: str = ""
    
    def get_text(self) -> str:
        """전체 텍스트"""
        lines = []
        for p in self.paragraphs:
            if p.text.strip():
                if p.is_heading:
                    prefix = '#' * p.heading_level + ' '
                    lines.append(prefix + p.text)
                else:
                    lines.append(p.text)
        return '\n\n'.join(lines)
    
    def get_headings(self) -> List[tuple]:
        """헤딩 목록"""
        return [(p.heading_level, p.text) for p in self.paragraphs if p.is_heading]


def parse_hwpx(filepath_or_bytes) -> HwpxDocument:
    """
    HWPX 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        HwpxDocument
    """
    # ZIP 열기
    if isinstance(filepath_or_bytes, str):
        zf = zipfile.ZipFile(filepath_or_bytes, 'r')
    else:
        zf = zipfile.ZipFile(BytesIO(filepath_or_bytes), 'r')
    
    doc = HwpxDocument()
    
    try:
        # 섹션 파일 찾기
        section_files = sorted([
            f for f in zf.namelist()
            if re.match(r'Contents/section\d+\.xml$', f)
        ])
        
        # 각 섹션 파싱
        for section_file in section_files:
            content = zf.read(section_file)
            _parse_section_xml(content, doc)
        
        # 이미지 추출
        _extract_images(zf, doc)
        
        # 메타데이터
        if 'Contents/content.hpf' in zf.namelist():
            hpf_content = zf.read('Contents/content.hpf')
            _parse_hpf(hpf_content, doc)
    
    finally:
        zf.close()
    
    return doc


def _parse_section_xml(content: bytes, doc: HwpxDocument):
    """섹션 XML 파싱"""
    try:
        root = ET.fromstring(content)
    except ET.ParseError:
        return
    
    # 모든 문단 찾기 (p 태그)
    # HWPX는 네임스페이스가 다양할 수 있음
    for elem in root.iter():
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        
        if tag == 'p':
            para = _parse_paragraph(elem)
            if para and para.text.strip():
                doc.paragraphs.append(para)
        
        elif tag == 'tbl':
            table = _parse_table(elem)
            if table and table.rows:
                doc.tables.append(table)


def _parse_paragraph(elem) -> Optional[HwpxParagraph]:
    """문단 파싱"""
    texts = []
    
    # 모든 텍스트 노드 수집
    for node in elem.iter():
        tag = node.tag.split('}')[-1] if '}' in node.tag else node.tag
        
        if tag == 't' and node.text:
            texts.append(node.text)
        elif tag == 'char' and node.text:
            texts.append(node.text)
    
    if not texts:
        # 직접 텍스트도 체크
        if elem.text:
            texts.append(elem.text)
        for child in elem:
            if child.tail:
                texts.append(child.tail)
    
    text = ''.join(texts)
    
    # 스타일/아웃라인 레벨 체크
    is_heading = False
    heading_level = 0
    
    # 아웃라인 레벨 속성 체크
    for attr in elem.attrib:
        if 'outlineLevel' in attr or 'level' in attr.lower():
            try:
                level = int(elem.attrib[attr])
                if 0 <= level <= 9:
                    is_heading = True
                    heading_level = level + 1
            except ValueError:
                pass
    
    return HwpxParagraph(
        text=text,
        is_heading=is_heading,
        heading_level=heading_level
    )


def _parse_table(elem) -> Optional[HwpxTable]:
    """테이블 파싱"""
    table = HwpxTable()
    
    # 행 찾기
    for row_elem in elem.iter():
        row_tag = row_elem.tag.split('}')[-1] if '}' in row_elem.tag else row_elem.tag
        
        if row_tag == 'tr':
            row = []
            
            # 셀 찾기
            for cell_elem in row_elem.iter():
                cell_tag = cell_elem.tag.split('}')[-1] if '}' in cell_elem.tag else cell_elem.tag
                
                if cell_tag == 'tc':
                    cell_texts = []
                    for t in cell_elem.iter():
                        t_tag = t.tag.split('}')[-1] if '}' in t.tag else t.tag
                        if t_tag in ['t', 'char'] and t.text:
                            cell_texts.append(t.text)
                    row.append(' '.join(cell_texts))
            
            if row:
                table.rows.append(row)
    
    return table if table.rows else None


def _extract_images(zf: zipfile.ZipFile, doc: HwpxDocument):
    """이미지 추출"""
    for name in zf.namelist():
        # BinData 또는 Media 폴더
        if '/BinData/' in name or '/media/' in name.lower():
            ext = name.split('.')[-1].lower()
            if ext in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'emf', 'wmf']:
                data = zf.read(name)
                filename = name.split('/')[-1]
                
                content_types = {
                    'png': 'image/png',
                    'jpg': 'image/jpeg',
                    'jpeg': 'image/jpeg',
                    'gif': 'image/gif',
                    'bmp': 'image/bmp',
                }
                
                doc.images.append(HwpxImage(
                    filename=filename,
                    data=data,
                    content_type=content_types.get(ext, 'application/octet-stream')
                ))


def _parse_hpf(content: bytes, doc: HwpxDocument):
    """content.hpf 메타데이터 파싱"""
    try:
        root = ET.fromstring(content)
    except ET.ParseError:
        return
    
    # 다양한 네임스페이스에서 메타데이터 찾기
    for elem in root.iter():
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        
        if tag == 'title' and elem.text:
            doc.title = elem.text
        elif tag == 'creator' and elem.text:
            doc.author = elem.text
        elif tag == 'date' and elem.text:
            doc.created = elem.text


def extract_text(doc: HwpxDocument) -> str:
    """순수 텍스트 추출"""
    return doc.get_text()


def extract_tables(doc: HwpxDocument) -> List[HwpxTable]:
    """테이블 추출"""
    return doc.tables
