"""
DOCX Parser (Office Open XML)

.docx 파일 파싱 - 순수 Python, 외부 라이브러리 없음
DOCX = ZIP 파일 (XML 문서들 포함)

구조:
  [Content_Types].xml
  _rels/.rels
  word/document.xml  <- 본문
  word/styles.xml    <- 스타일
  word/media/        <- 이미지
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, BinaryIO
from io import BytesIO
import re


# OOXML 네임스페이스
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}


@dataclass
class DocxParagraph:
    """문단"""
    text: str
    style: str = ""  # Heading1, Normal, etc.
    is_heading: bool = False
    heading_level: int = 0
    is_list_item: bool = False
    list_level: int = 0


@dataclass
class DocxTable:
    """테이블"""
    rows: List[List[str]] = field(default_factory=list)
    
    def to_markdown(self) -> str:
        if not self.rows:
            return ""
        lines = []
        # 헤더
        lines.append('| ' + ' | '.join(self.rows[0]) + ' |')
        lines.append('| ' + ' | '.join(['---'] * len(self.rows[0])) + ' |')
        # 본문
        for row in self.rows[1:]:
            # 열 수 맞추기
            while len(row) < len(self.rows[0]):
                row.append('')
            lines.append('| ' + ' | '.join(row[:len(self.rows[0])]) + ' |')
        return '\n'.join(lines)


@dataclass
class DocxImage:
    """이미지"""
    filename: str
    data: bytes
    content_type: str = ""
    width: int = 0
    height: int = 0


@dataclass
class DocxDocument:
    """파싱된 DOCX 문서"""
    paragraphs: List[DocxParagraph] = field(default_factory=list)
    tables: List[DocxTable] = field(default_factory=list)
    images: List[DocxImage] = field(default_factory=list)
    
    # 메타데이터
    title: str = ""
    author: str = ""
    created: str = ""
    modified: str = ""
    
    def get_text(self) -> str:
        """전체 텍스트 추출"""
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
        """헤딩 목록 [(level, text), ...]"""
        return [(p.heading_level, p.text) for p in self.paragraphs if p.is_heading]


def parse_docx(filepath_or_bytes) -> DocxDocument:
    """
    DOCX 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
    
    Returns:
        DocxDocument
    """
    # ZIP 열기
    if isinstance(filepath_or_bytes, str):
        zf = zipfile.ZipFile(filepath_or_bytes, 'r')
    else:
        zf = zipfile.ZipFile(BytesIO(filepath_or_bytes), 'r')
    
    doc = DocxDocument()
    
    try:
        # 본문 파싱
        if 'word/document.xml' in zf.namelist():
            content = zf.read('word/document.xml')
            _parse_document_xml(content, doc)
        
        # 스타일 정보 (헤딩 판별용)
        styles = {}
        if 'word/styles.xml' in zf.namelist():
            styles_content = zf.read('word/styles.xml')
            styles = _parse_styles_xml(styles_content)
        
        # 스타일 기반 헤딩 업데이트
        _update_heading_info(doc, styles)
        
        # 이미지 추출
        _extract_images(zf, doc)
        
        # 메타데이터
        if 'docProps/core.xml' in zf.namelist():
            core_content = zf.read('docProps/core.xml')
            _parse_core_xml(core_content, doc)
    
    finally:
        zf.close()
    
    return doc


def _parse_document_xml(content: bytes, doc: DocxDocument):
    """document.xml 파싱"""
    root = ET.fromstring(content)
    body = root.find('.//w:body', NS)
    
    if body is None:
        return
    
    for elem in body:
        tag = elem.tag.split('}')[-1]
        
        if tag == 'p':  # 문단
            para = _parse_paragraph(elem)
            if para:
                doc.paragraphs.append(para)
        
        elif tag == 'tbl':  # 테이블
            table = _parse_table(elem)
            if table:
                doc.tables.append(table)


def _parse_paragraph(elem) -> Optional[DocxParagraph]:
    """문단 파싱"""
    # 텍스트 추출
    texts = []
    for t in elem.findall('.//w:t', NS):
        if t.text:
            texts.append(t.text)
    
    text = ''.join(texts)
    
    # 스타일 정보
    style = ""
    pPr = elem.find('w:pPr', NS)
    if pPr is not None:
        pStyle = pPr.find('w:pStyle', NS)
        if pStyle is not None:
            style = pStyle.get(f'{{{NS["w"]}}}val', '')
    
    # 리스트 아이템 체크
    is_list = False
    list_level = 0
    if pPr is not None:
        numPr = pPr.find('w:numPr', NS)
        if numPr is not None:
            is_list = True
            ilvl = numPr.find('w:ilvl', NS)
            if ilvl is not None:
                list_level = int(ilvl.get(f'{{{NS["w"]}}}val', '0'))
    
    return DocxParagraph(
        text=text,
        style=style,
        is_list_item=is_list,
        list_level=list_level
    )


def _parse_table(elem) -> Optional[DocxTable]:
    """테이블 파싱"""
    table = DocxTable()
    
    for tr in elem.findall('.//w:tr', NS):
        row = []
        for tc in tr.findall('.//w:tc', NS):
            # 셀 내 모든 텍스트
            cell_texts = []
            for t in tc.findall('.//w:t', NS):
                if t.text:
                    cell_texts.append(t.text)
            row.append(' '.join(cell_texts))
        
        if row:
            table.rows.append(row)
    
    return table if table.rows else None


def _parse_styles_xml(content: bytes) -> Dict[str, dict]:
    """styles.xml 파싱 - 스타일 ID와 헤딩 레벨 매핑"""
    styles = {}
    root = ET.fromstring(content)
    
    for style in root.findall('.//w:style', NS):
        style_id = style.get(f'{{{NS["w"]}}}styleId', '')
        style_type = style.get(f'{{{NS["w"]}}}type', '')
        
        if style_type == 'paragraph':
            # outlineLvl로 헤딩 레벨 판별
            outline = style.find('.//w:outlineLvl', NS)
            if outline is not None:
                level = int(outline.get(f'{{{NS["w"]}}}val', '-1')) + 1
                if 1 <= level <= 9:
                    styles[style_id] = {'heading_level': level}
            
            # 또는 이름으로 판별
            name = style.find('w:name', NS)
            if name is not None:
                name_val = name.get(f'{{{NS["w"]}}}val', '')
                match = re.match(r'[Hh]eading\s*(\d)', name_val)
                if match:
                    styles[style_id] = {'heading_level': int(match.group(1))}
    
    return styles


def _update_heading_info(doc: DocxDocument, styles: Dict[str, dict]):
    """스타일 정보로 헤딩 업데이트"""
    for para in doc.paragraphs:
        if para.style in styles:
            info = styles[para.style]
            if 'heading_level' in info:
                para.is_heading = True
                para.heading_level = info['heading_level']
        
        # 스타일 이름으로도 체크
        if para.style.lower().startswith('heading'):
            match = re.search(r'\d', para.style)
            if match:
                para.is_heading = True
                para.heading_level = int(match.group())


def _extract_images(zf: zipfile.ZipFile, doc: DocxDocument):
    """이미지 추출"""
    for name in zf.namelist():
        if name.startswith('word/media/'):
            data = zf.read(name)
            filename = name.split('/')[-1]
            
            # Content-Type 추정
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
            
            doc.images.append(DocxImage(
                filename=filename,
                data=data,
                content_type=content_types.get(ext, 'application/octet-stream')
            ))


def _parse_core_xml(content: bytes, doc: DocxDocument):
    """메타데이터 파싱"""
    # 네임스페이스
    dc_ns = 'http://purl.org/dc/elements/1.1/'
    cp_ns = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
    dcterms_ns = 'http://purl.org/dc/terms/'
    
    root = ET.fromstring(content)
    
    # 제목
    title = root.find(f'{{{dc_ns}}}title')
    if title is not None and title.text:
        doc.title = title.text
    
    # 작성자
    creator = root.find(f'{{{dc_ns}}}creator')
    if creator is not None and creator.text:
        doc.author = creator.text
    
    # 생성일
    created = root.find(f'{{{dcterms_ns}}}created')
    if created is not None and created.text:
        doc.created = created.text
    
    # 수정일
    modified = root.find(f'{{{dcterms_ns}}}modified')
    if modified is not None and modified.text:
        doc.modified = modified.text


def extract_text(doc: DocxDocument) -> str:
    """순수 텍스트 추출"""
    return doc.get_text()


def extract_tables(doc: DocxDocument) -> List[DocxTable]:
    """테이블 추출"""
    return doc.tables
