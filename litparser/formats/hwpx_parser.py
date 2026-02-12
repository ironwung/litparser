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
from typing import List, Dict, Optional, Set
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


def _get_tag(elem) -> str:
    """네임스페이스를 제거한 태그명 반환"""
    return elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag


def _collect_all_texts(elem) -> str:
    """
    요소 내 모든 텍스트 수집 (text + tail 포함)
    lineBreak 등의 tail 텍스트도 누락 없이 수집
    """
    texts = []
    for node in elem.iter():
        tag = _get_tag(node)
        if tag in ('t', 'char') and node.text:
            texts.append(node.text)
        # lineBreak 등의 tail에 있는 텍스트 수집
        if tag in ('lineBreak', 'linebreak') and node.tail:
            texts.append(node.tail)
    return ''.join(texts)


def _parse_section_xml(content: bytes, doc: HwpxDocument):
    """
    섹션 XML 파싱 - 문서 순서 보존
    
    핵심: 최상위 <p> 요소들을 순서대로 처리하여
    레이아웃 테이블 내 텍스트도 올바른 위치에 삽입
    """
    try:
        root = ET.fromstring(content)
    except ET.ParseError:
        return
    
    # 먼저 모든 tbl 요소 내부의 p 요소 id를 수집 (중복 방지용)
    table_inner_p_ids: Set[int] = set()
    for elem in root.iter():
        tag = _get_tag(elem)
        if tag == 'tbl':
            for inner in elem.iter():
                if _get_tag(inner) == 'p':
                    table_inner_p_ids.add(id(inner))
    
    # 최상위 p 요소들을 문서 순서대로 처리
    for elem in root:
        tag = _get_tag(elem)
        
        if tag != 'p':
            continue
        
        # 이 p에 tbl이 포함되어 있는지 확인
        has_tbl = any(_get_tag(sub) == 'tbl' for sub in elem.iter())
        
        if has_tbl:
            # tbl 포함 p: 테이블과 비테이블 텍스트를 분리 처리
            _process_p_with_table(elem, doc)
        else:
            # 일반 p: 테이블 내부 p는 건너뜀
            if id(elem) in table_inner_p_ids:
                continue
            para = _parse_paragraph(elem)
            if para and para.text.strip():
                doc.paragraphs.append(para)


def _process_p_with_table(p_elem, doc: HwpxDocument):
    """
    tbl을 포함한 p 요소 처리
    - 데이터 테이블 → doc.tables에 추가
    - 레이아웃 테이블 → 텍스트를 doc.paragraphs에 추가 (문서 순서 유지)
    - p 자체의 비테이블 텍스트도 수집
    """
    # 이 p 내의 최상위 tbl만 찾기 (중첩 방지)
    tbl_elements = []
    _find_top_level_tables(p_elem, tbl_elements)
    
    for tbl_elem in tbl_elements:
        table = _parse_table(tbl_elem)
        if table and table.rows:
            if _is_data_table(table):
                doc.tables.append(table)
            # 데이터든 레이아웃이든 텍스트는 항상 paragraphs에 추가 (문서 순서 유지)
            _add_layout_table_as_text(table, doc)
    
    # p 자체의 비테이블 텍스트 (tbl 바깥의 텍스트)
    p_texts = []
    _collect_texts_skip_tables(p_elem, p_texts)
    p_text = ''.join(p_texts).strip()
    if p_text:
        existing = {p.text.strip() for p in doc.paragraphs}
        if p_text not in existing:
            doc.paragraphs.append(HwpxParagraph(text=p_text))


def _find_top_level_tables(elem, result: list):
    """최상위 tbl만 찾기 (중첩 tbl 제외)"""
    for child in elem:
        tag = _get_tag(child)
        if tag == 'tbl':
            result.append(child)
            # tbl 내부는 더 이상 탐색하지 않음 (중첩 방지)
        else:
            _find_top_level_tables(child, result)


def _is_data_table(table: HwpxTable) -> bool:
    """
    데이터 테이블인지 레이아웃 테이블인지 판별
    """
    if not table.rows:
        return False
    
    num_rows = len(table.rows)
    max_cols = max(len(row) for row in table.rows)
    
    # 1행짜리 테이블: 거의 항상 레이아웃 용도
    if num_rows == 1:
        return False
    
    # 2행이상이지만, 실질적 데이터가 없는 경우
    non_empty_cells = 0
    total_cells = 0
    for row in table.rows:
        for cell in row:
            total_cells += 1
            if cell.strip():
                non_empty_cells += 1
    
    # 비어있는 셀이 대부분이면 레이아웃
    if total_cells > 0 and non_empty_cells / total_cells < 0.3:
        return False
    
    # 2행 이상이고 2열 이상이면 데이터 테이블
    if num_rows >= 2 and max_cols >= 2:
        return True
    
    # 2행 이상이지만 1열이면 레이아웃 (세로로 나열된 텍스트)
    if max_cols <= 1:
        return False
    
    return True


def _add_layout_table_as_text(table: HwpxTable, doc: HwpxDocument):
    """레이아웃 테이블의 텍스트를 일반 문단으로 추가"""
    for row in table.rows:
        texts = [cell.strip() for cell in row if cell.strip()]
        if texts:
            combined = ' '.join(texts)
            # 이미 동일한 텍스트가 문단에 있는지 확인 (중복 방지)
            existing_texts = {p.text.strip() for p in doc.paragraphs}
            if combined not in existing_texts:
                doc.paragraphs.append(HwpxParagraph(text=combined))


def _parse_paragraph(elem) -> Optional[HwpxParagraph]:
    """문단 파싱"""
    texts = []
    
    # 모든 텍스트 노드 수집 (단, tbl 내부는 건너뜀)
    _collect_texts_skip_tables(elem, texts)
    
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


def _collect_texts_skip_tables(elem, texts: list):
    """
    텍스트 노드를 수집하되 tbl 요소 내부는 건너뜀
    lineBreak의 tail 텍스트도 수집
    """
    for node in elem:
        tag = _get_tag(node)
        
        # 테이블 내부는 건너뜀
        if tag == 'tbl':
            continue
        
        # 텍스트 노드
        if tag in ('t', 'char') and node.text:
            texts.append(node.text)
        
        # lineBreak의 tail에 있는 텍스트 수집
        if tag in ('lineBreak', 'linebreak') and node.tail:
            texts.append(node.tail)
        
        # 재귀적으로 자식 탐색
        _collect_texts_skip_tables(node, texts)


def _parse_table(elem) -> Optional[HwpxTable]:
    """테이블 파싱 - lineBreak tail 텍스트도 수집"""
    table = HwpxTable()
    
    # 행 찾기
    for row_elem in elem.iter():
        row_tag = _get_tag(row_elem)
        
        if row_tag == 'tr':
            row = []
            
            # 셀 찾기
            for cell_elem in row_elem.iter():
                cell_tag = _get_tag(cell_elem)
                
                if cell_tag == 'tc':
                    # 셀 내 모든 텍스트 수집 (lineBreak tail 포함)
                    cell_text = _collect_all_texts(cell_elem)
                    row.append(cell_text)
            
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
        tag = _get_tag(elem)
        
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