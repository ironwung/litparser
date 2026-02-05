"""
Document Output Formatter

파싱 결과를 다양한 포맷으로 변환:
- Markdown: 읽기 좋은 문서 형태
- JSON: 프로그래밍 처리용 구조화 데이터
"""

import json
import base64
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional, Union
from pathlib import Path


@dataclass
class DocumentOutput:
    """통합 문서 출력 구조"""
    # 메타데이터
    filename: str = ""
    format: str = ""  # pdf, docx, pptx, hwpx, txt, md
    page_count: int = 0
    title: str = ""
    author: str = ""
    
    # 콘텐츠
    text: str = ""  # 전체 텍스트
    pages: List[Dict] = None  # 페이지별 데이터
    tables: List[Dict] = None  # 테이블 목록
    images: List[Dict] = None  # 이미지 목록
    headings: List[Dict] = None  # 헤딩/개요
    
    def __post_init__(self):
        if self.pages is None:
            self.pages = []
        if self.tables is None:
            self.tables = []
        if self.images is None:
            self.images = []
        if self.headings is None:
            self.headings = []


def to_markdown(doc_output: DocumentOutput, include_images: bool = False) -> str:
    """
    DocumentOutput을 마크다운으로 변환
    
    Args:
        doc_output: 파싱 결과
        include_images: 이미지를 base64로 포함할지
    
    Returns:
        str: 마크다운 문서
    """
    lines = []
    
    # 제목
    title = doc_output.title or Path(doc_output.filename).stem
    lines.append(f"# {title}")
    lines.append("")
    
    # 메타데이터
    if doc_output.author or doc_output.page_count > 1:
        lines.append("> **문서 정보**")
        if doc_output.author:
            lines.append(f"> - 작성자: {doc_output.author}")
        if doc_output.format:
            lines.append(f"> - 포맷: {doc_output.format.upper()}")
        if doc_output.page_count > 1:
            lines.append(f"> - 페이지: {doc_output.page_count}")
        lines.append("")
    
    # 목차 (헤딩이 있으면)
    if doc_output.headings:
        lines.append("## 목차")
        lines.append("")
        for h in doc_output.headings:
            level = h.get('level', 1)
            text = h.get('text', '')
            indent = "  " * (level - 1)
            lines.append(f"{indent}- {text}")
        lines.append("")
        lines.append("---")
        lines.append("")
    
    # 페이지별 내용
    if doc_output.pages:
        for page in doc_output.pages:
            page_num = page.get('page', 1)
            
            if doc_output.page_count > 1:
                lines.append(f"## 페이지 {page_num}")
                lines.append("")
            
            # 텍스트
            text = page.get('text', '')
            if text:
                lines.append(text)
                lines.append("")
            
            # 테이블
            page_tables = page.get('tables', [])
            for i, table in enumerate(page_tables, 1):
                if doc_output.page_count > 1:
                    lines.append(f"### 테이블 {page_num}-{i}")
                else:
                    lines.append(f"### 테이블 {i}")
                lines.append("")
                lines.append(table.get('markdown', ''))
                lines.append("")
            
            # 이미지
            if include_images:
                page_images = page.get('images', [])
                for i, img in enumerate(page_images, 1):
                    alt = img.get('alt', f'이미지 {i}')
                    if img.get('base64'):
                        mime = img.get('mime', 'image/png')
                        b64 = img['base64']
                        lines.append(f"![{alt}](data:{mime};base64,{b64})")
                    elif img.get('path'):
                        lines.append(f"![{alt}]({img['path']})")
                    lines.append("")
    
    # 페이지 구분 없이 전체 텍스트만 있는 경우
    elif doc_output.text:
        lines.append(doc_output.text)
        lines.append("")
        
        # 테이블
        for i, table in enumerate(doc_output.tables, 1):
            lines.append(f"### 테이블 {i}")
            lines.append("")
            lines.append(table.get('markdown', ''))
            lines.append("")
    
    return "\n".join(lines)


def to_json(doc_output: DocumentOutput, 
            include_images: bool = False,
            image_format: str = "base64") -> str:
    """
    DocumentOutput을 JSON으로 변환
    
    Args:
        doc_output: 파싱 결과
        include_images: 이미지 데이터 포함 여부
        image_format: "base64" 또는 "path"
    
    Returns:
        str: JSON 문자열
    """
    data = {
        "metadata": {
            "filename": doc_output.filename,
            "format": doc_output.format,
            "page_count": doc_output.page_count,
            "title": doc_output.title,
            "author": doc_output.author,
        },
        "content": {
            "text": doc_output.text,
            "headings": doc_output.headings,
        },
        "pages": [],
        "tables": [],
        "images": [],
    }
    
    # 페이지 데이터
    for page in doc_output.pages:
        page_data = {
            "page": page.get('page', 1),
            "text": page.get('text', ''),
            "tables": [],
            "images": [],
        }
        
        # 테이블
        for table in page.get('tables', []):
            table_data = {
                "rows": table.get('rows', 0),
                "cols": table.get('cols', 0),
                "data": table.get('data', []),  # 2D array
                "markdown": table.get('markdown', ''),
            }
            page_data["tables"].append(table_data)
            data["tables"].append({
                "page": page.get('page', 1),
                **table_data
            })
        
        # 이미지
        if include_images:
            for img in page.get('images', []):
                img_data = {
                    "width": img.get('width', 0),
                    "height": img.get('height', 0),
                    "format": img.get('format', ''),
                }
                if image_format == "base64" and img.get('base64'):
                    img_data["base64"] = img['base64']
                    img_data["mime"] = img.get('mime', 'image/png')
                elif image_format == "path" and img.get('path'):
                    img_data["path"] = img['path']
                
                page_data["images"].append(img_data)
                data["images"].append({
                    "page": page.get('page', 1),
                    **img_data
                })
        
        data["pages"].append(page_data)
    
    # 페이지 없이 전체 데이터만 있는 경우
    if not doc_output.pages and doc_output.tables:
        for table in doc_output.tables:
            data["tables"].append({
                "rows": table.get('rows', 0),
                "cols": table.get('cols', 0),
                "data": table.get('data', []),
                "markdown": table.get('markdown', ''),
            })
    
    return json.dumps(data, ensure_ascii=False, indent=2)


def to_dict(doc_output: DocumentOutput, include_images: bool = False) -> dict:
    """DocumentOutput을 딕셔너리로 변환"""
    return json.loads(to_json(doc_output, include_images))


# =============================================================================
# 파서별 변환 함수
# =============================================================================

def pdf_to_output(doc, include_images: bool = True) -> DocumentOutput:
    """PDF 파싱 결과를 DocumentOutput으로 변환"""
    from . import (
        get_page_count, extract_text, extract_tables, 
        extract_images, get_document_outline
    )
    
    output = DocumentOutput()
    output.format = "pdf"
    output.page_count = get_page_count(doc)
    
    # 전체 텍스트
    all_text = []
    
    # 페이지별 처리
    for page_num in range(output.page_count):
        page_data = {"page": page_num + 1}
        
        # 텍스트
        text = extract_text(doc, page_num)
        page_data["text"] = text
        all_text.append(text)
        
        # 테이블
        tables = extract_tables(doc, page_num)
        page_data["tables"] = []
        for t in tables:
            table_data = {
                "rows": t.rows,
                "cols": t.cols,
                "data": t.to_list(),
                "markdown": t.to_markdown(),
            }
            page_data["tables"].append(table_data)
            output.tables.append(table_data)  # 전체 테이블 목록에도 추가
        
        # 이미지 (첫 페이지에서만 전체 추출)
        page_data["images"] = []
        
        output.pages.append(page_data)
    
    output.text = "\n\n".join(all_text)
    
    # 이미지
    if include_images:
        images = extract_images(doc)
        for img in images:
            img_data = {
                "width": img.width,
                "height": img.height,
                "format": img.color_space,
            }
            # base64 인코딩
            if img.data:
                img_data["base64"] = base64.b64encode(img.data).decode('ascii')
                img_data["mime"] = _guess_mime(img)
            output.images.append(img_data)
    
    # 개요
    try:
        outline = get_document_outline(doc)
        output.headings = [{"level": level, "text": text} for level, text in outline]
    except:
        pass
    
    return output


def docx_to_output(doc, include_images: bool = True) -> DocumentOutput:
    """DOCX 파싱 결과를 DocumentOutput으로 변환"""
    output = DocumentOutput()
    output.format = "docx"
    output.title = doc.title
    output.author = doc.author
    output.page_count = 1  # DOCX는 페이지 개념 없음
    
    # 텍스트
    output.text = doc.get_text()
    
    # 헤딩
    output.headings = [{"level": level, "text": text} for level, text in doc.get_headings()]
    
    # 테이블
    for t in doc.tables:
        output.tables.append({
            "rows": len(t.rows),
            "cols": len(t.rows[0]) if t.rows else 0,
            "data": t.rows,
            "markdown": t.to_markdown(),
        })
    
    # 이미지
    if include_images:
        for img in doc.images:
            img_data = {
                "filename": img.filename,
                "format": img.content_type,
                "width": img.width,
                "height": img.height,
            }
            if img.data:
                img_data["base64"] = base64.b64encode(img.data).decode('ascii')
                img_data["mime"] = img.content_type
            output.images.append(img_data)
    
    return output


def pptx_to_output(doc, include_images: bool = True) -> DocumentOutput:
    """PPTX 파싱 결과를 DocumentOutput으로 변환"""
    output = DocumentOutput()
    output.format = "pptx"
    output.title = doc.title
    output.author = doc.author
    output.page_count = doc.slide_count
    
    all_text = []
    
    for slide in doc.slides:
        page_data = {
            "page": slide.number,
            "title": slide.title,
            "text": slide.get_text(),
            "tables": [],
            "images": [],
        }
        all_text.append(slide.get_text())
        
        # 테이블
        for t in slide.tables:
            page_data["tables"].append({
                "rows": len(t.rows),
                "cols": len(t.rows[0]) if t.rows else 0,
                "data": t.rows,
                "markdown": t.to_markdown(),
            })
        
        output.pages.append(page_data)
        
        # 헤딩 (슬라이드 제목)
        if slide.title:
            output.headings.append({"level": 1, "text": slide.title})
    
    output.text = "\n\n".join(all_text)
    
    # 이미지
    if include_images:
        for img in doc.images:
            img_data = {
                "filename": img.filename,
                "format": img.content_type,
            }
            if img.data:
                img_data["base64"] = base64.b64encode(img.data).decode('ascii')
                img_data["mime"] = img.content_type
            output.images.append(img_data)
    
    return output


def hwpx_to_output(doc, include_images: bool = True) -> DocumentOutput:
    """HWPX 파싱 결과를 DocumentOutput으로 변환"""
    output = DocumentOutput()
    output.format = "hwpx"
    output.title = doc.title
    output.author = doc.author
    output.page_count = 1
    
    output.text = doc.get_text()
    output.headings = [{"level": level, "text": text} for level, text in doc.get_headings()]
    
    # 테이블
    for t in doc.tables:
        output.tables.append({
            "rows": len(t.rows),
            "cols": len(t.rows[0]) if t.rows else 0,
            "data": t.rows,
            "markdown": t.to_markdown(),
        })
    
    # 이미지
    if include_images:
        for img in doc.images:
            img_data = {
                "filename": img.filename,
                "format": img.content_type,
            }
            if img.data:
                img_data["base64"] = base64.b64encode(img.data).decode('ascii')
                img_data["mime"] = img.content_type
            output.images.append(img_data)
    
    return output


def text_to_output(doc, is_markdown: bool = False) -> DocumentOutput:
    """텍스트/마크다운 파싱 결과를 DocumentOutput으로 변환"""
    output = DocumentOutput()
    output.format = "md" if is_markdown else "txt"
    output.page_count = 1
    output.text = doc.content
    
    if is_markdown and doc.headings:
        output.headings = [{"level": level, "text": text} for level, text in doc.headings]
    
    return output


def _guess_mime(img) -> str:
    """이미지 MIME 타입 추측"""
    if hasattr(img, 'color_space'):
        cs = img.color_space.lower()
        if 'jpeg' in cs or 'jpg' in cs:
            return 'image/jpeg'
        elif 'jp2' in cs:
            return 'image/jp2'
    
    # 데이터 시그니처로 판단
    if img.data:
        if img.data[:2] == b'\xff\xd8':
            return 'image/jpeg'
        elif img.data[:8] == b'\x89PNG\r\n\x1a\n':
            return 'image/png'
        elif img.data[:12] == b'\x00\x00\x00\x0cjP  \r\n':
            return 'image/jp2'
    
    return 'image/png'
