"""
Text/Markdown Parser

txt, md 파일 파싱
"""

from dataclasses import dataclass
from typing import List, Optional
import re


@dataclass
class TextDocument:
    """파싱된 텍스트 문서"""
    content: str
    lines: List[str]
    filename: str = ""
    encoding: str = "utf-8"
    
    # Markdown 전용
    headings: List[tuple] = None  # [(level, text), ...]
    code_blocks: List[dict] = None  # [{'lang': 'python', 'code': '...'}, ...]
    links: List[tuple] = None  # [(text, url), ...]
    images: List[tuple] = None  # [(alt, url), ...]


def parse_text(filepath_or_bytes, encoding: str = None) -> TextDocument:
    """
    텍스트 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
        encoding: 인코딩 (자동 감지 시도)
    
    Returns:
        TextDocument
    """
    # 데이터 읽기
    if isinstance(filepath_or_bytes, str):
        data = _read_file(filepath_or_bytes, encoding)
        filename = filepath_or_bytes
    else:
        data = _decode_bytes(filepath_or_bytes, encoding)
        filename = ""
    
    lines = data.splitlines()
    
    return TextDocument(
        content=data,
        lines=lines,
        filename=filename,
        encoding=encoding or "utf-8"
    )


def parse_markdown(filepath_or_bytes, encoding: str = None) -> TextDocument:
    """
    마크다운 파일 파싱
    
    Args:
        filepath_or_bytes: 파일 경로 또는 바이트
        encoding: 인코딩
    
    Returns:
        TextDocument (마크다운 요소 포함)
    """
    doc = parse_text(filepath_or_bytes, encoding)
    
    # 마크다운 요소 추출
    doc.headings = _extract_headings(doc.content)
    doc.code_blocks = _extract_code_blocks(doc.content)
    doc.links = _extract_links(doc.content)
    doc.images = _extract_images(doc.content)
    
    return doc


def _read_file(filepath: str, encoding: str = None) -> str:
    """파일 읽기 (인코딩 자동 감지)"""
    # 인코딩 시도 순서
    encodings = [encoding] if encoding else []
    encodings.extend(['utf-8', 'utf-8-sig', 'cp949', 'euc-kr', 'latin-1'])
    
    with open(filepath, 'rb') as f:
        raw_data = f.read()
    
    return _decode_bytes(raw_data, encoding)


def _decode_bytes(data: bytes, encoding: str = None) -> str:
    """바이트를 문자열로 디코딩 (인코딩 자동 감지)"""
    # BOM 체크
    if data.startswith(b'\xef\xbb\xbf'):
        return data[3:].decode('utf-8')
    if data.startswith(b'\xff\xfe'):
        return data[2:].decode('utf-16-le')
    if data.startswith(b'\xfe\xff'):
        return data[2:].decode('utf-16-be')
    
    # 인코딩 시도
    encodings = [encoding] if encoding else []
    encodings.extend(['utf-8', 'cp949', 'euc-kr', 'latin-1'])
    
    for enc in encodings:
        if not enc:
            continue
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    
    # 최후의 수단
    return data.decode('utf-8', errors='replace')


def _extract_headings(content: str) -> List[tuple]:
    """마크다운 헤딩 추출"""
    headings = []
    
    # ATX 스타일: # Heading
    for match in re.finditer(r'^(#{1,6})\s+(.+)$', content, re.MULTILINE):
        level = len(match.group(1))
        text = match.group(2).strip()
        headings.append((level, text))
    
    return headings


def _extract_code_blocks(content: str) -> List[dict]:
    """코드 블록 추출"""
    blocks = []
    
    # Fenced 코드 블록: ```lang ... ```
    pattern = r'```(\w*)\n(.*?)```'
    for match in re.finditer(pattern, content, re.DOTALL):
        blocks.append({
            'lang': match.group(1) or 'text',
            'code': match.group(2)
        })
    
    return blocks


def _extract_links(content: str) -> List[tuple]:
    """링크 추출 [text](url)"""
    links = []
    
    # 이미지가 아닌 링크만
    pattern = r'(?<!!)\[([^\]]+)\]\(([^)]+)\)'
    for match in re.finditer(pattern, content):
        links.append((match.group(1), match.group(2)))
    
    return links


def _extract_images(content: str) -> List[tuple]:
    """이미지 추출 ![alt](url)"""
    images = []
    
    pattern = r'!\[([^\]]*)\]\(([^)]+)\)'
    for match in re.finditer(pattern, content):
        images.append((match.group(1), match.group(2)))
    
    return images


def extract_text(doc: TextDocument) -> str:
    """순수 텍스트 추출 (마크다운 문법 제거)"""
    text = doc.content
    
    # 코드 블록 제거
    text = re.sub(r'```.*?```', '', text, flags=re.DOTALL)
    
    # 인라인 코드 제거
    text = re.sub(r'`[^`]+`', '', text)
    
    # 이미지 제거
    text = re.sub(r'!\[([^\]]*)\]\([^)]+\)', '', text)
    
    # 링크를 텍스트만 남기기
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    
    # 헤딩 마커 제거
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    
    # 강조 제거
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
    text = re.sub(r'\*([^*]+)\*', r'\1', text)
    text = re.sub(r'__([^_]+)__', r'\1', text)
    text = re.sub(r'_([^_]+)_', r'\1', text)
    
    # 수평선 제거
    text = re.sub(r'^[-*_]{3,}$', '', text, flags=re.MULTILINE)
    
    # 연속 빈 줄 정리
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text.strip()
