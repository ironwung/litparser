"""
PDF Parser - Stage 1: 기본 구조 파싱

목표:
1. PDF 헤더 파싱 (버전 확인)
2. Trailer 파싱 (파일 끝에서부터)
3. XRef 테이블 파싱 (객체 위치 인덱스)
4. 기본 객체 타입 파싱 (dict, array, string, number, name, ref)
"""

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union
from enum import Enum


class PDFTokenType(Enum):
    """PDF 토큰 타입"""
    NUMBER = "number"
    STRING = "string"          # (hello)
    HEX_STRING = "hex_string"  # <48656C6C6F>
    NAME = "name"              # /Type
    BOOL = "bool"
    NULL = "null"
    KEYWORD = "keyword"        # obj, endobj, stream, etc.
    DICT_START = "dict_start"  # <<
    DICT_END = "dict_end"      # >>
    ARRAY_START = "array_start"  # [
    ARRAY_END = "array_end"      # ]
    REF = "ref"                # 1 0 R


@dataclass
class PDFToken:
    """파싱된 토큰"""
    type: PDFTokenType
    value: Any
    pos: int  # 파일 내 위치


@dataclass
class PDFRef:
    """객체 참조 (예: 1 0 R)"""
    obj_num: int
    gen_num: int
    
    def __repr__(self):
        return f"Ref({self.obj_num} {self.gen_num} R)"


@dataclass
class XRefEntry:
    """XRef 테이블 항목"""
    offset: int      # 파일 내 바이트 오프셋 (또는 Object Stream 번호)
    gen_num: int     # 세대 번호 (또는 Object Stream 내 인덱스)
    in_use: bool     # 사용 중 여부 (n=True, f=False)
    compressed: bool = False  # Object Stream 내 압축 객체 여부
    obj_stream_num: int = 0   # 압축 객체가 속한 Object Stream 번호
    obj_stream_idx: int = 0   # Object Stream 내 인덱스


@dataclass 
class PDFDocument:
    """파싱된 PDF 문서"""
    version: str
    xref: Dict[int, XRefEntry] = field(default_factory=dict)
    trailer: Dict[str, Any] = field(default_factory=dict)
    objects: Dict[Tuple[int, int], Any] = field(default_factory=dict)  # (obj_num, gen_num) -> value


class PDFLexer:
    """PDF 토크나이저 - 바이트 스트림을 토큰으로 변환"""
    
    # 구분자 문자
    WHITESPACE = b' \t\n\r\x00\x0c'
    DELIMITERS = b'()<>[]{}/%'
    
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        self.length = len(data)
    
    def peek(self, count: int = 1) -> bytes:
        """현재 위치에서 count 바이트 미리보기"""
        return self.data[self.pos:self.pos + count]
    
    def read(self, count: int = 1) -> bytes:
        """count 바이트 읽고 위치 이동"""
        result = self.data[self.pos:self.pos + count]
        self.pos += count
        return result
    
    def skip_whitespace(self):
        """공백 문자 스킵"""
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos + 1]
            if ch in self.WHITESPACE:
                self.pos += 1
            elif ch == b'%':
                # 주석 스킵 (줄 끝까지)
                while self.pos < self.length and self.data[self.pos:self.pos + 1] not in b'\r\n':
                    self.pos += 1
            else:
                break
    
    def read_token(self) -> Optional[PDFToken]:
        """다음 토큰 읽기"""
        self.skip_whitespace()
        
        if self.pos >= self.length:
            return None
        
        start_pos = self.pos
        ch = self.peek()
        
        # Dictionary 시작/끝
        if ch == b'<':
            if self.peek(2) == b'<<':
                self.pos += 2
                return PDFToken(PDFTokenType.DICT_START, "<<", start_pos)
            else:
                return self._read_hex_string(start_pos)
        
        if ch == b'>':
            if self.peek(2) == b'>>':
                self.pos += 2
                return PDFToken(PDFTokenType.DICT_END, ">>", start_pos)
            else:
                raise ValueError(f"Unexpected '>' at position {self.pos}")
        
        # Array
        if ch == b'[':
            self.pos += 1
            return PDFToken(PDFTokenType.ARRAY_START, "[", start_pos)
        if ch == b']':
            self.pos += 1
            return PDFToken(PDFTokenType.ARRAY_END, "]", start_pos)
        
        # Name
        if ch == b'/':
            return self._read_name(start_pos)
        
        # String
        if ch == b'(':
            return self._read_string(start_pos)
        
        # Number 또는 Keyword
        if ch in b'-+0123456789.':
            return self._read_number_or_keyword(start_pos)
        
        # Keyword (true, false, null, obj, endobj, etc.)
        if ch.isalpha():
            return self._read_keyword(start_pos)
        
        raise ValueError(f"Unexpected character {ch!r} at position {self.pos}")
    
    def _read_name(self, start_pos: int) -> PDFToken:
        """Name 토큰 읽기: /Type, /Pages, etc."""
        self.pos += 1  # '/' 스킵
        name = b''
        
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos + 1]
            if ch in self.WHITESPACE or ch in self.DELIMITERS:
                break
            
            # #XX 이스케이프 처리
            if ch == b'#' and self.pos + 2 < self.length:
                hex_val = self.data[self.pos + 1:self.pos + 3]
                try:
                    name += bytes([int(hex_val, 16)])
                    self.pos += 3
                    continue
                except ValueError:
                    pass
            
            name += ch
            self.pos += 1
        
        return PDFToken(PDFTokenType.NAME, name.decode('utf-8', errors='replace'), start_pos)
    
    def _read_string(self, start_pos: int) -> PDFToken:
        """리터럴 문자열 읽기: (Hello World)"""
        self.pos += 1  # '(' 스킵
        result = b''
        depth = 1  # 괄호 중첩 추적
        
        while self.pos < self.length and depth > 0:
            ch = self.data[self.pos:self.pos + 1]
            
            if ch == b'\\':
                # 이스케이프 시퀀스
                self.pos += 1
                if self.pos >= self.length:
                    break
                esc = self.data[self.pos:self.pos + 1]
                
                escape_map = {
                    b'n': b'\n', b'r': b'\r', b't': b'\t',
                    b'b': b'\b', b'f': b'\f',
                    b'(': b'(', b')': b')', b'\\': b'\\'
                }
                
                if esc in escape_map:
                    result += escape_map[esc]
                    self.pos += 1
                elif esc in b'0123456789':
                    # 8진수 이스케이프
                    octal = b''
                    for _ in range(3):
                        if self.pos < self.length and self.data[self.pos:self.pos + 1] in b'01234567':
                            octal += self.data[self.pos:self.pos + 1]
                            self.pos += 1
                        else:
                            break
                    if octal:
                        result += bytes([int(octal, 8) & 0xFF])
                elif esc in b'\r\n':
                    # 줄 연속
                    if esc == b'\r' and self.peek() == b'\n':
                        self.pos += 1
                    self.pos += 1
                else:
                    result += esc
                    self.pos += 1
            elif ch == b'(':
                depth += 1
                result += ch
                self.pos += 1
            elif ch == b')':
                depth -= 1
                if depth > 0:
                    result += ch
                self.pos += 1
            else:
                result += ch
                self.pos += 1
        
        return PDFToken(PDFTokenType.STRING, result, start_pos)
    
    def _read_hex_string(self, start_pos: int) -> PDFToken:
        """16진수 문자열 읽기: <48656C6C6F>"""
        self.pos += 1  # '<' 스킵
        hex_str = b''
        
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos + 1]
            if ch == b'>':
                self.pos += 1
                break
            if ch in self.WHITESPACE:
                self.pos += 1
                continue
            if ch in b'0123456789ABCDEFabcdef':
                hex_str += ch
                self.pos += 1
            else:
                raise ValueError(f"Invalid hex character {ch!r} at position {self.pos}")
        
        # 홀수 길이면 0 추가
        if len(hex_str) % 2 == 1:
            hex_str += b'0'
        
        result = bytes.fromhex(hex_str.decode('ascii'))
        return PDFToken(PDFTokenType.HEX_STRING, result, start_pos)
    
    def _read_number_or_keyword(self, start_pos: int) -> PDFToken:
        """숫자 또는 키워드 읽기"""
        token = b''
        
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos + 1]
            if ch in self.WHITESPACE or ch in self.DELIMITERS:
                break
            token += ch
            self.pos += 1
        
        token_str = token.decode('ascii')
        
        # 정수 또는 실수 판별
        try:
            if '.' in token_str:
                return PDFToken(PDFTokenType.NUMBER, float(token_str), start_pos)
            else:
                return PDFToken(PDFTokenType.NUMBER, int(token_str), start_pos)
        except ValueError:
            return PDFToken(PDFTokenType.KEYWORD, token_str, start_pos)
    
    def _read_keyword(self, start_pos: int) -> PDFToken:
        """키워드 읽기: true, false, null, obj, endobj, stream, etc."""
        token = b''
        
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos + 1]
            if ch in self.WHITESPACE or ch in self.DELIMITERS:
                break
            token += ch
            self.pos += 1
        
        token_str = token.decode('ascii')
        
        if token_str == 'true':
            return PDFToken(PDFTokenType.BOOL, True, start_pos)
        elif token_str == 'false':
            return PDFToken(PDFTokenType.BOOL, False, start_pos)
        elif token_str == 'null':
            return PDFToken(PDFTokenType.NULL, None, start_pos)
        else:
            return PDFToken(PDFTokenType.KEYWORD, token_str, start_pos)


class PDFParser:
    """PDF 파서 - 토큰을 파싱해서 객체로 변환"""
    
    def __init__(self, data: bytes):
        self.data = data
        self.lexer = PDFLexer(data)
        self.document = PDFDocument(version="")
    
    def parse(self) -> PDFDocument:
        """PDF 문서 전체 파싱"""
        # 1. 헤더 파싱
        self._parse_header()
        
        # 2. Trailer와 XRef 파싱 (파일 끝에서부터)
        self._parse_trailer_and_xref()
        
        # 3. 모든 객체 파싱
        self._parse_all_objects()
        
        return self.document
    
    def _parse_header(self):
        """PDF 헤더 파싱"""
        # %PDF-X.X
        match = re.match(rb'%PDF-(\d+\.\d+)', self.data)
        if not match:
            raise ValueError("Invalid PDF: missing header")
        self.document.version = match.group(1).decode('ascii')
    
    def _parse_trailer_and_xref(self):
        """Trailer와 XRef 테이블 파싱 (파일 끝에서부터)"""
        # %%EOF 찾기
        eof_pos = self.data.rfind(b'%%EOF')
        if eof_pos == -1:
            raise ValueError("Invalid PDF: missing %%EOF")
        
        # startxref 찾기
        startxref_pos = self.data.rfind(b'startxref', 0, eof_pos)
        if startxref_pos == -1:
            raise ValueError("Invalid PDF: missing startxref")
        
        # xref 오프셋 읽기
        after_startxref = self.data[startxref_pos + 9:eof_pos].strip()
        xref_offset = int(after_startxref.split()[0])
        
        # XRef 파싱
        self._parse_xref_at(xref_offset)
    
    def _parse_xref_at(self, offset: int):
        """특정 오프셋에서 XRef 테이블 또는 XRef 스트림 파싱"""
        self.lexer.pos = offset
        self.lexer.skip_whitespace()
        
        # 'xref' 키워드 확인 - 기존 XRef 테이블
        if self.data[offset:offset + 4] == b'xref':
            self._parse_xref_table(offset)
        else:
            # XRef 스트림 (PDF 1.5+)
            self._parse_xref_stream(offset)
    
    def _parse_xref_table(self, offset: int):
        """기존 XRef 테이블 파싱"""
        self.lexer.pos = offset + 4
        self.lexer.skip_whitespace()
        
        # 섹션들 파싱
        while True:
            token = self.lexer.read_token()
            if token is None:
                break
            
            if token.type == PDFTokenType.KEYWORD and token.value == 'trailer':
                break
            
            if token.type != PDFTokenType.NUMBER:
                break
            
            start_obj = token.value
            
            token = self.lexer.read_token()
            if token.type != PDFTokenType.NUMBER:
                break
            
            count = token.value
            
            # 각 항목 파싱
            for i in range(count):
                self.lexer.skip_whitespace()
                
                # 20바이트 고정 형식: OOOOOOOOOO GGGGG n/f
                entry_data = self.data[self.lexer.pos:self.lexer.pos + 20]
                
                if len(entry_data) < 18:
                    break
                
                offset_str = entry_data[0:10].decode('ascii').strip()
                gen_str = entry_data[11:16].decode('ascii').strip()
                flag = entry_data[17:18].decode('ascii')
                
                obj_num = start_obj + i
                
                # 이미 있는 항목은 건너뜀 (최신 XRef 우선)
                if obj_num not in self.document.xref:
                    self.document.xref[obj_num] = XRefEntry(
                        offset=int(offset_str),
                        gen_num=int(gen_str),
                        in_use=(flag == 'n')
                    )
                
                self.lexer.pos += 20
        
        # Trailer dictionary 파싱
        self.lexer.skip_whitespace()
        trailer_dict = self._parse_value()
        
        if isinstance(trailer_dict, dict):
            # 기존 trailer와 병합 (incremental updates)
            if not self.document.trailer:
                self.document.trailer = trailer_dict
            else:
                for k, v in trailer_dict.items():
                    if k not in self.document.trailer:
                        self.document.trailer[k] = v
            
            # Prev가 있으면 이전 xref도 파싱 (incremental updates)
            if 'Prev' in trailer_dict:
                self._parse_xref_at(int(trailer_dict['Prev']))
    
    def _parse_xref_stream(self, offset: int):
        """XRef 스트림 파싱 (PDF 1.5+)"""
        from .stream_decoder import StreamDecoder
        
        self.lexer.pos = offset
        
        # 객체 헤더 파싱: obj_num gen_num obj
        token1 = self.lexer.read_token()
        token2 = self.lexer.read_token()
        token3 = self.lexer.read_token()
        
        if not (token1 and token2 and token3):
            raise ValueError(f"Invalid XRef stream at offset {offset}")
        
        # 딕셔너리 파싱
        xref_dict = self._parse_value()
        
        if not isinstance(xref_dict, dict):
            raise ValueError(f"Expected XRef stream dictionary at offset {offset}")
        
        # stream 키워드 확인 및 스트림 데이터 추출
        self.lexer.skip_whitespace()
        if self.data[self.lexer.pos:self.lexer.pos + 6] == b'stream':
            xref_dict = self._parse_stream(xref_dict)
        
        if '_stream_data' not in xref_dict:
            raise ValueError("XRef stream has no stream data")
        
        # 스트림 디코딩
        filters = xref_dict.get('Filter', [])
        decode_parms = xref_dict.get('DecodeParms', {})
        if filters:
            stream_data = StreamDecoder.decode(xref_dict['_stream_data'], filters, decode_parms)
        else:
            stream_data = xref_dict['_stream_data']
        
        # XRef 스트림 파라미터
        w = xref_dict.get('W', [1, 2, 1])  # 기본값: [타입 1바이트, 오프셋 2바이트, gen 1바이트]
        size = xref_dict.get('Size', 0)
        index = xref_dict.get('Index', [0, size])  # 기본값: 0부터 Size개
        
        # W 배열: [타입 필드 크기, 값1 크기, 값2 크기]
        w0, w1, w2 = w[0], w[1], w[2] if len(w) > 2 else 0
        entry_size = w0 + w1 + w2
        
        # Index 배열 파싱: [시작1, 개수1, 시작2, 개수2, ...]
        subsections = []
        for i in range(0, len(index), 2):
            start = index[i]
            count = index[i + 1] if i + 1 < len(index) else 0
            subsections.append((start, count))
        
        # 스트림 데이터 파싱
        pos = 0
        for start_obj, count in subsections:
            for i in range(count):
                if pos + entry_size > len(stream_data):
                    break
                
                # 각 필드 읽기
                field0 = self._read_bytes_as_int(stream_data, pos, w0) if w0 > 0 else 1
                field1 = self._read_bytes_as_int(stream_data, pos + w0, w1) if w1 > 0 else 0
                field2 = self._read_bytes_as_int(stream_data, pos + w0 + w1, w2) if w2 > 0 else 0
                
                obj_num = start_obj + i
                
                # 이미 있는 항목은 건너뜀 (최신 XRef 우선)
                if obj_num in self.document.xref:
                    pos += entry_size
                    continue
                
                if field0 == 0:
                    # Free 객체
                    self.document.xref[obj_num] = XRefEntry(
                        offset=field1,
                        gen_num=field2,
                        in_use=False
                    )
                elif field0 == 1:
                    # 사용중 객체 (일반)
                    self.document.xref[obj_num] = XRefEntry(
                        offset=field1,
                        gen_num=field2,
                        in_use=True
                    )
                elif field0 == 2:
                    # 압축 객체 (Object Stream 내)
                    # field1 = Object Stream 번호, field2 = 스트림 내 인덱스
                    self.document.xref[obj_num] = XRefEntry(
                        offset=field1,  # Object Stream 번호로 사용
                        gen_num=field2,  # 인덱스로 사용
                        in_use=True,
                        compressed=True,  # 압축 객체 표시
                        obj_stream_num=field1,
                        obj_stream_idx=field2
                    )
                
                pos += entry_size
        
        # Trailer 정보 추출 (XRef 스트림 딕셔너리에서)
        if not self.document.trailer:
            self.document.trailer = {}
        
        for key in ['Root', 'Info', 'ID', 'Size', 'Encrypt']:
            if key in xref_dict and key not in self.document.trailer:
                self.document.trailer[key] = xref_dict[key]
        
        # Prev가 있으면 이전 xref도 파싱
        if 'Prev' in xref_dict:
            self._parse_xref_at(int(xref_dict['Prev']))
    
    def _read_bytes_as_int(self, data: bytes, offset: int, length: int) -> int:
        """바이트 시퀀스를 big-endian 정수로 변환"""
        if length == 0:
            return 0
        value = 0
        for i in range(length):
            if offset + i < len(data):
                value = (value << 8) | data[offset + i]
        return value
    
    def _parse_all_objects(self):
        """XRef에 등록된 모든 객체 파싱"""
        # Object Stream 캐시 (한 번 파싱하면 재사용)
        obj_stream_cache = {}
        
        # 먼저 일반 객체 파싱 (압축되지 않은 것들)
        for obj_num, entry in self.document.xref.items():
            if not entry.in_use:
                continue
            
            if entry.compressed:
                # 압축 객체는 나중에 처리
                continue
            
            try:
                obj_value = self._parse_object_at(entry.offset, obj_num, entry.gen_num)
                self.document.objects[(obj_num, entry.gen_num)] = obj_value
            except Exception as e:
                # 조용히 실패 (디버그 시 출력 가능)
                pass
        
        # 압축 객체 파싱 (Object Stream에서)
        for obj_num, entry in self.document.xref.items():
            if not entry.in_use or not entry.compressed:
                continue
            
            try:
                stream_num = entry.obj_stream_num
                stream_idx = entry.obj_stream_idx
                
                # Object Stream 캐시 확인
                if stream_num not in obj_stream_cache:
                    # Object Stream 파싱
                    stream_obj = self.document.objects.get((stream_num, 0))
                    if stream_obj and isinstance(stream_obj, dict):
                        parsed = self._parse_object_stream(stream_obj)
                        obj_stream_cache[stream_num] = parsed
                    else:
                        continue
                
                # 캐시에서 객체 가져오기
                if stream_num in obj_stream_cache:
                    stream_objects = obj_stream_cache[stream_num]
                    if stream_idx < len(stream_objects):
                        obj_num_in_stream, obj_value = stream_objects[stream_idx]
                        self.document.objects[(obj_num, 0)] = obj_value
            except Exception as e:
                pass
    
    def _parse_object_stream(self, stream_obj: dict) -> list:
        """
        Object Stream (ObjStm) 파싱
        
        Returns:
            list of (obj_num, obj_value) 튜플
        """
        from .stream_decoder import StreamDecoder
        
        if stream_obj.get('Type') != 'ObjStm':
            return []
        
        n = stream_obj.get('N', 0)  # 객체 수
        first = stream_obj.get('First', 0)  # 첫 객체 데이터 시작 위치
        
        if '_stream_data' not in stream_obj:
            return []
        
        # 스트림 디코딩
        filters = stream_obj.get('Filter', [])
        if filters:
            stream_data = StreamDecoder.decode(stream_obj['_stream_data'], filters)
        else:
            stream_data = stream_obj['_stream_data']
        
        # 헤더 파싱: obj_num1 offset1 obj_num2 offset2 ...
        header_data = stream_data[:first]
        
        # 헤더에서 객체번호와 오프셋 쌍 추출
        header_lexer = PDFLexer(header_data)
        obj_info = []
        
        for _ in range(n):
            num_token = header_lexer.read_token()
            offset_token = header_lexer.read_token()
            
            if num_token and offset_token:
                obj_info.append((int(num_token.value), int(offset_token.value)))
        
        # 각 객체 파싱
        results = []
        body_data = stream_data[first:]
        
        for i, (obj_num, rel_offset) in enumerate(obj_info):
            try:
                # 다음 객체까지의 범위 계산
                if i + 1 < len(obj_info):
                    next_offset = obj_info[i + 1][1]
                    obj_data = body_data[rel_offset:next_offset]
                else:
                    obj_data = body_data[rel_offset:]
                
                # 객체 값 파싱
                obj_lexer = PDFLexer(obj_data)
                temp_parser = PDFParser.__new__(PDFParser)
                temp_parser.data = obj_data
                temp_parser.lexer = obj_lexer
                temp_parser.document = PDFDocument(version="")
                
                obj_value = temp_parser._parse_value()
                results.append((obj_num, obj_value))
            except Exception as e:
                results.append((obj_num, None))
        
        return results
    
    def _parse_object_at(self, offset: int, expected_obj: int, expected_gen: int) -> Any:
        """특정 오프셋에서 객체 파싱"""
        self.lexer.pos = offset
        
        # obj_num gen_num obj
        token = self.lexer.read_token()
        if token.type != PDFTokenType.NUMBER or token.value != expected_obj:
            raise ValueError(f"Expected object {expected_obj}, got {token.value}")
        
        token = self.lexer.read_token()
        if token.type != PDFTokenType.NUMBER or token.value != expected_gen:
            raise ValueError(f"Expected generation {expected_gen}, got {token.value}")
        
        token = self.lexer.read_token()
        if token.type != PDFTokenType.KEYWORD or token.value != 'obj':
            raise ValueError(f"Expected 'obj', got {token.value}")
        
        # 객체 값 파싱
        value = self._parse_value()
        
        # stream 체크
        self.lexer.skip_whitespace()
        if self.data[self.lexer.pos:self.lexer.pos + 6] == b'stream':
            value = self._parse_stream(value)
        
        return value
    
    def _parse_value(self) -> Any:
        """값 파싱 (재귀적)"""
        token = self.lexer.read_token()
        if token is None:
            return None
        
        # Dictionary
        if token.type == PDFTokenType.DICT_START:
            return self._parse_dict()
        
        # Array
        if token.type == PDFTokenType.ARRAY_START:
            return self._parse_array()
        
        # 기본 타입들
        if token.type in (PDFTokenType.NUMBER, PDFTokenType.STRING, 
                          PDFTokenType.HEX_STRING, PDFTokenType.NAME,
                          PDFTokenType.BOOL, PDFTokenType.NULL):
            
            # Reference 체크: number number R
            if token.type == PDFTokenType.NUMBER:
                saved_pos = self.lexer.pos
                token2 = self.lexer.read_token()
                
                if token2 and token2.type == PDFTokenType.NUMBER:
                    token3 = self.lexer.read_token()
                    
                    if token3 and token3.type == PDFTokenType.KEYWORD and token3.value == 'R':
                        return PDFRef(int(token.value), int(token2.value))
                
                # Reference가 아니면 위치 복원
                self.lexer.pos = saved_pos
            
            return token.value
        
        if token.type == PDFTokenType.KEYWORD:
            return token.value
        
        raise ValueError(f"Unexpected token: {token}")
    
    def _parse_dict(self) -> Dict[str, Any]:
        """Dictionary 파싱"""
        result = {}
        
        while True:
            token = self.lexer.read_token()
            
            if token is None:
                raise ValueError("Unexpected end of file in dictionary")
            
            if token.type == PDFTokenType.DICT_END:
                break
            
            if token.type != PDFTokenType.NAME:
                raise ValueError(f"Expected name in dictionary, got {token}")
            
            key = token.value
            value = self._parse_value()
            result[key] = value
        
        return result
    
    def _parse_array(self) -> List[Any]:
        """Array 파싱"""
        result = []
        
        while True:
            # 다음 토큰 미리보기
            saved_pos = self.lexer.pos
            token = self.lexer.read_token()
            
            if token is None:
                raise ValueError("Unexpected end of file in array")
            
            if token.type == PDFTokenType.ARRAY_END:
                break
            
            # 위치 복원하고 값 파싱
            self.lexer.pos = saved_pos
            value = self._parse_value()
            result.append(value)
        
        return result
    
    def _parse_stream(self, stream_dict: Dict[str, Any]) -> Dict[str, Any]:
        """Stream 파싱"""
        # 'stream' 키워드 스킵
        self.lexer.pos += 6
        
        # stream 뒤의 EOL 스킵 (\r\n 또는 \n)
        if self.data[self.lexer.pos:self.lexer.pos + 1] == b'\r':
            self.lexer.pos += 1
        if self.data[self.lexer.pos:self.lexer.pos + 1] == b'\n':
            self.lexer.pos += 1
        
        # Length 가져오기
        length = stream_dict.get('Length', 0)
        
        # Length가 참조인 경우
        if isinstance(length, PDFRef):
            # 참조된 객체에서 실제 길이 가져오기
            ref_key = (length.obj_num, length.gen_num)
            if ref_key in self.document.objects:
                length = self.document.objects[ref_key]
            else:
                # 아직 파싱 안됨 - endstream으로 찾기
                endstream_pos = self.data.find(b'endstream', self.lexer.pos)
                if endstream_pos != -1:
                    length = endstream_pos - self.lexer.pos
                    # 끝의 EOL 제거
                    while length > 0 and self.data[self.lexer.pos + length - 1:self.lexer.pos + length] in b'\r\n':
                        length -= 1
                else:
                    length = 0
        
        # Stream 데이터 읽기
        stream_data = self.data[self.lexer.pos:self.lexer.pos + length]
        self.lexer.pos += length
        
        # Stream dict에 raw 데이터 추가
        stream_dict['_stream_data'] = stream_data
        
        return stream_dict


def parse_pdf(filepath: str) -> PDFDocument:
    """PDF 파일 파싱"""
    with open(filepath, 'rb') as f:
        data = f.read()
    
    parser = PDFParser(data)
    return parser.parse()


# 테스트
if __name__ == '__main__':
    import sys
    
    # 간단한 테스트 PDF 생성
    test_pdf = b"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792]
   /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length 44 >>
stream
BT
/F1 24 Tf
100 700 Td
(Hello PDF!) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000244 00000 n 
0000000336 00000 n 
trailer
<< /Size 6 /Root 1 0 R >>
startxref
406
%%EOF"""
    
    print("=" * 60)
    print("PDF Parser Stage 1 - 테스트")
    print("=" * 60)
    
    parser = PDFParser(test_pdf)
    doc = parser.parse()
    
    print(f"\n[헤더]")
    print(f"  PDF 버전: {doc.version}")
    
    print(f"\n[XRef 테이블]")
    for obj_num, entry in sorted(doc.xref.items()):
        status = "사용중" if entry.in_use else "free"
        print(f"  객체 {obj_num}: offset={entry.offset}, gen={entry.gen_num}, {status}")
    
    print(f"\n[Trailer]")
    for key, value in doc.trailer.items():
        print(f"  {key}: {value}")
    
    print(f"\n[파싱된 객체들]")
    for (obj_num, gen_num), value in sorted(doc.objects.items()):
        print(f"\n  === 객체 {obj_num} {gen_num} ===")
        if isinstance(value, dict):
            for k, v in value.items():
                if k == '_stream_data':
                    print(f"    {k}: <{len(v)} bytes>")
                    if len(v) < 200:
                        try:
                            print(f"    [내용] {v.decode('latin-1')}")
                        except:
                            print(f"    [내용] (바이너리)")
                else:
                    print(f"    {k}: {v}")
        else:
            print(f"    {value}")