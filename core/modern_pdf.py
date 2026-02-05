"""
PDF Modern Features - XRef Stream & Object Stream Parser

PDF 1.5+ 에서 도입된 기능들:
1. XRef Stream: 기존 텍스트 XRef 테이블을 압축 바이너리로 대체
2. Object Stream: 여러 객체를 하나의 압축 스트림에 저장

XRef Stream 구조:
- Type: XRef
- Size: 전체 객체 수
- W: [w1, w2, w3] - 각 필드의 바이트 수
- Index: [start1, count1, start2, count2, ...] - 객체 범위
- 스트림 데이터: 각 항목이 w1+w2+w3 바이트

XRef 항목 타입:
- Type 0: Free object (다음 free 객체 번호, generation)
- Type 1: Uncompressed object (byte offset, generation)
- Type 2: Compressed object (object stream 번호, index within stream)

Object Stream 구조:
- Type: ObjStm
- N: 스트림 내 객체 수
- First: 첫 객체 데이터의 오프셋
- 스트림 데이터: "obj1_num obj1_offset obj2_num obj2_offset ... obj1_data obj2_data ..."
"""

from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Any
import zlib


@dataclass
class XRefStreamEntry:
    """XRef 스트림 항목"""
    entry_type: int  # 0=free, 1=uncompressed, 2=compressed
    field2: int      # type에 따라 다른 의미
    field3: int      # type에 따라 다른 의미
    
    @property
    def is_free(self) -> bool:
        return self.entry_type == 0
    
    @property
    def is_uncompressed(self) -> bool:
        return self.entry_type == 1
    
    @property
    def is_compressed(self) -> bool:
        return self.entry_type == 2
    
    @property
    def offset(self) -> int:
        """Type 1: 바이트 오프셋"""
        return self.field2 if self.entry_type == 1 else 0
    
    @property
    def generation(self) -> int:
        """Type 0, 1: generation number"""
        return self.field3 if self.entry_type in (0, 1) else 0
    
    @property
    def objstm_num(self) -> int:
        """Type 2: Object Stream 객체 번호"""
        return self.field2 if self.entry_type == 2 else 0
    
    @property
    def objstm_index(self) -> int:
        """Type 2: Object Stream 내 인덱스"""
        return self.field3 if self.entry_type == 2 else 0


def parse_xref_stream(stream_dict: dict, stream_data: bytes) -> Dict[int, XRefStreamEntry]:
    """
    XRef 스트림 파싱
    
    Args:
        stream_dict: 스트림 딕셔너리 (W, Index, Size 등 포함)
        stream_data: 디코딩된 스트림 데이터
    
    Returns:
        Dict[int, XRefStreamEntry]: 객체 번호 → XRef 항목
    """
    # W 배열: 각 필드의 바이트 수 [type_bytes, field2_bytes, field3_bytes]
    w = stream_dict.get('W', [1, 2, 1])
    if len(w) != 3:
        raise ValueError(f"Invalid W array: {w}")
    
    w1, w2, w3 = w
    entry_size = w1 + w2 + w3
    
    # Index 배열: [start1, count1, start2, count2, ...]
    # 없으면 [0, Size] 로 가정
    size = stream_dict.get('Size', 0)
    index = stream_dict.get('Index', [0, size])
    
    # Index가 리스트가 아니면 변환
    if not isinstance(index, list):
        index = [0, size]
    
    result = {}
    data_pos = 0
    
    # Index 배열 순회 (start, count 쌍)
    i = 0
    while i < len(index) - 1:
        start_obj = index[i]
        count = index[i + 1]
        i += 2
        
        for j in range(count):
            obj_num = start_obj + j
            
            if data_pos + entry_size > len(stream_data):
                break
            
            # 각 필드 읽기 (big-endian)
            entry_type = _read_int(stream_data, data_pos, w1) if w1 > 0 else 1
            field2 = _read_int(stream_data, data_pos + w1, w2) if w2 > 0 else 0
            field3 = _read_int(stream_data, data_pos + w1 + w2, w3) if w3 > 0 else 0
            
            result[obj_num] = XRefStreamEntry(entry_type, field2, field3)
            data_pos += entry_size
    
    return result


def _read_int(data: bytes, offset: int, length: int) -> int:
    """Big-endian 정수 읽기"""
    if length == 0:
        return 0
    value = 0
    for i in range(length):
        if offset + i < len(data):
            value = (value << 8) | data[offset + i]
    return value


def parse_object_stream(stream_dict: dict, stream_data: bytes, lexer_class, parser_instance) -> Dict[int, Any]:
    """
    Object Stream 파싱
    
    Args:
        stream_dict: 스트림 딕셔너리 (N, First 포함)
        stream_data: 디코딩된 스트림 데이터
        lexer_class: PDFLexer 클래스
        parser_instance: PDFParser 인스턴스 (값 파싱용)
    
    Returns:
        Dict[int, Any]: 객체 번호 → 객체 값
    """
    n = stream_dict.get('N', 0)  # 객체 수
    first = stream_dict.get('First', 0)  # 첫 객체 데이터 오프셋
    
    if n == 0:
        return {}
    
    # 헤더 부분 파싱: "obj1_num obj1_offset obj2_num obj2_offset ..."
    header_data = stream_data[:first]
    
    # 숫자들 추출
    header_text = header_data.decode('latin-1', errors='replace')
    numbers = []
    for part in header_text.split():
        try:
            numbers.append(int(part))
        except ValueError:
            pass
    
    # (obj_num, offset) 쌍으로 분리
    obj_entries = []
    for i in range(0, min(len(numbers), n * 2), 2):
        if i + 1 < len(numbers):
            obj_num = numbers[i]
            offset = numbers[i + 1]
            obj_entries.append((obj_num, offset))
    
    # 각 객체 파싱
    result = {}
    object_data = stream_data[first:]
    
    for i, (obj_num, offset) in enumerate(obj_entries):
        # 다음 객체까지의 범위 계산
        if i + 1 < len(obj_entries):
            next_offset = obj_entries[i + 1][1]
            obj_bytes = object_data[offset:next_offset]
        else:
            obj_bytes = object_data[offset:]
        
        # 객체 값 파싱
        try:
            # 임시 렉서로 파싱
            temp_lexer = lexer_class(obj_bytes)
            parser_instance.lexer = temp_lexer
            value = parser_instance._parse_value()
            result[obj_num] = value
        except Exception as e:
            # 파싱 실패 시 무시
            pass
    
    return result


class ModernPDFParser:
    """
    PDF 1.5+ 기능을 지원하는 파서 확장
    
    기존 PDFParser를 확장하여:
    1. XRef 스트림 지원
    2. Object 스트림 지원
    3. Incremental updates 지원
    """
    
    def __init__(self, base_parser):
        """
        Args:
            base_parser: 기존 PDFParser 인스턴스
        """
        self.base_parser = base_parser
        self.data = base_parser.data
        self.document = base_parser.document
        
        # Object Stream 캐시
        self._objstm_cache: Dict[int, Dict[int, Any]] = {}
    
    def parse_xref_stream_at(self, offset: int) -> Tuple[Dict[int, XRefStreamEntry], dict]:
        """
        특정 위치의 XRef 스트림 파싱
        
        Returns:
            (xref_entries, trailer_dict)
        """
        from .parser import PDFLexer
        from .stream_decoder import StreamDecoder
        
        # 객체 파싱
        self.base_parser.lexer.pos = offset
        
        # "obj_num gen_num obj" 읽기
        tokens = []
        for _ in range(3):
            token = self.base_parser.lexer.read_token()
            if token:
                tokens.append(token)
        
        # 스트림 딕셔너리 파싱
        stream_dict = self.base_parser._parse_value()
        
        if not isinstance(stream_dict, dict):
            raise ValueError("XRef stream: expected dictionary")
        
        # 스트림 타입 확인
        if stream_dict.get('Type') != 'XRef':
            raise ValueError(f"Not an XRef stream: Type={stream_dict.get('Type')}")
        
        # 스트림 데이터 파싱
        self.base_parser.lexer.skip_whitespace()
        
        if self.data[self.base_parser.lexer.pos:self.base_parser.lexer.pos + 6] != b'stream':
            raise ValueError("XRef stream: missing stream keyword")
        
        stream_dict = self.base_parser._parse_stream(stream_dict)
        
        # 스트림 디코딩
        raw_data = stream_dict.get('_stream_data', b'')
        filters = stream_dict.get('Filter', [])
        
        if filters:
            decoded_data = StreamDecoder.decode(raw_data, filters)
        else:
            decoded_data = raw_data
        
        # XRef 항목 파싱
        xref_entries = parse_xref_stream(stream_dict, decoded_data)
        
        # trailer 정보 추출 (XRef 스트림 딕셔너리에 포함됨)
        trailer = {
            'Size': stream_dict.get('Size'),
            'Root': stream_dict.get('Root'),
            'Info': stream_dict.get('Info'),
            'Prev': stream_dict.get('Prev'),
            'Encrypt': stream_dict.get('Encrypt'),
        }
        # None 값 제거
        trailer = {k: v for k, v in trailer.items() if v is not None}
        
        return xref_entries, trailer
    
    def get_object_from_stream(self, objstm_num: int, index: int) -> Any:
        """
        Object Stream에서 객체 가져오기
        
        Args:
            objstm_num: Object Stream 객체 번호
            index: 스트림 내 인덱스
        
        Returns:
            파싱된 객체 값
        """
        # 캐시 확인
        if objstm_num not in self._objstm_cache:
            self._load_object_stream(objstm_num)
        
        cache = self._objstm_cache.get(objstm_num, {})
        
        # index로 검색 (cache는 obj_num → value 매핑)
        # index 순서대로 저장되어 있으므로, 인덱스 기반 검색 필요
        # 현재 구현에서는 obj_num을 직접 반환
        for obj_num, value in cache.items():
            if index == 0:
                return value
            index -= 1
        
        return None
    
    def _load_object_stream(self, objstm_num: int):
        """Object Stream 로드 및 캐시"""
        from .parser import PDFLexer
        from .stream_decoder import StreamDecoder
        
        # Object Stream 객체 가져오기
        objstm = self.document.objects.get((objstm_num, 0))
        
        if not objstm or not isinstance(objstm, dict):
            self._objstm_cache[objstm_num] = {}
            return
        
        if objstm.get('Type') != 'ObjStm':
            self._objstm_cache[objstm_num] = {}
            return
        
        # 스트림 디코딩
        raw_data = objstm.get('_stream_data', b'')
        filters = objstm.get('Filter', [])
        
        if filters:
            decoded_data = StreamDecoder.decode(raw_data, filters)
        else:
            decoded_data = raw_data
        
        # Object Stream 파싱
        objects = parse_object_stream(
            objstm, decoded_data,
            PDFLexer, self.base_parser
        )
        
        self._objstm_cache[objstm_num] = objects


def find_all_xref_positions(data: bytes) -> List[int]:
    """
    PDF 파일에서 모든 XRef 위치 찾기 (Incremental updates 지원)
    
    Returns:
        List[int]: XRef 시작 오프셋 리스트 (최신 → 오래된 순)
    """
    positions = []
    
    # 모든 %%EOF 찾기
    pos = len(data)
    while True:
        eof_pos = data.rfind(b'%%EOF', 0, pos)
        if eof_pos == -1:
            break
        
        # startxref 찾기
        startxref_pos = data.rfind(b'startxref', 0, eof_pos)
        if startxref_pos != -1:
            # startxref 다음의 숫자 추출
            after = data[startxref_pos + 9:eof_pos].strip()
            try:
                xref_offset = int(after.split()[0])
                positions.append(xref_offset)
            except (ValueError, IndexError):
                pass
        
        pos = eof_pos
    
    return positions


def is_xref_stream(data: bytes, offset: int) -> bool:
    """해당 위치가 XRef 스트림인지 확인"""
    # 'xref' 키워드로 시작하면 기존 XRef 테이블
    # 숫자로 시작하면 XRef 스트림 (객체)
    
    # 앞쪽 공백 스킵
    while offset < len(data) and data[offset:offset+1] in b' \t\n\r':
        offset += 1
    
    if offset >= len(data):
        return False
    
    # 'xref'로 시작하면 기존 테이블
    if data[offset:offset+4] == b'xref':
        return False
    
    # 숫자로 시작하면 XRef 스트림
    return data[offset:offset+1] in b'0123456789'


# 테스트
if __name__ == '__main__':
    # XRef 스트림 파싱 테스트
    # W = [1, 2, 1], 항목: type=1, offset=1000, gen=0
    test_w = [1, 2, 1]
    test_data = bytes([
        1, 0x03, 0xE8, 0,  # type=1, offset=1000, gen=0
        1, 0x07, 0xD0, 0,  # type=1, offset=2000, gen=0
        2, 0x00, 0x05, 3,  # type=2, objstm=5, index=3
    ])
    
    test_dict = {'W': test_w, 'Size': 3, 'Index': [0, 3]}
    
    entries = parse_xref_stream(test_dict, test_data)
    
    print("XRef Stream 테스트:")
    for obj_num, entry in entries.items():
        print(f"  객체 {obj_num}: type={entry.entry_type}, "
              f"field2={entry.field2}, field3={entry.field3}")
        if entry.is_uncompressed:
            print(f"    → 오프셋 {entry.offset}, gen {entry.generation}")
        elif entry.is_compressed:
            print(f"    → ObjStm {entry.objstm_num}, index {entry.objstm_index}")
