"""
PDF Content Stream Parser

Content Stream의 연산자들을 파싱하고 텍스트 상태를 추적

주요 연산자:
- BT/ET: 텍스트 블록 시작/끝
- Tf: 폰트 설정
- Td, TD, Tm, T*: 위치 이동
- Tj, TJ, ', ": 텍스트 출력
- Tc, Tw, TL, Tz: 텍스트 속성
"""

import re
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Any, Dict
from enum import Enum


class CSTokenType(Enum):
    """Content Stream 토큰 타입"""
    NUMBER = "number"
    STRING = "string"       # (Hello)
    HEX_STRING = "hex"      # <48656C6C6F>
    NAME = "name"           # /F1
    OPERATOR = "operator"   # Tj, BT, ET
    ARRAY_START = "["
    ARRAY_END = "]"


@dataclass
class CSToken:
    type: CSTokenType
    value: Any
    pos: int = 0


@dataclass
class TextItem:
    """추출된 텍스트 항목"""
    text: str               # 실제 텍스트
    x: float               # X 좌표
    y: float               # Y 좌표
    font_name: str         # 폰트 이름
    font_size: float       # 폰트 크기
    char_spacing: float = 0
    word_spacing: float = 0


@dataclass
class TextState:
    """텍스트 상태 머신"""
    # 폰트
    font_name: str = ""
    font_size: float = 12.0
    
    # 위치 (텍스트 매트릭스)
    tm: List[float] = field(default_factory=lambda: [1, 0, 0, 1, 0, 0])
    
    # 줄 매트릭스 (T* 연산자용)
    tlm: List[float] = field(default_factory=lambda: [1, 0, 0, 1, 0, 0])
    
    # 텍스트 속성
    char_spacing: float = 0      # Tc
    word_spacing: float = 0      # Tw
    leading: float = 0           # TL (줄 간격)
    horizontal_scale: float = 100  # Tz (%)
    rise: float = 0              # Ts (베이스라인)
    
    def get_position(self) -> Tuple[float, float]:
        """현재 텍스트 위치 반환"""
        return (self.tm[4], self.tm[5])
    
    def copy(self) -> 'TextState':
        """상태 복사"""
        new_state = TextState()
        new_state.font_name = self.font_name
        new_state.font_size = self.font_size
        new_state.tm = self.tm.copy()
        new_state.tlm = self.tlm.copy()
        new_state.char_spacing = self.char_spacing
        new_state.word_spacing = self.word_spacing
        new_state.leading = self.leading
        new_state.horizontal_scale = self.horizontal_scale
        new_state.rise = self.rise
        return new_state


class ContentStreamLexer:
    """Content Stream 토크나이저"""
    
    WHITESPACE = b' \t\n\r\x00'
    
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        self.length = len(data)
    
    def tokenize(self) -> List[CSToken]:
        """전체 토큰화"""
        tokens = []
        while self.pos < self.length:
            token = self._read_token()
            if token:
                tokens.append(token)
        return tokens
    
    def _skip_whitespace(self):
        while self.pos < self.length and self.data[self.pos:self.pos+1] in self.WHITESPACE:
            self.pos += 1
    
    def _read_token(self) -> Optional[CSToken]:
        self._skip_whitespace()
        if self.pos >= self.length:
            return None
        
        start_pos = self.pos
        ch = self.data[self.pos:self.pos+1]
        
        # 주석 스킵
        if ch == b'%':
            while self.pos < self.length and self.data[self.pos:self.pos+1] not in b'\r\n':
                self.pos += 1
            return self._read_token()
        
        # 배열
        if ch == b'[':
            self.pos += 1
            return CSToken(CSTokenType.ARRAY_START, '[', start_pos)
        if ch == b']':
            self.pos += 1
            return CSToken(CSTokenType.ARRAY_END, ']', start_pos)
        
        # 이름 (/F1)
        if ch == b'/':
            return self._read_name(start_pos)
        
        # 문자열 (Hello)
        if ch == b'(':
            return self._read_string(start_pos)
        
        # Hex 문자열 <48656C6C6F>
        if ch == b'<':
            return self._read_hex_string(start_pos)
        
        # 숫자
        if ch in b'-+.0123456789':
            return self._read_number(start_pos)
        
        # 연산자 (알파벳)
        if ch.isalpha() or ch == b"'" or ch == b'"':
            return self._read_operator(start_pos)
        
        # 알 수 없는 문자 스킵
        self.pos += 1
        return self._read_token()
    
    def _read_name(self, start_pos: int) -> CSToken:
        self.pos += 1  # '/' 스킵
        name = b''
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos+1]
            if ch in self.WHITESPACE or ch in b'()<>[]{}/%':
                break
            name += ch
            self.pos += 1
        return CSToken(CSTokenType.NAME, name.decode('latin-1'), start_pos)
    
    def _read_string(self, start_pos: int) -> CSToken:
        self.pos += 1  # '(' 스킵
        result = b''
        depth = 1
        
        while self.pos < self.length and depth > 0:
            ch = self.data[self.pos:self.pos+1]
            
            if ch == b'\\':
                self.pos += 1
                if self.pos >= self.length:
                    break
                esc = self.data[self.pos:self.pos+1]
                
                escape_map = {
                    b'n': b'\n', b'r': b'\r', b't': b'\t',
                    b'b': b'\b', b'f': b'\f',
                    b'(': b'(', b')': b')', b'\\': b'\\'
                }
                
                if esc in escape_map:
                    result += escape_map[esc]
                    self.pos += 1
                elif esc in b'0123456789':
                    # 8진수
                    octal = b''
                    for _ in range(3):
                        if self.pos < self.length and self.data[self.pos:self.pos+1] in b'01234567':
                            octal += self.data[self.pos:self.pos+1]
                            self.pos += 1
                        else:
                            break
                    if octal:
                        result += bytes([int(octal, 8) & 0xFF])
                else:
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
        
        return CSToken(CSTokenType.STRING, result, start_pos)
    
    def _read_hex_string(self, start_pos: int) -> CSToken:
        self.pos += 1  # '<' 스킵
        hex_str = b''
        
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos+1]
            if ch == b'>':
                self.pos += 1
                break
            if ch in self.WHITESPACE:
                self.pos += 1
                continue
            if ch in b'0123456789ABCDEFabcdef':
                hex_str += ch
            self.pos += 1
        
        # 홀수 길이면 0 추가
        if len(hex_str) % 2 == 1:
            hex_str += b'0'
        
        try:
            result = bytes.fromhex(hex_str.decode('ascii'))
        except:
            result = b''
        
        return CSToken(CSTokenType.HEX_STRING, result, start_pos)
    
    def _read_number(self, start_pos: int) -> CSToken:
        num_str = b''
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos+1]
            if ch in b'-+.0123456789':
                num_str += ch
                self.pos += 1
            else:
                break
        
        try:
            if b'.' in num_str:
                return CSToken(CSTokenType.NUMBER, float(num_str), start_pos)
            else:
                return CSToken(CSTokenType.NUMBER, int(num_str), start_pos)
        except:
            return CSToken(CSTokenType.NUMBER, 0, start_pos)
    
    def _read_operator(self, start_pos: int) -> CSToken:
        op = b''
        while self.pos < self.length:
            ch = self.data[self.pos:self.pos+1]
            # 연산자는 알파벳, *, ', "
            if ch.isalpha() or ch in b"*'\"":
                op += ch
                self.pos += 1
            else:
                break
        return CSToken(CSTokenType.OPERATOR, op.decode('latin-1'), start_pos)


class ContentStreamParser:
    """Content Stream 파서 - 텍스트 추출"""
    
    def __init__(self, font_map: Dict[str, 'FontInfo'] = None):
        """
        Args:
            font_map: 폰트 이름 → FontInfo 매핑 (ToUnicode 등)
        """
        self.font_map = font_map or {}
        self.text_items: List[TextItem] = []
        self.state = TextState()
        self.state_stack: List[TextState] = []
    
    def parse(self, data: bytes) -> List[TextItem]:
        """Content Stream 파싱해서 텍스트 항목 추출"""
        lexer = ContentStreamLexer(data)
        tokens = lexer.tokenize()
        
        self.text_items = []
        self.state = TextState()
        self.state_stack = []
        
        # 스택 기반 파싱 (피연산자 → 연산자)
        operand_stack = []
        i = 0
        
        while i < len(tokens):
            token = tokens[i]
            
            if token.type == CSTokenType.OPERATOR:
                self._execute_operator(token.value, operand_stack)
                operand_stack = []
            elif token.type == CSTokenType.ARRAY_START:
                # 배열 시작 - 배열 끝까지 모아서 하나의 피연산자로
                array, end_idx = self._collect_array(tokens, i + 1)
                operand_stack.append(array)
                i = end_idx  # 배열 끝으로 점프
            elif token.type != CSTokenType.ARRAY_END:
                operand_stack.append(token)
            
            i += 1
        
        return self.text_items
    
    def _collect_array(self, tokens: List[CSToken], start_idx: int) -> Tuple[List[Any], int]:
        """배열 내용 수집, (배열, 끝 인덱스) 반환"""
        array = []
        depth = 1
        idx = start_idx
        
        while idx < len(tokens) and depth > 0:
            token = tokens[idx]
            if token.type == CSTokenType.ARRAY_START:
                depth += 1
            elif token.type == CSTokenType.ARRAY_END:
                depth -= 1
                if depth == 0:
                    break
            else:
                array.append(token)
            idx += 1
        
        return array, idx
    
    def _execute_operator(self, op: str, operands: List):
        """연산자 실행"""
        
        # 그래픽 상태
        if op == 'q':
            self.state_stack.append(self.state.copy())
        elif op == 'Q':
            if self.state_stack:
                self.state = self.state_stack.pop()
        
        # 텍스트 블록
        elif op == 'BT':
            # 텍스트 행렬 초기화
            self.state.tm = [1, 0, 0, 1, 0, 0]
            self.state.tlm = [1, 0, 0, 1, 0, 0]
        elif op == 'ET':
            pass
        
        # 폰트 설정: /F1 12 Tf
        elif op == 'Tf':
            if len(operands) >= 2:
                font_name = self._get_value(operands[-2])
                font_size = self._get_value(operands[-1])
                if isinstance(font_name, str):
                    self.state.font_name = font_name
                if isinstance(font_size, (int, float)):
                    self.state.font_size = float(font_size)
        
        # 텍스트 위치
        elif op == 'Td':  # tx ty Td
            if len(operands) >= 2:
                tx = self._get_number(operands[-2])
                ty = self._get_number(operands[-1])
                self._move_text(tx, ty)
        
        elif op == 'TD':  # tx ty TD (= -ty TL tx ty Td)
            if len(operands) >= 2:
                tx = self._get_number(operands[-2])
                ty = self._get_number(operands[-1])
                self.state.leading = -ty
                self._move_text(tx, ty)
        
        elif op == 'Tm':  # a b c d e f Tm
            if len(operands) >= 6:
                self.state.tm = [self._get_number(operands[i]) for i in range(6)]
                self.state.tlm = self.state.tm.copy()
        
        elif op == 'T*':  # 다음 줄
            self._move_text(0, -self.state.leading)
        
        # 텍스트 속성
        elif op == 'Tc':  # charSpace Tc
            if operands:
                self.state.char_spacing = self._get_number(operands[-1])
        
        elif op == 'Tw':  # wordSpace Tw
            if operands:
                self.state.word_spacing = self._get_number(operands[-1])
        
        elif op == 'TL':  # leading TL
            if operands:
                self.state.leading = self._get_number(operands[-1])
        
        elif op == 'Tz':  # scale Tz
            if operands:
                self.state.horizontal_scale = self._get_number(operands[-1])
        
        elif op == 'Ts':  # rise Ts
            if operands:
                self.state.rise = self._get_number(operands[-1])
        
        # 텍스트 출력
        elif op == 'Tj':  # (string) Tj
            if operands:
                self._show_text(operands[-1])
        
        elif op == 'TJ':  # [(string) num (string) ...] TJ
            if operands:
                self._show_text_array(operands[-1])
        
        elif op == "'":  # (string) ' (= T* string Tj)
            self._move_text(0, -self.state.leading)
            if operands:
                self._show_text(operands[-1])
        
        elif op == '"':  # aw ac (string) " (= aw Tw ac Tc string ')
            if len(operands) >= 3:
                self.state.word_spacing = self._get_number(operands[-3])
                self.state.char_spacing = self._get_number(operands[-2])
                self._move_text(0, -self.state.leading)
                self._show_text(operands[-1])
    
    def _get_value(self, item) -> Any:
        """토큰 또는 값에서 실제 값 추출"""
        if isinstance(item, CSToken):
            return item.value
        return item
    
    def _get_number(self, item) -> float:
        """숫자 값 추출"""
        value = self._get_value(item)
        if isinstance(value, (int, float)):
            return float(value)
        return 0.0
    
    def _move_text(self, tx: float, ty: float):
        """텍스트 위치 이동"""
        # Td: tlm을 기준으로 이동하고 tm 업데이트
        # tm = [[1 0 0], [0 1 0], [tx ty 1]] × tlm
        new_x = self.state.tlm[4] + tx * self.state.tlm[0] + ty * self.state.tlm[2]
        new_y = self.state.tlm[5] + tx * self.state.tlm[1] + ty * self.state.tlm[3]
        
        self.state.tm[4] = new_x
        self.state.tm[5] = new_y
        self.state.tlm[4] = new_x
        self.state.tlm[5] = new_y
    
    def _show_text(self, item):
        """텍스트 출력 (Tj)"""
        if isinstance(item, CSToken):
            raw_bytes = item.value
            is_hex = item.type == CSTokenType.HEX_STRING
        elif isinstance(item, bytes):
            raw_bytes = item
            is_hex = False
        else:
            return
        
        # 텍스트 디코딩
        text = self._decode_text(raw_bytes, is_hex)
        
        if text:
            x, y = self.state.get_position()
            self.text_items.append(TextItem(
                text=text,
                x=x,
                y=y,
                font_name=self.state.font_name,
                font_size=self.state.font_size,
                char_spacing=self.state.char_spacing,
                word_spacing=self.state.word_spacing
            ))
    
    def _show_text_array(self, array):
        """TJ 배열 처리"""
        if not isinstance(array, list):
            return
        
        texts = []
        for item in array:
            if isinstance(item, CSToken):
                if item.type in (CSTokenType.STRING, CSTokenType.HEX_STRING):
                    text = self._decode_text(item.value, item.type == CSTokenType.HEX_STRING)
                    if text:
                        texts.append(text)
                elif item.type == CSTokenType.NUMBER:
                    # 큰 음수 값은 공백으로 처리 (커닝)
                    if item.value < -100:
                        texts.append(' ')
        
        if texts:
            full_text = ''.join(texts)
            x, y = self.state.get_position()
            self.text_items.append(TextItem(
                text=full_text,
                x=x,
                y=y,
                font_name=self.state.font_name,
                font_size=self.state.font_size
            ))
    
    def _decode_text(self, raw_bytes: bytes, is_hex: bool = False) -> str:
        """바이트를 텍스트로 디코딩"""
        if not raw_bytes:
            return ""
        
        # 폰트의 ToUnicode CMap이 있으면 사용
        font_info = self.font_map.get(self.state.font_name)
        if font_info and font_info.to_unicode:
            return self._decode_with_cmap(raw_bytes, font_info.to_unicode)
        
        # Hex 문자열이고 2바이트 단위면 UTF-16BE 시도
        if is_hex and len(raw_bytes) >= 2 and len(raw_bytes) % 2 == 0:
            try:
                # BOM 체크
                if raw_bytes[:2] == b'\xfe\xff':
                    return raw_bytes[2:].decode('utf-16-be')
                # 일반 UTF-16BE 시도
                return raw_bytes.decode('utf-16-be')
            except:
                pass
        
        # 기본: latin-1 (PDFDocEncoding과 유사)
        try:
            return raw_bytes.decode('latin-1')
        except:
            return raw_bytes.decode('utf-8', errors='replace')
    
    def _decode_with_cmap(self, raw_bytes: bytes, cmap: Dict[int, str]) -> str:
        """CMap을 사용해서 디코딩"""
        result = []
        
        # 2바이트 단위로 처리 (CID 폰트)
        if len(raw_bytes) % 2 == 0:
            for i in range(0, len(raw_bytes), 2):
                code = (raw_bytes[i] << 8) | raw_bytes[i + 1]
                if code in cmap:
                    result.append(cmap[code])
                else:
                    # CMap에 없으면 1바이트로 시도
                    if raw_bytes[i] in cmap:
                        result.append(cmap[raw_bytes[i]])
                    if raw_bytes[i + 1] in cmap:
                        result.append(cmap[raw_bytes[i + 1]])
        else:
            # 1바이트 단위
            for b in raw_bytes:
                if b in cmap:
                    result.append(cmap[b])
                else:
                    result.append(chr(b))
        
        return ''.join(result)


@dataclass
class FontInfo:
    """폰트 정보"""
    name: str
    subtype: str = ""
    base_font: str = ""
    encoding: str = ""
    to_unicode: Dict[int, str] = field(default_factory=dict)  # GID → 유니코드


def parse_tounicode_cmap(cmap_data: bytes) -> Dict[int, str]:
    """
    ToUnicode CMap 파싱
    
    CMap 형식 예시:
    beginbfchar
    <0048> <0048>
    <0065> <0065>
    endbfchar
    beginbfrange
    <0000> <00FF> <0000>
    endbfrange
    """
    result = {}
    
    try:
        text = cmap_data.decode('latin-1')
    except:
        return result
    
    # bfchar 파싱: <src> <dst>
    bfchar_pattern = r'beginbfchar\s*(.*?)\s*endbfchar'
    for match in re.finditer(bfchar_pattern, text, re.DOTALL):
        content = match.group(1)
        # <XXXX> <YYYY> 패턴
        pairs = re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', content)
        for src, dst in pairs:
            try:
                src_code = int(src, 16)
                # dst를 유니코드 문자열로 변환
                dst_bytes = bytes.fromhex(dst)
                if len(dst_bytes) == 2:
                    dst_char = dst_bytes.decode('utf-16-be')
                else:
                    dst_char = ''.join(chr(b) for b in dst_bytes)
                result[src_code] = dst_char
            except:
                pass
    
    # bfrange 파싱: <start> <end> <dst> 또는 <start> <end> [<dst1> <dst2> ...]
    bfrange_pattern = r'beginbfrange\s*(.*?)\s*endbfrange'
    for match in re.finditer(bfrange_pattern, text, re.DOTALL):
        content = match.group(1)
        
        # 패턴 1: <start> <end> <dst> (단일 시작점)
        ranges = re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', content)
        for start, end, dst in ranges:
            try:
                start_code = int(start, 16)
                end_code = int(end, 16)
                dst_code = int(dst, 16)
                
                for i, code in enumerate(range(start_code, end_code + 1)):
                    char_code = dst_code + i
                    if char_code < 0x10000:
                        result[code] = chr(char_code)
            except:
                pass
        
        # 패턴 2: <start> <end> [<dst1> <dst2> ...] (배열)
        array_ranges = re.findall(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*\[([^\]]+)\]', content)
        for start, end, dst_array in array_ranges:
            try:
                start_code = int(start, 16)
                end_code = int(end, 16)
                
                # 배열에서 각 매핑 추출
                dst_codes = re.findall(r'<([0-9A-Fa-f]+)>', dst_array)
                
                for i, code in enumerate(range(start_code, end_code + 1)):
                    if i < len(dst_codes):
                        char_code = int(dst_codes[i], 16)
                        if char_code < 0x10000:
                            result[code] = chr(char_code)
            except:
                pass
    
    return result


# 테스트
if __name__ == '__main__':
    # 간단한 Content Stream 테스트
    test_content = b"""
    BT
    /F1 12 Tf
    100 700 Td
    (Hello World!) Tj
    0 -20 Td
    (Second Line) Tj
    0 -20 Td
    [(K) -20 (erning) 30 ( Test)] TJ
    ET
    """
    
    parser = ContentStreamParser()
    items = parser.parse(test_content)
    
    print("추출된 텍스트:")
    print("-" * 40)
    for item in items:
        print(f"  [{item.x:.0f}, {item.y:.0f}] {item.font_name} {item.font_size}pt: '{item.text}'")
