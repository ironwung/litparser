"""
PDF Parser - Stage 2: 스트림 디코딩

지원하는 필터:
1. FlateDecode (zlib)
2. ASCII85Decode
3. ASCIIHexDecode
4. LZWDecode
"""

import zlib
from typing import List, Union


class StreamDecoder:
    """PDF 스트림 디코더"""
    
    @staticmethod
    def decode(data: bytes, filters: Union[str, List[str]], params: dict = None) -> bytes:
        """
        필터 체인을 적용해서 스트림 디코딩
        
        Args:
            data: 원본 스트림 데이터
            filters: 필터 이름 또는 필터 리스트
            params: 디코딩 파라미터 (DecodeParms)
        
        Returns:
            디코딩된 데이터
        """
        if isinstance(filters, str):
            filters = [filters]
        
        if params is None:
            params = {}
        
        result = data
        
        for filter_name in filters:
            if filter_name == 'FlateDecode':
                result = StreamDecoder.decode_flate(result, params)
            elif filter_name == 'ASCII85Decode':
                result = StreamDecoder.decode_ascii85(result)
            elif filter_name == 'ASCIIHexDecode':
                result = StreamDecoder.decode_asciihex(result)
            elif filter_name == 'LZWDecode':
                result = StreamDecoder.decode_lzw(result, params)
            elif filter_name == 'RunLengthDecode':
                result = StreamDecoder.decode_runlength(result)
            elif filter_name == 'DCTDecode':
                # JPEG - 그대로 반환 (이미지용)
                pass
            elif filter_name == 'JPXDecode':
                # JPEG2000 - 그대로 반환
                pass
            elif filter_name == 'CCITTFaxDecode':
                # 팩스 인코딩 - 나중에 구현
                raise NotImplementedError(f"Filter not implemented: {filter_name}")
            else:
                raise ValueError(f"Unknown filter: {filter_name}")
        
        return result
    
    @staticmethod
    def decode_flate(data: bytes, params: dict = None) -> bytes:
        """FlateDecode (zlib) 압축 해제"""
        try:
            decompressed = zlib.decompress(data)
        except zlib.error:
            # 일부 PDF는 헤더 없이 raw deflate 사용
            decompressed = zlib.decompress(data, -15)
        
        # Predictor 처리 (PNG 필터 등)
        if params:
            predictor = params.get('Predictor', 1)
            if predictor > 1:
                decompressed = StreamDecoder._apply_predictor(
                    decompressed,
                    predictor,
                    params.get('Columns', 1),
                    params.get('Colors', 1),
                    params.get('BitsPerComponent', 8)
                )
        
        return decompressed
    
    @staticmethod
    def decode_ascii85(data: bytes) -> bytes:
        """ASCII85 (Base85) 디코딩"""
        # 공백 제거
        data = bytes(b for b in data if b not in b' \t\n\r')
        
        # ~> 종료 마커 제거
        if data.endswith(b'~>'):
            data = data[:-2]
        
        result = bytearray()
        i = 0
        
        while i < len(data):
            if data[i:i+1] == b'z':
                # 'z'는 4개의 0바이트를 의미
                result.extend(b'\x00\x00\x00\x00')
                i += 1
                continue
            
            # 5글자 그룹 읽기
            chunk = data[i:i+5]
            chunk_len = len(chunk)
            
            if chunk_len == 0:
                break
            
            # 패딩
            if chunk_len < 5:
                chunk = chunk + b'u' * (5 - chunk_len)
            
            # 디코딩
            value = 0
            for c in chunk:
                if c < 33 or c > 117:
                    raise ValueError(f"Invalid ASCII85 character: {chr(c)}")
                value = value * 85 + (c - 33)
            
            # 4바이트로 변환
            decoded = value.to_bytes(4, 'big')
            
            # 마지막 그룹은 패딩만큼 잘라냄
            if chunk_len < 5:
                decoded = decoded[:chunk_len - 1]
            
            result.extend(decoded)
            i += min(5, chunk_len)
        
        return bytes(result)
    
    @staticmethod
    def decode_asciihex(data: bytes) -> bytes:
        """ASCIIHex 디코딩"""
        # 공백 제거
        data = bytes(b for b in data if b not in b' \t\n\r')
        
        # > 종료 마커 제거
        if data.endswith(b'>'):
            data = data[:-1]
        
        # 홀수 길이면 0 추가
        if len(data) % 2 == 1:
            data = data + b'0'
        
        return bytes.fromhex(data.decode('ascii'))
    
    @staticmethod
    def decode_lzw(data: bytes, params: dict = None) -> bytes:
        """LZW 압축 해제"""
        # LZW 테이블 초기화
        CLEAR_CODE = 256
        EOD_CODE = 257
        
        table = {i: bytes([i]) for i in range(256)}
        table[CLEAR_CODE] = None
        table[EOD_CODE] = None
        
        next_code = 258
        code_bits = 9
        
        result = bytearray()
        bit_buffer = 0
        bits_in_buffer = 0
        
        pos = 0
        prev_seq = None
        
        while pos < len(data):
            # 비트 버퍼 채우기
            while bits_in_buffer < code_bits and pos < len(data):
                bit_buffer = (bit_buffer << 8) | data[pos]
                bits_in_buffer += 8
                pos += 1
            
            if bits_in_buffer < code_bits:
                break
            
            # 코드 추출
            code = (bit_buffer >> (bits_in_buffer - code_bits)) & ((1 << code_bits) - 1)
            bits_in_buffer -= code_bits
            
            if code == CLEAR_CODE:
                # 테이블 리셋
                table = {i: bytes([i]) for i in range(256)}
                table[CLEAR_CODE] = None
                table[EOD_CODE] = None
                next_code = 258
                code_bits = 9
                prev_seq = None
                continue
            
            if code == EOD_CODE:
                break
            
            # 시퀀스 가져오기
            if code in table:
                seq = table[code]
            elif code == next_code and prev_seq:
                seq = prev_seq + prev_seq[0:1]
            else:
                raise ValueError(f"Invalid LZW code: {code}")
            
            result.extend(seq)
            
            # 테이블에 추가
            if prev_seq and next_code < 4096:
                table[next_code] = prev_seq + seq[0:1]
                next_code += 1
                
                # 코드 비트 수 증가
                if next_code >= (1 << code_bits) and code_bits < 12:
                    code_bits += 1
            
            prev_seq = seq
        
        # Predictor 처리
        if params:
            predictor = params.get('Predictor', 1)
            if predictor > 1:
                result = StreamDecoder._apply_predictor(
                    bytes(result),
                    predictor,
                    params.get('Columns', 1),
                    params.get('Colors', 1),
                    params.get('BitsPerComponent', 8)
                )
        
        return bytes(result)
    
    @staticmethod
    def decode_runlength(data: bytes) -> bytes:
        """RunLength 디코딩"""
        result = bytearray()
        i = 0
        
        while i < len(data):
            length = data[i]
            i += 1
            
            if length == 128:
                # EOD
                break
            elif length < 128:
                # 다음 length+1 바이트를 그대로 복사
                count = length + 1
                result.extend(data[i:i + count])
                i += count
            else:
                # 다음 1바이트를 257-length번 반복
                count = 257 - length
                result.extend(data[i:i + 1] * count)
                i += 1
        
        return bytes(result)
    
    @staticmethod
    def _apply_predictor(data: bytes, predictor: int, columns: int, 
                         colors: int = 1, bits: int = 8) -> bytes:
        """Predictor 역변환 (PNG 필터 등)"""
        if predictor == 1:
            return data  # 변환 없음
        
        if predictor == 2:
            # TIFF Predictor 2
            row_size = columns * colors * bits // 8
            result = bytearray()
            
            for row_start in range(0, len(data), row_size):
                row = bytearray(data[row_start:row_start + row_size])
                for i in range(colors * bits // 8, len(row)):
                    row[i] = (row[i] + row[i - colors * bits // 8]) & 0xFF
                result.extend(row)
            
            return bytes(result)
        
        if predictor >= 10:
            # PNG 필터
            bytes_per_pixel = colors * bits // 8
            row_size = columns * bytes_per_pixel
            result = bytearray()
            prev_row = bytes(row_size)
            
            i = 0
            while i < len(data):
                filter_type = data[i]
                i += 1
                row = bytearray(data[i:i + row_size])
                i += row_size
                
                if filter_type == 0:
                    # None
                    pass
                elif filter_type == 1:
                    # Sub
                    for j in range(bytes_per_pixel, len(row)):
                        row[j] = (row[j] + row[j - bytes_per_pixel]) & 0xFF
                elif filter_type == 2:
                    # Up
                    for j in range(len(row)):
                        row[j] = (row[j] + prev_row[j]) & 0xFF
                elif filter_type == 3:
                    # Average
                    for j in range(len(row)):
                        left = row[j - bytes_per_pixel] if j >= bytes_per_pixel else 0
                        up = prev_row[j]
                        row[j] = (row[j] + (left + up) // 2) & 0xFF
                elif filter_type == 4:
                    # Paeth
                    for j in range(len(row)):
                        left = row[j - bytes_per_pixel] if j >= bytes_per_pixel else 0
                        up = prev_row[j]
                        up_left = prev_row[j - bytes_per_pixel] if j >= bytes_per_pixel else 0
                        row[j] = (row[j] + StreamDecoder._paeth(left, up, up_left)) & 0xFF
                
                result.extend(row)
                prev_row = row
            
            return bytes(result)
        
        return data
    
    @staticmethod
    def _paeth(a: int, b: int, c: int) -> int:
        """Paeth predictor"""
        p = a + b - c
        pa = abs(p - a)
        pb = abs(p - b)
        pc = abs(p - c)
        
        if pa <= pb and pa <= pc:
            return a
        elif pb <= pc:
            return b
        else:
            return c


# 테스트
if __name__ == '__main__':
    print("Stream Decoder 테스트")
    print("=" * 50)
    
    # FlateDecode 테스트
    original = b"Hello, PDF World! This is a test of FlateDecode compression."
    compressed = zlib.compress(original)
    decoded = StreamDecoder.decode(compressed, 'FlateDecode')
    print(f"\n[FlateDecode]")
    print(f"  원본: {original}")
    print(f"  압축: {len(compressed)} bytes")
    print(f"  해제: {decoded}")
    print(f"  일치: {original == decoded}")
    
    # ASCII85 테스트
    ascii85_data = b"9jqo^BlbD-BleB1DJ+*+F(f,q"
    decoded = StreamDecoder.decode(ascii85_data, 'ASCII85Decode')
    print(f"\n[ASCII85Decode]")
    print(f"  입력: {ascii85_data}")
    print(f"  출력: {decoded}")
    
    # ASCIIHex 테스트
    hex_data = b"48 65 6C 6C 6F>"
    decoded = StreamDecoder.decode(hex_data, 'ASCIIHexDecode')
    print(f"\n[ASCIIHexDecode]")
    print(f"  입력: {hex_data}")
    print(f"  출력: {decoded}")
    
    # 필터 체인 테스트
    original = b"Testing filter chain: ASCII85 + FlateDecode"
    step1 = zlib.compress(original)
    
    # ASCII85 인코딩 (간단한 버전)
    def encode_ascii85(data):
        result = bytearray()
        for i in range(0, len(data), 4):
            chunk = data[i:i+4]
            if len(chunk) < 4:
                chunk = chunk + b'\x00' * (4 - len(chunk))
            value = int.from_bytes(chunk, 'big')
            encoded = []
            for _ in range(5):
                encoded.append(value % 85 + 33)
                value //= 85
            result.extend(reversed(encoded))
        result.extend(b'~>')
        return bytes(result)
    
    step2 = encode_ascii85(step1)
    
    # 디코딩
    decoded = StreamDecoder.decode(step2, ['ASCII85Decode', 'FlateDecode'])
    print(f"\n[Filter Chain: ASCII85 + Flate]")
    print(f"  원본: {original}")
    print(f"  해제: {decoded}")
    print(f"  일치: {original == decoded}")
