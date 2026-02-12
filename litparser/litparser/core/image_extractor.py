"""
PDF Image Extractor - Stage 5

PDF에서 이미지 추출 기능

지원하는 이미지 타입:
- DCTDecode (JPEG)
- FlateDecode (PNG raw data)
- JPXDecode (JPEG2000)
- CCITTFaxDecode (Fax/TIFF)
"""

import zlib
from dataclasses import dataclass
from typing import List, Optional, Tuple, Any, Dict
from .stream_decoder import StreamDecoder


@dataclass
class PDFImage:
    """추출된 이미지 정보"""
    obj_num: int
    width: int
    height: int
    color_space: str
    bits_per_component: int
    filter: str  # 원본 필터 (DCTDecode, FlateDecode 등)
    data: bytes  # 이미지 데이터
    
    @property
    def format(self) -> str:
        """이미지 포맷 추정"""
        if self.filter == 'DCTDecode':
            return 'jpeg'
        elif self.filter == 'JPXDecode':
            return 'jp2'
        elif self.filter == 'CCITTFaxDecode':
            return 'tiff'
        elif self.filter == 'FlateDecode':
            return 'raw'  # PNG로 변환 필요
        return 'unknown'


def extract_images(doc, page_num: int = None) -> List[PDFImage]:
    """
    PDF에서 이미지 추출
    
    Args:
        doc: PDFDocument 객체
        page_num: 특정 페이지만 추출 (None이면 전체)
    
    Returns:
        List[PDFImage]: 추출된 이미지 목록
    """
    from . import PDFRef
    
    images = []
    
    # 이미지 XObject 찾기
    for (obj_num, gen_num), obj in doc.objects.items():
        if not isinstance(obj, dict):
            continue
        
        subtype = obj.get('Subtype')
        if subtype != 'Image':
            continue
        
        # 이미지 정보 추출
        width = obj.get('Width', 0)
        height = obj.get('Height', 0)
        bpc = obj.get('BitsPerComponent', 8)
        
        # 색상 공간 및 채널 수 결정
        cs = obj.get('ColorSpace', 'DeviceGray')
        channels = 1  # 기본값
        
        if isinstance(cs, list):
            # [/ICCBased ref] 또는 [/Indexed base ...] 형식
            cs_type = cs[0] if cs else 'DeviceGray'
            
            if cs_type == 'ICCBased' and len(cs) > 1:
                # ICC Profile에서 채널 수 가져오기
                icc_ref = cs[1]
                if isinstance(icc_ref, PDFRef):
                    icc_obj = doc.objects.get((icc_ref.obj_num, icc_ref.gen_num))
                    if icc_obj and isinstance(icc_obj, dict):
                        channels = icc_obj.get('N', 3)
                color_space = f'ICCBased({channels}ch)'
            elif cs_type == 'Indexed':
                color_space = 'Indexed'
                channels = 1  # 인덱스 색상
            else:
                color_space = str(cs_type)
        elif isinstance(cs, PDFRef):
            cs_obj = doc.objects.get((cs.obj_num, cs.gen_num))
            if isinstance(cs_obj, list) and cs_obj:
                color_space = str(cs_obj[0])
            else:
                color_space = str(cs_obj) if cs_obj else 'DeviceGray'
        else:
            color_space = str(cs) if cs else 'DeviceGray'
        
        # 색상 공간에서 채널 수 추론
        if channels == 1:  # 아직 결정 안됨
            if 'RGB' in color_space or color_space == 'DeviceRGB':
                channels = 3
            elif 'CMYK' in color_space or color_space == 'DeviceCMYK':
                channels = 4
            elif 'Gray' in color_space or color_space == 'DeviceGray':
                channels = 1
        
        # 필터
        filter_name = obj.get('Filter', '')
        if isinstance(filter_name, list):
            filter_name = filter_name[0] if filter_name else ''
        
        # 스트림 데이터
        if '_stream_data' not in obj:
            continue
        
        raw_data = obj['_stream_data']
        
        # DCTDecode (JPEG)는 그대로 사용
        if filter_name == 'DCTDecode':
            image_data = raw_data
        elif filter_name == 'JPXDecode':
            image_data = raw_data
        elif filter_name == 'FlateDecode':
            # zlib 압축 해제
            try:
                image_data = zlib.decompress(raw_data)
            except:
                image_data = raw_data
        else:
            # 다른 필터는 StreamDecoder 사용
            try:
                if filter_name:
                    image_data = StreamDecoder.decode(raw_data, filter_name)
                else:
                    image_data = raw_data
            except:
                image_data = raw_data
        
        # 채널 수를 color_space에 포함
        if 'ICCBased' in color_space:
            pass  # 이미 포함됨
        elif channels == 3:
            color_space = 'RGB'
        elif channels == 4:
            color_space = 'CMYK'
        elif channels == 1:
            color_space = 'Grayscale'
        
        images.append(PDFImage(
            obj_num=obj_num,
            width=width,
            height=height,
            color_space=color_space,
            bits_per_component=bpc,
            filter=filter_name,
            data=image_data
        ))
    
    return images


def save_image(image: PDFImage, filepath: str) -> bool:
    """
    이미지를 파일로 저장
    
    Args:
        image: PDFImage 객체
        filepath: 저장 경로 (확장자 자동 결정)
    
    Returns:
        bool: 성공 여부
    """
    try:
        # JPEG는 그대로 저장
        if image.filter == 'DCTDecode':
            if not filepath.lower().endswith(('.jpg', '.jpeg')):
                filepath += '.jpg'
            with open(filepath, 'wb') as f:
                f.write(image.data)
            return True
        
        # JPEG2000도 그대로 저장
        if image.filter == 'JPXDecode':
            if not filepath.lower().endswith('.jp2'):
                filepath += '.jp2'
            with open(filepath, 'wb') as f:
                f.write(image.data)
            return True
        
        # Raw 데이터 (FlateDecode 등)는 PNG로 저장
        if not filepath.lower().endswith('.png'):
            filepath += '.png'
        
        png_data = raw_to_png(image)
        with open(filepath, 'wb') as f:
            f.write(png_data)
        return True
        
    except Exception as e:
        print(f"이미지 저장 실패: {e}")
        return False


def _save_as_ppm(image: PDFImage, filepath: str) -> bool:
    """Raw 이미지 데이터를 PPM/PGM 포맷으로 저장"""
    try:
        width = image.width
        height = image.height
        data = image.data
        
        # 색상 공간에 따라 PPM 또는 PGM
        if 'RGB' in image.color_space or image.color_space == 'DeviceRGB':
            # PPM (컬러)
            header = f"P6\n{width} {height}\n255\n".encode('ascii')
            ext = '.ppm'
        else:
            # PGM (그레이스케일)
            header = f"P5\n{width} {height}\n255\n".encode('ascii')
            ext = '.pgm'
        
        # 확장자 변경
        if not filepath.endswith(ext):
            filepath = filepath.rsplit('.', 1)[0] + ext
        
        with open(filepath, 'wb') as f:
            f.write(header)
            f.write(data)
        
        return True
    except Exception as e:
        print(f"PPM 저장 실패: {e}")
        return False


def raw_to_png(image: PDFImage) -> bytes:
    """
    Raw 이미지 데이터를 PNG로 변환
    
    Note: 이 함수는 외부 라이브러리 없이 PNG를 생성합니다.
    복잡한 이미지의 경우 Pillow 사용을 권장합니다.
    """
    import struct
    
    width = image.width
    height = image.height
    data = image.data
    
    # 색상 타입 및 채널 수 결정
    color_space = image.color_space
    
    # 데이터 크기로 채널 수 추론
    expected_pixels = width * height
    actual_size = len(data)
    
    if actual_size == expected_pixels * 3:
        color_type = 2  # RGB
        channels = 3
    elif actual_size == expected_pixels * 4:
        color_type = 6  # RGBA
        channels = 4
    elif actual_size == expected_pixels:
        color_type = 0  # Grayscale
        channels = 1
    elif 'RGB' in color_space or '3ch' in color_space:
        color_type = 2  # RGB
        channels = 3
    elif 'CMYK' in color_space or '4ch' in color_space:
        # CMYK -> RGB 변환 필요
        data = _cmyk_to_rgb(data, width, height)
        color_type = 2
        channels = 3
    else:
        color_type = 0  # Grayscale
        channels = 1
    
    # PNG 헤더
    png_signature = b'\x89PNG\r\n\x1a\n'
    
    # IHDR chunk
    ihdr_data = struct.pack('>IIBBBBB', width, height, 8, color_type, 0, 0, 0)
    ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data) & 0xffffffff
    ihdr_chunk = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)
    
    # IDAT chunk (이미지 데이터)
    # 각 행 앞에 필터 바이트 추가 (0 = None)
    row_size = width * channels
    filtered_data = b''
    for y in range(height):
        filtered_data += b'\x00'  # 필터 타입: None
        row_start = y * row_size
        row_end = row_start + row_size
        if row_end <= len(data):
            filtered_data += data[row_start:row_end]
        else:
            # 데이터가 부족하면 0으로 채움
            available = data[row_start:] if row_start < len(data) else b''
            filtered_data += available + b'\x00' * (row_size - len(available))
    
    compressed_data = zlib.compress(filtered_data, 9)
    idat_crc = zlib.crc32(b'IDAT' + compressed_data) & 0xffffffff
    idat_chunk = struct.pack('>I', len(compressed_data)) + b'IDAT' + compressed_data + struct.pack('>I', idat_crc)
    
    # IEND chunk
    iend_crc = zlib.crc32(b'IEND') & 0xffffffff
    iend_chunk = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)
    
    return png_signature + ihdr_chunk + idat_chunk + iend_chunk


def _cmyk_to_rgb(data: bytes, width: int, height: int) -> bytes:
    """CMYK 데이터를 RGB로 변환"""
    rgb_data = bytearray()
    
    for i in range(0, len(data), 4):
        if i + 4 > len(data):
            break
        c, m, y, k = data[i], data[i+1], data[i+2], data[i+3]
        
        # CMYK to RGB 변환 공식
        r = int(255 * (1 - c/255) * (1 - k/255))
        g = int(255 * (1 - m/255) * (1 - k/255))
        b = int(255 * (1 - y/255) * (1 - k/255))
        
        rgb_data.extend([r, g, b])
    
    return bytes(rgb_data)


# 테스트
if __name__ == '__main__':
    print("이미지 추출 모듈 로드됨")
