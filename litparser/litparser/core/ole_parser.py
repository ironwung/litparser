"""
OLE2 (Compound File Binary Format) Parser

DOC, PPT, XLS, HWP(5.0 Binary) 파일의 기반이 되는 컨테이너 포맷 파서.
순수 Python 구현 - 외부 라이브러리 없음

구조:
- Header: 파일 메타데이터 (512바이트)
- FAT (File Allocation Table): 섹터 체인 관리
- Directory: 파일/스트림 목록 관리
- MiniFAT: 작은 스트림(4KB 미만) 관리용 FAT
"""

import struct
from dataclasses import dataclass
from typing import List, Optional, Union

# 상수 정의
HEADER_SIGNATURE = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
DIFAT_SECTORS_IN_HEADER = 109

# 특수 섹터 ID
FAT_SECT_FREE = 0xFFFFFFFF
FAT_SECT_END = 0xFFFFFFFE
FAT_SECT_FAT = 0xFFFFFFFD
FAT_SECT_DIFAT = 0xFFFFFFFC


@dataclass
class DirectoryEntry:
    """디렉토리 엔트리"""
    name: str
    entry_type: int  # 0: Empty, 1: Storage, 2: Stream, 5: Root
    color: int
    left_sibling: int
    right_sibling: int
    child: int
    start_sector: int
    size: int


class OLE2Reader:
    """OLE2 컨테이너 파일 리더"""
    
    def __init__(self, filepath_or_bytes: Union[str, bytes]):
        if isinstance(filepath_or_bytes, str):
            with open(filepath_or_bytes, 'rb') as f:
                self.data = f.read()
        else:
            self.data = filepath_or_bytes

        self._parse_header()
        self._build_fat()
        self._read_directory()
        self._build_minifat()

    def _parse_header(self):
        """헤더 파싱 (512 bytes)"""
        if len(self.data) < 512:
            raise ValueError("파일이 너무 작음 - OLE2 형식이 아님")
            
        header = self.data[:512]
        
        # 시그니처 확인
        if header[:8] != HEADER_SIGNATURE:
            raise ValueError("OLE2 시그니처가 아님")

        # 섹터 크기 (보통 512 bytes)
        self.sector_shift = struct.unpack('<H', header[30:32])[0]
        self.sector_size = 1 << self.sector_shift
        
        # Mini 섹터 크기 (보통 64 bytes)
        self.mini_sector_shift = struct.unpack('<H', header[32:34])[0]
        self.mini_sector_size = 1 << self.mini_sector_shift

        # 주요 정보 위치
        self.num_fat_sectors = struct.unpack('<I', header[44:48])[0]
        self.first_dir_sector = struct.unpack('<I', header[48:52])[0]
        self.mini_cutoff_size = struct.unpack('<I', header[56:60])[0]
        self.first_minifat_sector = struct.unpack('<I', header[60:64])[0]
        self.num_minifat_sectors = struct.unpack('<I', header[64:68])[0]
        self.first_difat_sector = struct.unpack('<I', header[68:72])[0]
        self.num_difat_sectors = struct.unpack('<I', header[72:76])[0]

        # 초기 DIFAT (Header 내 109개)
        self.difat = []
        for i in range(DIFAT_SECTORS_IN_HEADER):
            offset = 76 + i * 4
            sector = struct.unpack('<I', header[offset:offset+4])[0]
            if sector != FAT_SECT_FREE:
                self.difat.append(sector)

    def _get_sector_offset(self, sector: int) -> int:
        """섹터 번호를 파일 오프셋으로 변환"""
        return 512 + sector * self.sector_size

    def _build_fat(self):
        """FAT (File Allocation Table) 구축"""
        # DIFAT 확장 (Header 109개 초과 시)
        current_difat_sector = self.first_difat_sector
        while current_difat_sector != FAT_SECT_END and current_difat_sector != FAT_SECT_FREE:
            offset = self._get_sector_offset(current_difat_sector)
            sector_data = self.data[offset:offset + self.sector_size]
            
            # 마지막 4바이트는 다음 DIFAT 섹터 포인터
            ints_per_sector = self.sector_size // 4
            for i in range(ints_per_sector - 1):
                sector = struct.unpack('<I', sector_data[i*4:(i+1)*4])[0]
                if sector != FAT_SECT_FREE:
                    self.difat.append(sector)
            
            current_difat_sector = struct.unpack('<I', sector_data[-4:])[0]

        # FAT 로드
        self.fat = []
        for sector in self.difat:
            offset = self._get_sector_offset(sector)
            sector_data = self.data[offset:offset + self.sector_size]
            
            ints_per_sector = self.sector_size // 4
            for i in range(ints_per_sector):
                val = struct.unpack('<I', sector_data[i*4:(i+1)*4])[0]
                self.fat.append(val)

    def _get_chain(self, start_sector: int) -> List[int]:
        """FAT 체인을 따라가며 섹터 목록 반환"""
        chain = []
        current = start_sector
        visited = set()
        
        while current != FAT_SECT_END and current != FAT_SECT_FREE:
            if current >= len(self.fat) or current in visited:
                break  # 깨진 체인
            visited.add(current)
            chain.append(current)
            current = self.fat[current]
        
        return chain

    def _read_directory(self):
        """디렉토리 엔트리 파싱"""
        chain = self._get_chain(self.first_dir_sector)
        dir_data = bytearray()
        
        for sector in chain:
            offset = self._get_sector_offset(sector)
            dir_data.extend(self.data[offset:offset + self.sector_size])

        self.entries = []
        self.root_entry = None
        self._entry_map = {}  # 이름 -> 엔트리 매핑
        
        # 각 엔트리는 128 바이트
        num_entries = len(dir_data) // 128
        for i in range(num_entries):
            entry_bytes = dir_data[i*128:(i+1)*128]
            
            # 이름 길이 (바이트 단위, null 포함)
            name_len = struct.unpack('<H', entry_bytes[64:66])[0]
            if name_len > 0 and name_len <= 64:
                # UTF-16LE 디코딩 (마지막 null 문자 제외)
                name = entry_bytes[:name_len-2].decode('utf-16le', errors='ignore')
            else:
                name = ""

            entry_type = entry_bytes[66]  # 0:Empty, 1:Storage, 2:Stream, 5:Root
            
            if entry_type == 0:
                continue  # 빈 엔트리 스킵
                
            color = entry_bytes[67]
            left = struct.unpack('<I', entry_bytes[68:72])[0]
            right = struct.unpack('<I', entry_bytes[72:76])[0]
            child = struct.unpack('<I', entry_bytes[76:80])[0]
            start_sector = struct.unpack('<I', entry_bytes[116:120])[0]
            size = struct.unpack('<Q', entry_bytes[120:128])[0]

            entry = DirectoryEntry(
                name=name,
                entry_type=entry_type,
                color=color,
                left_sibling=left,
                right_sibling=right,
                child=child,
                start_sector=start_sector,
                size=size
            )
            self.entries.append(entry)
            self._entry_map[name] = entry

            if entry_type == 5:  # Root Entry
                self.root_entry = entry

    def _build_minifat(self):
        """MiniFAT 구축"""
        self.minifat = []
        self.mini_stream_data = bytearray()
        
        if not self.root_entry:
            return

        # MiniFAT 데이터 읽기
        if self.first_minifat_sector != FAT_SECT_END:
            chain = self._get_chain(self.first_minifat_sector)
            minifat_data = bytearray()
            
            for sector in chain:
                offset = self._get_sector_offset(sector)
                minifat_data.extend(self.data[offset:offset + self.sector_size])

            num_ints = len(minifat_data) // 4
            for i in range(num_ints):
                val = struct.unpack('<I', minifat_data[i*4:(i+1)*4])[0]
                self.minifat.append(val)

        # MiniStream 데이터 읽기 (Root Entry가 가리키는 스트림)
        if self.root_entry.start_sector != FAT_SECT_END:
            root_chain = self._get_chain(self.root_entry.start_sector)
            for sector in root_chain:
                offset = self._get_sector_offset(sector)
                self.mini_stream_data.extend(self.data[offset:offset + self.sector_size])

    def get_stream(self, stream_name: str) -> Optional[bytes]:
        """
        스트림 데이터 가져오기
        
        Args:
            stream_name: 스트림 이름 (예: "WordDocument", "Section0")
        
        Returns:
            bytes: 스트림 데이터 또는 None
        """
        entry = self._entry_map.get(stream_name)
        if entry and entry.entry_type == 2:
            return self._read_stream_data(entry)
        return None

    def _read_stream_data(self, entry: DirectoryEntry) -> bytes:
        """엔트리의 데이터 읽기"""
        data = bytearray()
        
        if entry.size == 0:
            return bytes()
        
        if entry.size < self.mini_cutoff_size and self.minifat:
            # MiniStream 사용
            current = entry.start_sector
            visited = set()
            
            while current != FAT_SECT_END and current != FAT_SECT_FREE:
                if current >= len(self.minifat) or current in visited:
                    break
                visited.add(current)
                
                offset = current * self.mini_sector_size
                if offset + self.mini_sector_size <= len(self.mini_stream_data):
                    chunk = self.mini_stream_data[offset:offset + self.mini_sector_size]
                    data.extend(chunk)
                
                current = self.minifat[current]
        else:
            # 일반 FAT 사용
            chain = self._get_chain(entry.start_sector)
            for sector in chain:
                offset = self._get_sector_offset(sector)
                data.extend(self.data[offset:offset + self.sector_size])
        
        # 크기에 맞게 자르기
        return bytes(data[:entry.size])

    def list_streams(self) -> List[str]:
        """모든 스트림 이름 나열"""
        return [e.name for e in self.entries if e.entry_type == 2]
    
    def list_storages(self) -> List[str]:
        """모든 스토리지 이름 나열"""
        return [e.name for e in self.entries if e.entry_type == 1]
    
    def list_all(self) -> List[str]:
        """모든 엔트리 이름 나열"""
        return [e.name for e in self.entries]


def is_ole2_file(data: bytes) -> bool:
    """OLE2 파일인지 확인"""
    return len(data) >= 8 and data[:8] == HEADER_SIGNATURE
