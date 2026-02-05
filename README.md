# LitParser

**Lit**eweight Document **Parser** - 순수 Python 문서 파서

**외부 라이브러리 없이** PDF, DOCX, PPTX, XLSX, HWPX 파싱

## 설치

```bash
pip install litparser
```

## 사용법

```python
from litparser import parse, to_markdown, to_json

# 자동 포맷 감지
result = parse('document.pdf')
result = parse('report.docx')
result = parse('data.xlsx')

# 결과 접근
print(result.text)      # 전체 텍스트
print(result.tables)    # 테이블 목록
print(result.pages)     # 페이지/시트별 데이터

# 변환
md = to_markdown(result)
json_str = to_json(result)
```

## CLI

```bash
litparser document.pdf
litparser document.pdf --markdown
litparser document.pdf --json
litparser data.xlsx --tables
litparser document.pdf --info
```

## 지원 포맷

| 포맷 | 확장자 | 텍스트 | 테이블 | 이미지 |
|------|--------|--------|--------|--------|
| PDF | .pdf | ✅ | ✅ | ✅ |
| Word | .docx | ✅ | ✅ | ✅ |
| PowerPoint | .pptx | ✅ | ✅ | ✅ |
| Excel | .xlsx | ✅ | ✅ | - |
| 한글 | .hwpx | ✅ | ✅ | ✅ |
| 텍스트 | .txt, .md | ✅ | - | - |

## 라이선스

MIT License
