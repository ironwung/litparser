# LitParser

**Lit**eweight Document **Parser** - 순수 Python으로 구현된 문서 파서.

**외부 라이브러리 없이** PDF, DOCX, PPTX, HWPX 파일을 파싱합니다.

## 특징

- ✅ **Zero Dependencies** - 표준 라이브러리만 사용
- ✅ **다양한 포맷** - PDF, DOCX, PPTX, HWPX, TXT, MD
- ✅ **텍스트 추출** - 위치 정보 포함
- ✅ **테이블 감지** - 마크다운 변환
- ✅ **이미지 추출** - PNG, JPEG, JP2
- ✅ **출력 포맷** - Markdown, JSON

## 설치

```bash
pip install litparser
```

소스에서:
```bash
pip install -e .
```

## CLI

```bash
# 텍스트 추출
litparser document.pdf
litparser report.docx

# 마크다운 변환
litparser document.pdf --markdown
litparser document.pdf --md -o result.md

# JSON 변환
litparser document.pdf --json
litparser document.pdf --json --include-images

# 테이블
litparser document.pdf --tables

# 분석
litparser document.pdf --analyze
```

## Python API

```python
from litparser import parse_pdf, extract_text, extract_tables

# PDF 파싱
doc = parse_pdf('document.pdf')

# 텍스트
text = extract_text(doc, page_num=0)

# 테이블
tables = extract_tables(doc, page_num=0)
for t in tables:
    print(t.to_markdown())

# 마크다운/JSON 변환
from litparser.output_formatter import pdf_to_output, to_markdown, to_json

output = pdf_to_output(doc)
md = to_markdown(output)
json_str = to_json(output)
```

## 지원 포맷

| 포맷 | 확장자 | 텍스트 | 테이블 | 이미지 |
|------|--------|--------|--------|--------|
| PDF | .pdf | ✅ | ✅ | ✅ |
| Word | .docx | ✅ | ✅ | ✅ |
| PowerPoint | .pptx | ✅ | ✅ | ✅ |
| 한글 | .hwpx | ✅ | ✅ | ✅ |
| 텍스트 | .txt, .md | ✅ | - | - |

## 라이선스

MIT License
