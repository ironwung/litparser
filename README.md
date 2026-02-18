# LitParser

Lightweight Document **Parser** - 순수 Python 문서 파서

**외부 라이브러리 없이** 다양한 문서 포맷 파싱

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
result = parse('문서.hwp')

# 결과 접근
print(result.text)
print(result.tables)

# 변환
md = to_markdown(result)
json_str = to_json(result)
```

## CLI

```bash
litparser document.pdf
litparser document.pdf --markdown
litparser document.pdf --json
litparser 문서.hwp --info
```

## 지원 포맷

| 포맷 | Modern | Legacy |
|------|--------|--------|
| Word | .docx ✅ | .doc ✅ |
| PowerPoint | .pptx ✅ | .ppt ✅ |
| Excel | .xlsx ✅ | .xls ✅ |
| 한글 | .hwpx ✅ | .hwp ✅ |
| PDF | .pdf ✅ | - |
| 텍스트 | .txt, .md ✅ | - |

## 라이선스

LitParser는 **듀얼 라이선스**로 제공됩니다.

### AGPL-3.0 (오픈소스)
- 오픈소스 프로젝트에서 자유롭게 사용 가능
- 네트워크 서비스(SaaS 등)로 제공 시 **전체 소스코드를 AGPL-3.0으로 공개** 필요
- 전문: https://www.gnu.org/licenses/agpl-3.0.html

### 상용 라이선스
- 소스코드 공개 없이 상업적 사용 가능
- 문의: ironwung@gmail.com
