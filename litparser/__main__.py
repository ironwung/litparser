"""
LitParser CLI

사용법:
    litparser document.pdf
    litparser report.docx --markdown
    litparser slides.pptx --json
"""

import sys
import os
import argparse
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        description='LitParser - Lightweight Document Parser',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
지원 포맷:
  .pdf, .docx, .pptx, .xlsx, .hwpx, .txt, .md

출력 포맷:
  --markdown    마크다운으로 변환
  --json        JSON으로 변환

예시:
  litparser document.pdf
  litparser document.pdf --markdown
  litparser document.pdf --json --include-images
  litparser report.docx -o result.md
'''
    )
    
    parser.add_argument('file', help='문서 파일 경로')
    parser.add_argument('--markdown', '--md', action='store_true', help='마크다운으로 변환')
    parser.add_argument('--json', '-j', action='store_true', help='JSON으로 변환')
    parser.add_argument('--include-images', action='store_true', help='이미지 포함 (base64)')
    parser.add_argument('--tables', action='store_true', help='테이블만 출력')
    parser.add_argument('--info', '-i', action='store_true', help='문서 정보')
    parser.add_argument('--page', '-p', type=int, help='특정 페이지 (0부터)')
    parser.add_argument('--output', '-o', help='출력 파일')
    
    args = parser.parse_args()
    
    filepath = Path(args.file)
    if not filepath.exists():
        print(f"오류: 파일을 찾을 수 없습니다: {filepath}", file=sys.stderr)
        sys.exit(1)
    
    ext = filepath.suffix.lower()
    supported = ['.pdf', '.docx', '.pptx', '.xlsx', '.hwpx', '.hwp', '.doc', '.ppt', '.xls', '.txt', '.md', '.markdown']
    if ext not in supported:
        print(f"오류: 지원하지 않는 포맷: {ext}", file=sys.stderr)
        sys.exit(1)
    
    try:
        from . import parse, to_markdown, to_json
        
        # 파싱
        result = parse(str(filepath), include_images=args.include_images)
        
        # 문서 정보
        if args.info:
            print(f"파일: {result.filename}")
            print(f"포맷: {result.format.upper()}")
            print(f"페이지: {result.page_count}")
            if result.title:
                print(f"제목: {result.title}")
            if result.author:
                print(f"작성자: {result.author}")
            print(f"테이블: {len(result.tables)}개")
            print(f"이미지: {len(result.images)}개")
            return
        
        # 테이블만
        if args.tables:
            for i, table in enumerate(result.tables, 1):
                print(f"\n테이블 {i} ({table['rows']}x{table['cols']}):")
                print(table['markdown'])
            return
        
        # 특정 페이지
        if args.page is not None and result.pages:
            if 0 <= args.page < len(result.pages):
                page = result.pages[args.page]
                print(f"--- 페이지 {page['page']} ---")
                print(page['text'])
                for t in page.get('tables', []):
                    print(f"\n{t['markdown']}")
            else:
                print(f"오류: 페이지 범위 초과 (0-{len(result.pages)-1})")
            return
        
        # 출력 포맷
        if args.markdown:
            output = to_markdown(result, include_images=args.include_images)
        elif args.json:
            output = to_json(result, include_images=args.include_images)
        else:
            # 기본: 텍스트
            output = result.text
        
        # 출력
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(output)
            print(f"저장됨: {args.output}", file=sys.stderr)
        else:
            print(output)
            
    except Exception as e:
        print(f"오류: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
