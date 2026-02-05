"""
Document Parser CLI

모든 지원 포맷을 자동으로 처리하는 통합 CLI

사용법:
    litparser document.pdf
    litparser document.docx --tables
    litparser slides.pptx --outline
"""

import sys
import os
import argparse
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        description='LitParser - PDF, DOCX, PPTX, HWPX, TXT, MD 지원',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
지원 포맷:
  .pdf          PDF 문서
  .docx         Microsoft Word
  .pptx         Microsoft PowerPoint  
  .hwpx         한글 (개방형)
  .txt, .md     텍스트/마크다운

출력 포맷:
  --markdown    마크다운으로 변환
  --json        JSON으로 변환 (구조화 데이터)

예시:
  litparser document.pdf
  litparser document.pdf --markdown
  litparser document.pdf --json --include-images
  litparser report.docx --tables
'''
    )
    
    parser.add_argument('file', help='문서 파일 경로')
    parser.add_argument('--text', '-t', action='store_true', help='텍스트만 추출')
    parser.add_argument('--all-text', '-a', action='store_true', help='모든 페이지 처리')
    parser.add_argument('--tables', action='store_true', help='테이블 추출')
    parser.add_argument('--images', action='store_true', help='이미지 정보')
    parser.add_argument('--save', '-s', action='store_true', help='이미지를 파일로 저장')
    parser.add_argument('--output-dir', default='.', help='이미지 저장 디렉토리')
    parser.add_argument('--outline', '-o', action='store_true', help='문서 구조/개요')
    parser.add_argument('--info', '-i', action='store_true', help='문서 정보')
    parser.add_argument('--analyze', action='store_true', help='상세 분석 (PDF)')
    parser.add_argument('--markdown', '--md', action='store_true', help='마크다운으로 변환')
    parser.add_argument('--json', '-j', action='store_true', help='JSON으로 변환')
    parser.add_argument('--include-images', action='store_true', help='출력에 이미지 포함 (base64)')
    parser.add_argument('--page', '-p', type=int, help='특정 페이지 (0부터 시작)')
    parser.add_argument('--output', help='출력 파일')
    
    args = parser.parse_args()
    
    filepath = Path(args.file)
    if not filepath.exists():
        print(f"오류: 파일을 찾을 수 없습니다: {filepath}", file=sys.stderr)
        sys.exit(1)
    
    ext = filepath.suffix.lower()
    
    try:
        if ext == '.pdf':
            process_pdf(filepath, args)
        elif ext == '.docx':
            process_docx(filepath, args)
        elif ext == '.pptx':
            process_pptx(filepath, args)
        elif ext == '.hwpx':
            process_hwpx(filepath, args)
        elif ext in ['.txt', '.md', '.markdown']:
            process_text(filepath, args)
        elif ext in ['.doc', '.ppt', '.hwp']:
            print(f"오류: {ext} 포맷은 지원하지 않습니다.", file=sys.stderr)
            print("바이너리 포맷으로 별도 라이브러리가 필요합니다.", file=sys.stderr)
            sys.exit(1)
        else:
            print(f"오류: 알 수 없는 파일 형식: {ext}", file=sys.stderr)
            sys.exit(1)
            
    except Exception as e:
        print(f"오류: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


def process_pdf(filepath, args):
    """PDF 처리"""
    from . import (
        parse_pdf, extract_text, extract_all_text,
        get_page_count, get_pages, extract_tables, extract_images,
        analyze_layout, is_tagged_pdf, save_image
    )
    
    doc = parse_pdf(str(filepath))
    page_count = get_page_count(doc)
    
    # Markdown/JSON 출력
    if args.markdown or args.json:
        from .output_formatter import pdf_to_output, to_markdown, to_json
        
        output = pdf_to_output(doc, include_images=args.include_images)
        output.filename = str(filepath)
        
        if args.markdown:
            result = to_markdown(output, include_images=args.include_images)
        else:
            result = to_json(output, include_images=args.include_images)
        
        _write_output(args, result)
        return
    
    print(f"PDF 분석: {filepath}")
    print("=" * 60)
    print(f"버전: PDF {doc.version}")
    print(f"객체 수: {len(doc.objects)}")
    print(f"페이지 수: {page_count}")
    
    if args.info:
        print(f"Tagged PDF: {is_tagged_pdf(doc)}")
        return
    
    # 상세 분석 모드
    if args.analyze:
        _run_integrated_analysis(doc, args, page_count)
        return
    
    if args.outline:
        from . import get_document_outline
        print("\n문서 개요:")
        try:
            outline = get_document_outline(doc)
            for level, text in outline:
                print("  " * (level - 1) + f"H{level}: {text}")
            if not outline:
                print("  (개요 없음)")
        except:
            print("  (개요 없음)")
        return
    
    if args.tables:
        pages = [args.page] if args.page is not None else range(page_count)
        for p in pages:
            tables = extract_tables(doc, p)
            if tables:
                print(f"\n테이블 감지 (페이지 {p + 1})")
                print("-" * 60)
                print(f"발견된 테이블: {len(tables)}개")
                for i, t in enumerate(tables, 1):
                    print(f"\n테이블 {i}: {t.rows}행 x {t.cols}열")
                    md = t.to_markdown()
                    lines = md.split('\n')
                    if len(lines) > 10:
                        print('\n'.join(lines[:8]))
                        print(f"   ... ({len(lines) - 8}행 더)")
                    else:
                        print(md)
        return
    
    if args.images:
        images = extract_images(doc)
        print(f"\n이미지: {len(images)}개")
        for i, img in enumerate(images, 1):
            print(f"  {i}. {img.width}x{img.height} {img.color_space} ({len(img.data)} bytes)")

            if args.save:
                os.makedirs(args.output_dir, exist_ok=True)
                filename = os.path.join(args.output_dir, f"img{i}_{img.obj_num}")
                if save_image(img, filename):
                    for ext in ['.jpg', '.jpeg', '.png', '.jp2']:
                        if os.path.exists(filename + ext):
                            print(f"     → 저장: {filename + ext}")
                            break
        return
    
    # 기본: 텍스트 추출
    if args.page is not None:
        text = extract_text(doc, args.page)
        print(f"\n--- 페이지 {args.page + 1} ---")
        print(text)
    else:
        text = extract_all_text(doc)
        print(text)
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"\n저장됨: {args.output}")


def _run_integrated_analysis(doc, args, page_count):
    """통합 분석 실행"""
    from . import (
        extract_images, save_image, extract_tables, 
        analyze_page_layout, extract_text
    )
    
    # 이미지는 문서 전체에서 한 번만 추출
    all_images = extract_images(doc)
    image_pages = _map_images_to_pages(doc, all_images)
    
    # 페이지 범위 결정
    if args.all_text:
        pages_to_analyze = range(page_count)
    elif args.page is not None:
        pages_to_analyze = [args.page]
    else:
        pages_to_analyze = [0]  # 기본: 첫 페이지
    
    for page_num in pages_to_analyze:
        print()
        print("=" * 60)
        print(f"📄 페이지 {page_num + 1} / {page_count}")
        print("=" * 60)
        
        # 1. 레이아웃 분석
        layout = analyze_page_layout(doc, page_num)
        print(f"\n📐 레이아웃: {layout.width:.0f}x{layout.height:.0f}, "
              f"{layout.columns}컬럼, {len(layout.blocks)}블록")
        
        if layout.has_header:
            print("   헤더 있음")
        if layout.has_footer:
            print("   푸터 있음")
        
        # 2. 텍스트 (읽기 순서대로)
        print(f"\n📝 텍스트:")
        print("-" * 40)
        
        for block in layout.get_reading_order():
            block_type = block.block_type.value
            text = block.text.strip()
            if text:
                type_emoji = {
                    'title': '📌',
                    'heading': '📎',
                    'paragraph': '  ',
                    'list_item': '  •',
                    'header': '🔼',
                    'footer': '🔽',
                    'caption': '  ',
                }.get(block_type, '  ')
                
                if len(text) > 80:
                    lines = [text[i:i+76] for i in range(0, len(text), 76)]
                    print(f"{type_emoji} {lines[0]}")
                    for line in lines[1:]:
                        print(f"     {line}")
                else:
                    print(f"{type_emoji} {text}")
        
        # 3. 이미지
        page_images = image_pages.get(page_num, [])
        if page_images:
            print(f"\n🖼️  이미지: {len(page_images)}개")
            for img in page_images:
                print(f"   - {img.width}x{img.height} {img.color_space}")
                
                if args.save:
                    os.makedirs(args.output_dir, exist_ok=True)
                    filename = os.path.join(args.output_dir, 
                                           f"page{page_num+1}_img{img.obj_num}")
                    if save_image(img, filename):
                        for ext in ['.jpg', '.jpeg', '.png', '.jp2']:
                            if os.path.exists(filename + ext):
                                print(f"     → 저장: {filename + ext}")
                                break
        
        # 4. 테이블
        tables = extract_tables(doc, page_num)
        if tables:
            print(f"\n📊 테이블: {len(tables)}개")
            for i, table in enumerate(tables):
                print(f"\n   테이블 {i+1} ({table.rows}x{table.cols}):")
                md = table.to_markdown()
                for line in md.split('\n')[:5]:
                    print(f"   {line}")
                if table.rows > 5:
                    print(f"   ... ({table.rows - 5}행 더)")
    
    print()
    print("=" * 60)
    print("분석 완료")


def _map_images_to_pages(doc, images):
    """이미지를 페이지에 매핑"""
    from . import get_pages, PDFRef
    
    pages = get_pages(doc)
    image_pages = {}
    
    for page_num, page in enumerate(pages):
        resources = page.get('Resources', {})
        if isinstance(resources, PDFRef):
            resources = doc.objects.get((resources.obj_num, resources.gen_num), {})
        
        xobjects = resources.get('XObject', {})
        if isinstance(xobjects, PDFRef):
            xobjects = doc.objects.get((xobjects.obj_num, xobjects.gen_num), {})
        
        page_images = []
        for name, ref in xobjects.items():
            if isinstance(ref, PDFRef):
                for img in images:
                    if img.obj_num == ref.obj_num:
                        page_images.append(img)
                        break
        
        if page_images:
            image_pages[page_num] = page_images
    
    return image_pages


def process_docx(filepath, args):
    """DOCX 처리"""
    from .formats.docx_parser import parse_docx
    
    doc = parse_docx(str(filepath))
    
    # Markdown/JSON 출력
    if args.markdown or args.json:
        from .output_formatter import docx_to_output, to_markdown, to_json
        
        output = docx_to_output(doc, include_images=args.include_images)
        output.filename = str(filepath)
        
        if args.markdown:
            result = to_markdown(output, include_images=args.include_images)
        else:
            result = to_json(output, include_images=args.include_images)
        
        _write_output(args, result)
        return
    
    print(f"DOCX 분석: {filepath}")
    print("=" * 60)
    
    if args.info:
        print(f"제목: {doc.title or '(없음)'}")
        print(f"작성자: {doc.author or '(없음)'}")
        print(f"문단: {len(doc.paragraphs)}개")
        print(f"테이블: {len(doc.tables)}개")
        print(f"이미지: {len(doc.images)}개")
        return
    
    if args.outline:
        print("\n문서 개요:")
        headings = doc.get_headings()
        if headings:
            for level, text in headings:
                print("  " * (level - 1) + f"H{level}: {text}")
        else:
            print("  (헤딩 없음)")
        return
    
    if args.tables:
        print(f"\n테이블: {len(doc.tables)}개")
        for i, t in enumerate(doc.tables, 1):
            print(f"\n테이블 {i}:")
            print(t.to_markdown())
        return
    
    if args.images:
        print(f"\n이미지: {len(doc.images)}개")
        for img in doc.images:
            print(f"  - {img.filename} ({img.content_type})")
        return
    
    text = doc.get_text()
    print(text)
    _save_text_output(args, text)


def process_pptx(filepath, args):
    """PPTX 처리"""
    from .formats.pptx_parser import parse_pptx
    
    doc = parse_pptx(str(filepath))
    
    # Markdown/JSON 출력
    if args.markdown or args.json:
        from .output_formatter import pptx_to_output, to_markdown, to_json
        
        output = pptx_to_output(doc, include_images=args.include_images)
        output.filename = str(filepath)
        
        if args.markdown:
            result = to_markdown(output, include_images=args.include_images)
        else:
            result = to_json(output, include_images=args.include_images)
        
        _write_output(args, result)
        return
    
    print(f"PPTX 분석: {filepath}")
    print("=" * 60)
    print(f"슬라이드: {doc.slide_count}개")
    
    if args.info:
        print(f"제목: {doc.title or '(없음)'}")
        print(f"작성자: {doc.author or '(없음)'}")
        print(f"이미지: {len(doc.images)}개")
        return
    
    if args.outline:
        print("\n슬라이드 목록:")
        for slide in doc.slides:
            print(f"  {slide.number}. {slide.title or '(제목 없음)'}")
        return
    
    if args.tables:
        for slide in doc.slides:
            if slide.tables:
                print(f"\n슬라이드 {slide.number} 테이블: {len(slide.tables)}개")
                for i, t in enumerate(slide.tables, 1):
                    print(f"\n테이블 {i}:")
                    print(t.to_markdown())
        return
    
    if args.images:
        print(f"\n이미지: {len(doc.images)}개")
        for img in doc.images:
            print(f"  - {img.filename} ({img.content_type})")
        return
    
    if args.page is not None:
        if 0 <= args.page < len(doc.slides):
            slide = doc.slides[args.page]
            print(f"\n--- 슬라이드 {slide.number} ---")
            print(slide.get_text())
        else:
            print(f"오류: 슬라이드 번호 범위 초과 (0-{len(doc.slides) - 1})")
        return
    
    text = doc.get_text()
    print(text)
    _save_text_output(args, text)


def process_hwpx(filepath, args):
    """HWPX 처리"""
    from .formats.hwpx_parser import parse_hwpx
    
    doc = parse_hwpx(str(filepath))
    
    # Markdown/JSON 출력
    if args.markdown or args.json:
        from .output_formatter import hwpx_to_output, to_markdown, to_json
        
        output = hwpx_to_output(doc, include_images=args.include_images)
        output.filename = str(filepath)
        
        if args.markdown:
            result = to_markdown(output, include_images=args.include_images)
        else:
            result = to_json(output, include_images=args.include_images)
        
        _write_output(args, result)
        return
    
    print(f"HWPX 분석: {filepath}")
    print("=" * 60)
    
    if args.info:
        print(f"제목: {doc.title or '(없음)'}")
        print(f"작성자: {doc.author or '(없음)'}")
        print(f"문단: {len(doc.paragraphs)}개")
        print(f"테이블: {len(doc.tables)}개")
        print(f"이미지: {len(doc.images)}개")
        return
    
    if args.outline:
        print("\n문서 개요:")
        headings = doc.get_headings()
        if headings:
            for level, text in headings:
                print("  " * (level - 1) + f"H{level}: {text}")
        else:
            print("  (헤딩 없음)")
        return
    
    if args.tables:
        print(f"\n테이블: {len(doc.tables)}개")
        for i, t in enumerate(doc.tables, 1):
            print(f"\n테이블 {i}:")
            print(t.to_markdown())
        return
    
    if args.images:
        print(f"\n이미지: {len(doc.images)}개")
        for img in doc.images:
            print(f"  - {img.filename} ({img.content_type})")
        return
    
    text = doc.get_text()
    print(text)
    _save_text_output(args, text)


def process_text(filepath, args):
    """TXT/MD 처리"""
    from .formats.text_parser import parse_text, parse_markdown, extract_text as txt_extract
    
    ext = filepath.suffix.lower()
    is_markdown = ext in ['.md', '.markdown']
    
    if is_markdown:
        doc = parse_markdown(str(filepath))
    else:
        doc = parse_text(str(filepath))
    
    # Markdown/JSON 출력
    if args.markdown or args.json:
        from .output_formatter import text_to_output, to_markdown, to_json
        
        output = text_to_output(doc, is_markdown=is_markdown)
        output.filename = str(filepath)
        
        if args.markdown:
            result = to_markdown(output)
        else:
            result = to_json(output)
        
        _write_output(args, result)
        return
    
    print(f"{'마크다운' if is_markdown else '텍스트'} 분석: {filepath}")
    print("=" * 60)
    
    if args.info:
        print(f"줄 수: {len(doc.lines)}")
        print(f"인코딩: {doc.encoding}")
        if is_markdown and doc.headings:
            print(f"헤딩: {len(doc.headings)}개")
            print(f"코드블록: {len(doc.code_blocks)}개")
            print(f"링크: {len(doc.links)}개")
            print(f"이미지: {len(doc.images)}개")
        return
    
    if args.outline and is_markdown and doc.headings:
        print("\n문서 개요:")
        for level, text in doc.headings:
            print("  " * (level - 1) + f"H{level}: {text}")
        return
    
    if is_markdown and not args.text:
        text = doc.content
    else:
        text = txt_extract(doc) if is_markdown else doc.content
    print(text)
    _save_text_output(args, text)


def _write_output(args, content):
    """출력 (파일 또는 stdout)"""
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"저장됨: {args.output}", file=sys.stderr)
    else:
        print(content)


def _save_text_output(args, text):
    """텍스트 출력 파일 저장"""
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"\n저장됨: {args.output}")


if __name__ == '__main__':
    main()
