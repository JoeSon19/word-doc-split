"""
Word 문서(.docx)를 지정된 페이지 수 단위로 분할하는 스크립트.
Windows + Microsoft Word 설치 필요 (win32com 사용).
"""

import argparse
import io
import os
import sys
import time

# 한글 출력을 위한 UTF-8 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

try:
    import win32com.client
    from win32com.client import constants
except ImportError:
    print("오류: pywin32가 설치되어 있지 않습니다.")
    print("  pip install pywin32")
    sys.exit(1)


def split_word_document(input_path: str, pages_per_split: int = 100, output_dir: str | None = None):
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        print(f"오류: 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    if output_dir is None:
        output_dir = os.path.dirname(input_path)
    os.makedirs(output_dir, exist_ok=True)

    word = None
    try:
        print("Word 애플리케이션을 시작합니다...")
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        print(f"문서를 열고 있습니다: {input_path}")
        doc = word.Documents.Open(input_path, ReadOnly=True)

        # 레이아웃 계산을 위해 잠시 대기
        word.ActiveWindow.View.Type = 3  # wdPrintView
        time.sleep(2)

        total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
        total_parts = (total_pages + pages_per_split - 1) // pages_per_split

        print(f"총 페이지 수: {total_pages}")
        print(f"분할 단위: {pages_per_split}페이지")
        print(f"생성될 파일 수: {total_parts}")
        print("-" * 40)

        for part_idx in range(total_parts):
            start_page = part_idx * pages_per_split + 1
            end_page = min((part_idx + 1) * pages_per_split, total_pages)

            part_num = f"{part_idx + 1:03d}"
            output_path = os.path.join(output_dir, f"{base_name}_{part_num}.docx")

            # 해당 페이지 범위를 선택
            rng_start = doc.GoTo(
                What=1,  # wdGoToPage
                Which=1,  # wdGoToAbsolute
                Count=start_page
            )

            if end_page < total_pages:
                rng_end = doc.GoTo(
                    What=1,  # wdGoToPage
                    Which=1,  # wdGoToAbsolute
                    Count=end_page + 1
                )
                # 다음 페이지 시작 직전까지 선택
                rng_start.End = rng_end.Start - 1
            else:
                # 마지막 파트: 문서 끝까지
                rng_start.End = doc.Content.End

            # 선택 영역을 복사하여 새 문서에 붙여넣기
            rng_start.Copy()

            new_doc = word.Documents.Add()
            # 기본 빈 단락 제거 후 붙여넣기
            new_doc.Content.Delete()
            new_doc.Content.Paste()

            # 원본 문서의 페이지 설정 복사
            for i in range(1, doc.Sections.Count + 1):
                if i <= new_doc.Sections.Count:
                    try:
                        src = doc.Sections(i).PageSetup
                        dst = new_doc.Sections(min(i, new_doc.Sections.Count)).PageSetup
                        dst.TopMargin = src.TopMargin
                        dst.BottomMargin = src.BottomMargin
                        dst.LeftMargin = src.LeftMargin
                        dst.RightMargin = src.RightMargin
                        dst.PageWidth = src.PageWidth
                        dst.PageHeight = src.PageHeight
                        dst.Orientation = src.Orientation
                    except Exception:
                        pass
                    break  # 첫 번째 섹션 설정만 복사

            new_doc.SaveAs2(os.path.abspath(output_path), FileFormat=12)  # 12 = wdFormatXMLDocument (.docx)
            new_doc.Close(SaveChanges=0)

            print(f"파트 {part_idx + 1}/{total_parts} 저장 완료: {os.path.basename(output_path)} (p.{start_page}-{end_page})")

        doc.Close(SaveChanges=0)
        print("-" * 40)
        print(f"완료! {total_parts}개 파일이 '{output_dir}'에 저장되었습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")
        sys.exit(1)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(description="Word 문서(.docx)를 페이지 단위로 분할합니다.")
    parser.add_argument("input", help="입력 .docx 파일 경로")
    parser.add_argument("-p", "--pages", type=int, default=100, help="분할 단위 페이지 수 (기본값: 100)")
    parser.add_argument("-o", "--output-dir", default=None, help="출력 디렉토리 (기본값: 입력 파일과 같은 위치)")
    args = parser.parse_args()

    split_word_document(args.input, args.pages, args.output_dir)


if __name__ == "__main__":
    main()
