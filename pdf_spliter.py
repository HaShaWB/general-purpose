import tkinter as tk
from tkinter import filedialog, messagebox
import os
import PyPDF2

def extract_bookmarks_to_pdf(pdf_path, output_dir):
    # 원본 pdf 파일명에서 확장자 제거
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    # base_name 폴더 생성
    target_dir = os.path.join(output_dir, base_name)
    os.makedirs(target_dir, exist_ok=True)
    
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)

        # outine(북마크) 가져오기
        outlines = reader.outline

        # 북마크 리스트 (최대 2단계까지)
        bookmarks = []
        
        def parse_outlines(outline_items, depth=1):
            for item in outline_items:
                if isinstance(item, list):
                    # depth 2까지만 파싱
                    if depth < 2:
                        parse_outlines(item, depth+1)
                else:
                    # 여기가 실제 북마크(목차 항목)
                    bookmarks.append(item)

        parse_outlines(outlines, depth=1)

        # 각 북마크 구간별 PDF 분할 저장
        for i, bookmark in enumerate(bookmarks):
            title = bookmark.title

            # 파일명에 문제될 수 있는 특수문자 제거
            safe_title = ''.join(c for c in title if c.isalnum() or c in (' ', '_', '-')).strip()
            if not safe_title:
                safe_title = f"untitled_{i}"

            start_page_num = reader.get_destination_page_number(bookmark)
            if i < len(bookmarks) - 1:
                next_bookmark = bookmarks[i+1]
                end_page_num = reader.get_destination_page_number(next_bookmark)
            else:
                end_page_num = len(reader.pages)

            writer = PyPDF2.PdfWriter()
            for p in range(start_page_num, end_page_num):
                writer.add_page(reader.pages[p])

            output_file = os.path.join(target_dir, f"{safe_title}.pdf")
            with open(output_file, 'wb') as out_f:
                writer.write(out_f)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(title="PDF 선택", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_path:
        messagebox.showerror("오류", "PDF 파일을 선택하지 않았습니다.")
        exit()

    output_dir = filedialog.askdirectory(title="저장할 기본 폴더 선택")
    if not output_dir:
        messagebox.showerror("오류", "저장할 폴더를 선택하지 않았습니다.")
        exit()

    try:
        extract_bookmarks_to_pdf(pdf_path, output_dir)
        print("ALL DONE")
    except Exception as e:
        messagebox.showerror("오류", f"에러 발생: {e}")
