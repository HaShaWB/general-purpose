import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Checkbutton, IntVar
import os
import PyPDF2

def extract_bookmarks_to_pdf(pdf_path, output_dir, max_depth=2, add_index=False):
    # 원본 pdf 파일명에서 확장자 제거
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    # base_name 폴더 생성
    target_dir = os.path.join(output_dir, base_name)
    os.makedirs(target_dir, exist_ok=True)
    
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)

        # outine(북마크) 가져오기
        outlines = reader.outline

        # 북마크 리스트 (사용자 설정 depth까지)
        bookmarks = []
        
        def parse_outlines(outline_items, depth=1):
            for item in outline_items:
                if isinstance(item, list):
                    # 사용자가 설정한 depth까지만 파싱
                    if depth < max_depth:
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
                
            # 인덱스 추가 옵션이 활성화된 경우 파일명 앞에 인덱스 추가
            if add_index:
                safe_title = f"{i+1:02d}_{safe_title}"

            start_page_num = reader.get_destination_page_number(bookmark)
            if i < len(bookmarks) - 1:
                next_bookmark = bookmarks[i+1]
                end_page_num = reader.get_destination_page_number(next_bookmark)
            else:
                end_page_num = len(reader.pages)

            writer = PyPDF2.PdfWriter()
            
            # 페이지 추가
            for p in range(start_page_num, end_page_num):
                writer.add_page(reader.pages[p])
            
            # 해당 범위에 포함된 모든 북마크 추가
            bookmark_tree = extract_bookmark_hierarchy(reader)
            add_bookmarks_to_writer(writer, bookmark_tree, start_page_num, end_page_num)

            output_file = os.path.join(target_dir, f"{safe_title}.pdf")
            with open(output_file, 'wb') as out_f:
                writer.write(out_f)


def extract_bookmark_hierarchy(pdf_reader, max_depth=5):
    """PDF의 북마크 계층 구조를 추출하는 함수"""
    bookmark_tree = []
    
    def process_bookmark_list(outline_items, current_depth=0, parent=None):
        result = []
        i = 0
        while i < len(outline_items):
            item = outline_items[i]
            
            if isinstance(item, list):
                # 이전 항목의 하위 목록인 경우 스킵
                i += 1
                continue
                
            # 북마크 정보
            page_num = pdf_reader.get_destination_page_number(item)
            bookmark_info = {
                "title": item.title,
                "page": page_num,
                "depth": current_depth,
                "children": []
            }
            
            # 하위 북마크가 있는지 확인
            if i + 1 < len(outline_items) and isinstance(outline_items[i + 1], list):
                # 하위 계층 처리
                children = process_bookmark_list(outline_items[i + 1], current_depth + 1, bookmark_info)
                bookmark_info["children"] = children
                i += 2  # 현재 항목과 하위 목록 건너뛰기
            else:
                i += 1
                
            result.append(bookmark_info)
            
        return result
        
    # 최상위 수준의 북마크 처리
    if pdf_reader.outline:
        bookmark_tree = process_bookmark_list(pdf_reader.outline)
        
    return bookmark_tree


def add_bookmarks_to_writer(writer, bookmark_tree, start_page, end_page, parent=None):
    """북마크 트리를 PDF 작성기에 추가하는 함수"""
    for bookmark in bookmark_tree:
        page_num = bookmark["page"]
        
        # 현재 범위에 속하는 북마크만 추가
        if start_page <= page_num < end_page:
            # 상대적 페이지 번호로 변환
            rel_page = page_num - start_page
            
            # 북마크 추가
            bookmark_ref = writer.add_outline_item(bookmark["title"], rel_page, parent=parent)
            
            # 하위 북마크도 추가
            if bookmark["children"]:
                add_bookmarks_to_writer(writer, bookmark["children"], start_page, end_page, bookmark_ref)


if __name__ == "__main__":
    # 메인 창 생성
    root = tk.Tk()
    root.title("PDF 북마크 분할기")
    root.geometry("400x200")
    
    # 파일 선택 함수
    def select_file():
        file_path = filedialog.askopenfilename(title="PDF 선택", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, file_path)
    
    # 디렉토리 선택 함수
    def select_directory():
        dir_path = filedialog.askdirectory(title="저장할 기본 폴더 선택")
        if dir_path:
            dir_entry.delete(0, tk.END)
            dir_entry.insert(0, dir_path)
    
    # 실행 함수
    def run_extraction():
        pdf_path = file_entry.get()
        output_dir = dir_entry.get()
        
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("오류", "유효한 PDF 파일을 선택하세요.")
            return
            
        if not output_dir or not os.path.exists(output_dir):
            messagebox.showerror("오류", "유효한 저장 폴더를 선택하세요.")
            return
        
        max_depth = simpledialog.askinteger("Depth 설정", "북마크 Depth를 설정하세요 (1-5):", 
                                           minvalue=1, maxvalue=5, initialvalue=2)
        if max_depth is None:
            max_depth = 2  # 취소 시 기본값 2로 설정
        
        add_index = index_var.get() == 1  # 체크박스 상태 확인
        
        try:
            extract_bookmarks_to_pdf(pdf_path, output_dir, max_depth, add_index)
            messagebox.showinfo("완료", "PDF 분할이 완료되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"에러 발생: {e}")
    
    # 레이아웃 구성
    tk.Label(root, text="PDF 파일:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
    file_entry = tk.Entry(root, width=30)
    file_entry.grid(row=0, column=1, padx=5, pady=10)
    tk.Button(root, text="찾아보기", command=select_file).grid(row=0, column=2, padx=5, pady=10)
    
    tk.Label(root, text="저장 폴더:").grid(row=1, column=0, sticky="w", padx=10, pady=10)
    dir_entry = tk.Entry(root, width=30)
    dir_entry.grid(row=1, column=1, padx=5, pady=10)
    tk.Button(root, text="찾아보기", command=select_directory).grid(row=1, column=2, padx=5, pady=10)
    
    # 인덱스 추가 체크박스
    index_var = IntVar()
    index_checkbox = Checkbutton(root, text="파일명에 인덱스 추가", variable=index_var)
    index_checkbox.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="w")
    
    # 실행 버튼
    tk.Button(root, text="PDF 분할하기", command=run_extraction, bg="#4CAF50", fg="white", 
             height=2).grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
    
    root.mainloop()
