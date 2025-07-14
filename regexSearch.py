import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_var.set(folder_selected)

def search_files():
    folder = folder_var.get()
    pattern = regex_var.get()
    result_list.delete(*result_list.get_children())
    results.clear()
    progress_var.set(0)
    progress_label.config(text="진행률: 0%")

    if not folder or not pattern:
        messagebox.showwarning("입력 필요", "폴더와 정규표현식을 입력해주세요.")
        return

    try:
        regex = re.compile(pattern)
    except re.error as e:
        messagebox.showerror("정규표현식 오류", f"유효하지 않은 정규표현식입니다:\n{e}")
        return

    # 파일 목록 수집
    matched_files = []
    for dirpath, _, files in os.walk(folder):
        for file in files:
            if file.endswith(('.py', '.java', '.cs', '.txt')):
                matched_files.append(os.path.join(dirpath, file))

    total_files = len(matched_files)
    if total_files == 0:
        progress_label.config(text="검색할 파일이 없습니다.")
        return

    # 파일 검색
    for idx, full_path in enumerate(matched_files):
        try:
            with open(full_path, encoding='utf-8', errors='ignore') as f:
                for i, line in enumerate(f, 1):
                    if regex.search(line):
                        result = (os.path.basename(full_path), i, line.strip(), full_path)
                        results.append(result)
                        result_list.insert('', 'end', values=result)
        except Exception as e:
            print(f"파일 읽기 오류: {full_path}: {e}")

        percent = int((idx + 1) / total_files * 100)
        progress_var.set(percent)
        progress_label.config(text=f"진행률: {percent}%")
        root.update_idletasks()

    adjust_column_widths()
    progress_label.config(text="검색 완료")

def export_to_excel():
    if not results:
        messagebox.showinfo("알림", "출력할 결과가 없습니다.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "검색 결과"
    ws.append(['파일명', '라인', '내용', '전체경로'])

    for row in results:
        ws.append(row)

    wb.save(save_path)
    messagebox.showinfo("저장 완료", f"엑셀 파일로 저장되었습니다:\n{save_path}")

def adjust_column_widths():
    col_max_width = [len(h) for h in cols]
    for row in results:
        for i, value in enumerate(row):
            length = len(str(value))
            if length > col_max_width[i]:
                col_max_width[i] = length

    for i, col in enumerate(cols):
        width = col_max_width[i] * 7 + 20
        result_list.column(col, width=width)

# GUI 구성
root = tk.Tk()
root.title("🔍 정규표현식 소스 검색기")
root.geometry("1000x600")
root.minsize(800, 500)

folder_var = tk.StringVar()
regex_var = tk.StringVar()
results = []

# 창 조정 가능하게
root.columnconfigure(1, weight=1)
root.rowconfigure(2, weight=1)

# 폴더 입력
tk.Label(root, text="📁 폴더 경로").grid(row=0, column=0, sticky="w", padx=5, pady=5)
tk.Entry(root, textvariable=folder_var).grid(row=0, column=1, sticky="ew", padx=5)
tk.Button(root, text="찾아보기", command=browse_folder).grid(row=0, column=2, padx=5)

# 정규표현식 입력
tk.Label(root, text="🔍 정규표현식").grid(row=1, column=0, sticky="w", padx=5)
tk.Entry(root, textvariable=regex_var).grid(row=1, column=1, sticky="ew", padx=5)
tk.Button(root, text="검색", command=search_files).grid(row=1, column=2, padx=5)

# 결과 테이블
cols = ('파일명', '라인', '내용', '전체경로')
result_list = ttk.Treeview(root, columns=cols, show='headings')
for col in cols:
    result_list.heading(col, text=col)
    result_list.column(col, anchor="w")

# 스크롤바
scroll_y = ttk.Scrollbar(root, orient='vertical', command=result_list.yview)
scroll_x = ttk.Scrollbar(root, orient='horizontal', command=result_list.xview)
result_list.configure(yscroll=scroll_y.set, xscroll=scroll_x.set)

# 결과 테이블 위치
result_list.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)
scroll_y.grid(row=2, column=3, sticky="ns")
scroll_x.grid(row=3, column=0, columnspan=3, sticky="ew")

# 진행률 표시
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=4, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

progress_label = tk.Label(root, text="진행률: 0%")
progress_label.grid(row=4, column=2, sticky="e", padx=10)

# 엑셀 저장 버튼
tk.Button(root, text="엑셀로 저장", command=export_to_excel).grid(row=5, column=2, sticky="e", pady=10)

root.mainloop()
