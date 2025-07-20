import os
import re
import tkinter as tk
from tkinter import filedialog, ttk
import openpyxl

class FileSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("파일 검색 도구")
        self.root.geometry("1000x700")

        # 검색 조건 프레임
        condition_frame = tk.Frame(root, padx=10, pady=10)
        condition_frame.pack(fill=tk.X)

        # 폴더 선택
        tk.Label(condition_frame, text="검색할 폴더:").grid(row=0, column=0, sticky='w')
        self.folder_var = tk.StringVar()
        tk.Entry(condition_frame, textvariable=self.folder_var, width=80).grid(row=0, column=1, padx=5)
        tk.Button(condition_frame, text="폴더 선택", command=self.browse_folder).grid(row=0, column=2)

        # 찾을 파일명 목록
        tk.Label(condition_frame, text="찾을 파일명 목록:").grid(row=1, column=0, sticky='nw')
        self.filenames_text = tk.Text(condition_frame, width=80, height=5)
        self.filenames_text.grid(row=1, column=1, padx=5, pady=5)

        # 확장자 선택 및 무시 옵션
        ext_frame = tk.Frame(condition_frame)
        ext_frame.grid(row=1, column=2, sticky='nw', pady=5)
        self.ext_cs = tk.BooleanVar(value=True)
        self.ext_regx = tk.BooleanVar(value=True)
        self.ext_java = tk.BooleanVar(value=True)
        self.ignore_ext = tk.BooleanVar()
        tk.Checkbutton(ext_frame, text=".cs", variable=self.ext_cs).pack(anchor='w')
        tk.Checkbutton(ext_frame, text=".regx", variable=self.ext_regx).pack(anchor='w')
        tk.Checkbutton(ext_frame, text=".java", variable=self.ext_java).pack(anchor='w')
        tk.Checkbutton(ext_frame, text="확장자 무시", variable=self.ignore_ext, command=self.toggle_ext_checkboxes).pack(anchor='w')

        # 대소문자 구분 체크박스
        self.case_sensitive = tk.BooleanVar()
        tk.Checkbutton(condition_frame, text="대소문자 구분", variable=self.case_sensitive).grid(row=2, column=2, sticky='w')

        # 문자 검색어
        tk.Label(condition_frame, text="문자 검색어 (,로 구분):").grid(row=2, column=0, sticky='w')
        self.keywords_var = tk.StringVar()
        tk.Entry(condition_frame, textvariable=self.keywords_var, width=80).grid(row=2, column=1, padx=5, pady=5)

        # 검색 버튼
        tk.Button(root, text="🔍 검색 시작", command=self.search_files, height=2, bg="lightblue").pack(fill=tk.X, padx=10, pady=5)

        # 결과 출력 테이블
        result_frame = tk.Frame(root, padx=10, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("파일명", "경로", "크기(KB)", "문자 검색 수")
        self.result_list = ttk.Treeview(result_frame, columns=columns, show='headings')
        for col in columns:
            self.result_list.heading(col, text=col)
            self.result_list.column(col, width=200, anchor='w')
        self.result_list.pack(fill=tk.BOTH, expand=True)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_var.set(folder_selected)

    def toggle_ext_checkboxes(self):
        state = tk.DISABLED if self.ignore_ext.get() else tk.NORMAL
        for var in [self.ext_cs, self.ext_regx, self.ext_java]:
            var.set(False if self.ignore_ext.get() else True)

    def search_files(self):
        self.result_list.delete(*self.result_list.get_children())

        base_dir = self.folder_var.get()
        if not base_dir:
            return

        input_names = self.filenames_text.get("1.0", tk.END).strip().splitlines()
        search_keywords = [kw.strip() for kw in self.keywords_var.get().split(',') if kw.strip()]

        extensions = []
        if self.ext_cs.get(): extensions.append(".cs")
        if self.ext_regx.get(): extensions.append(".regx")
        if self.ext_java.get(): extensions.append(".java")

        results = []
        for dirpath, _, filenames in os.walk(base_dir):
            for filename in filenames:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, base_dir)
                name_only, ext = os.path.splitext(filename)

                if not self.ignore_ext.get() and ext not in extensions:
                    continue

                match = False
                for target in input_names:
                    if not target.strip():
                        continue

                    is_path = '\\' in target or '/' in target
                    norm_target = target.replace('/', os.sep).replace('\\', os.sep)

                    if self.ignore_ext.get():
                        cmp_filename = os.path.splitext(rel_path if is_path else filename)[0]
                        cmp_target = os.path.splitext(norm_target)[0]
                    else:
                        cmp_filename = rel_path if is_path else filename
                        cmp_target = norm_target

                    if not self.case_sensitive.get():
                        cmp_filename = cmp_filename.lower()
                        cmp_target = cmp_target.lower()

                    if cmp_target in cmp_filename:
                        match = True
                        break

                if not match:
                    continue

                keyword_counts = {}
                try:
                    with open(full_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                        for kw in search_keywords:
                            count = content.count(kw) if self.case_sensitive.get() else content.lower().count(kw.lower())
                            keyword_counts[kw] = count
                except Exception:
                    continue

                display_name = filename if not self.ignore_ext.get() else os.path.splitext(filename)[0]
                total_size_kb = os.path.getsize(full_path) // 1024
                keyword_display = ' / '.join([f"{k}({v})" for k, v in keyword_counts.items()])

                results.append((display_name, rel_path, total_size_kb, keyword_display))

        for row in results:
            self.result_list.insert('', tk.END, values=row)

if __name__ == '__main__':
    root = tk.Tk()
    app = FileSearchApp(root)
    root.mainloop()
