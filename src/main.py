import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

class ExcelFillerGUI:
    def __init__(self, root):
        self.root = root
        # 初始化語言變數（預設繁中）
        self.current_lang = tk.StringVar(value="Zh")
        # 語言字典：繁中/英文對照
        self.lang_dict = {
            "Zh": {
                "title": "Excel數據匹配填充工具",
                "file_frame": "1. 文件選擇",
                "file1_label": "數據源文件（表1）：",
                "file2_label": "目標文件（表2）：",
                "browse_btn": "瀏覽",
                "sheet_label": "工作表：",
                "load_cols_btn": "加載列名",
                "output_label": "輸出文件：",
                "save_path_btn": "保存位置",
                "match_frame": "2. 匹配列配置（必選：按此列關聯數據 | 支持多列匹配）",
                "match1_label": "表1匹配列：",
                "match2_label": "表2匹配列：",
                "confirm_btn": "確認選擇",
                "add_match_btn": "添加匹配列對",
                "remove_match_btn": "移除選中匹配列",
                "clear_match_btn": "清空匹配列",
                "selected_match": "已選匹配列：",
                "fill_frame": "3. 填充列配置（表1列 → 表2列）",
                "fill1_label": "表1填充列：",
                "fill2_label": "表2填充列：",
                "add_fill_btn": "添加填充列",
                "remove_fill_btn": "移除選中填充列",
                "clear_fill_btn": "清空填充列",
                "selected_fill": "已選填充列：",
                "preview_btn": "✅ 預覽數據",
                "run_btn": "🚀 執行填充",
                "reset_btn": "🔄 重置所有",
                "lang_select": "語言 / Language",
                "no_select": "未選擇",
                "confirm_match1": "表1匹配列已選擇：",
                "confirm_match2": "表2匹配列已選擇：",
                "confirm_fill1": "表1填充列已選擇：",
                "confirm_fill2": "表2填充列已選擇：",
                "add_match_success": "匹配列對已添加：{} → {}",
                "add_fill_success": "填充列對已添加：{} → {}",
                "remove_match_success": "已移除 {} 組匹配列對",
                "remove_fill_success": "已移除 {} 組填充列對",
                "clear_match_success": "已清空所有匹配列對",
                "clear_fill_success": "已清空所有填充列對",
                "warn_no_file": "請選擇表1和表2文件！",
                "warn_no_match": "請先配置匹配列！",
                "warn_no_fill": "請先配置填充列！",
                "warn_no_col1": "請先從下拉框選擇列名！",
                "warn_no_col2": "請先選擇並確認表1填充列！",
                "warn_no_col3": "請先選擇並確認表2填充列！",
                "warn_match_col_count": "表1和表2選擇的列數量必須相同！",
                "warn_no_selected_match": "請先選擇要移除的匹配列對！",
                "warn_no_selected_fill": "請先選擇要移除的填充列對！",
                "success_load_file": "{}加載完成\n工作表：{}",
                "success_load_cols": "{}列名加載完成\n列名：{}",
                "success_fill": "🎉 填充完成",
                "fill_result": "結果已保存至：{}\n表2原行數：{}\n填充後行數：{}\n匹配列：{}\n填充列：{}",
                "reset_success": "所有配置已重置為初始狀態",
                "error_load_file": "加載文件失敗：{}\n建議：檢查文件是否損壞/關閉Excel後重試",
                "error_load_cols": "加載列名失敗：{}\n建議：檢查文件格式/工作表名是否正確",
                "error_preview": "預覽失敗：{}",
                "error_fill": "填充失敗：{}\n建議：檢查文件是否被佔用/列名是否正確",
                "table1": "表1",
                "table2": "表2",
                "preview_title": "數據預覽（前10行）",
                "select_output_title": "選擇結果保存位置"
            },
            "en": {
                "title": "Excel Data Matching & Filling Tool",
                "file_frame": "1. File Selection",
                "file1_label": "Source File (Table 1)：",
                "file2_label": "Target File (Table 2)：",
                "browse_btn": "Browse",
                "sheet_label": "Worksheet：",
                "load_cols_btn": "Load Columns",
                "output_path_btn": "Save Path",
                "match_frame": "2. Match Column Config (Required: Link data by columns | Support multi-column match)",
                "match1_label": "Table 1 Match Column：",
                "match2_label": "Table 2 Match Column：",
                "confirm_btn": "Confirm Selection",
                "add_match_btn": "Add Match Pair",
                "remove_match_btn": "Remove Selected Match",
                "clear_match_btn": "Clear All Match",
                "selected_match": "Selected Match Columns：",
                "fill_frame": "3. Fill Column Config (Table 1 → Table 2)",
                "fill1_label": "Table 1 Fill Column：",
                "fill2_label": "Table 2 Fill Column：",
                "add_fill_btn": "Add Fill Column",
                "remove_fill_btn": "Remove Selected Fill",
                "clear_fill_btn": "Clear All Fill",
                "selected_fill": "Selected Fill Columns：",
                "preview_btn": "✅ Preview Data",
                "run_btn": "🚀 Run Filling",
                "reset_btn": "🔄 Reset All",
                "lang_select": "Language / 語言",
                "no_select": "Not Selected",
                "confirm_match1": "Table 1 Match Column Selected：",
                "confirm_match2": "Table 2 Match Column Selected：",
                "confirm_fill1": "Table 1 Fill Column Selected：",
                "confirm_fill2": "Table 2 Fill Column Selected：",
                "add_match_success": "Match Pair Added：{} → {}",
                "add_fill_success": "Fill Column Pair Added：{} → {}",
                "remove_match_success": "Removed {} match pairs",
                "remove_fill_success": "Removed {} fill pairs",
                "clear_match_success": "All match pairs cleared",
                "clear_fill_success": "All fill pairs cleared",
                "warn_no_file": "Please select Table 1 and Table 2 files！",
                "warn_no_match": "Please configure match columns first！",
                "warn_no_fill": "Please configure fill columns first！",
                "warn_no_col1": "Please select a column from the dropdown first！",
                "warn_no_col2": "Please select and confirm Table 1 fill column first！",
                "warn_no_col3": "Please select and confirm Table 2 fill column first！",
                "warn_match_col_count": "The number of selected columns in Table 1 and Table 2 must be the same！",
                "warn_no_selected_match": "Please select match pairs to remove first！",
                "warn_no_selected_fill": "Please select fill pairs to remove first！",
                "success_load_file": "{} loaded successfully\nWorksheets：{}",
                "success_load_cols": "{} columns loaded successfully\nColumns：{}",
                "success_fill": "🎉 Filling Completed",
                "fill_result": "Result saved to：{}\nOriginal Table 2 rows：{}\nFilled rows：{}\nMatch Columns：{}\nFill Columns：{}",
                "reset_success": "All configurations reset to initial state",
                "error_load_file": "Failed to load file：{}\nSuggestion：Check if file is damaged / Close Excel and try again",
                "error_load_cols": "Failed to load columns：{}\nSuggestion：Check file format / Worksheet name",
                "error_preview": "Preview failed：{}",
                "error_fill": "Filling failed：{}\nSuggestion：Check if file is occupied / Column names are correct",
                "table1": "Table 1",
                "table2": "Table 2",
                "preview_title": "Data Preview (First 10 Rows)",
                "select_output_title": "Select Result Save Location"
            }
        }
        # 綁定語言變化事件
        self.current_lang.trace_add("write", self.update_all_texts)
        
        # 初始化變量
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar(value="填充結果.xlsx" if self.current_lang.get() == "Zh" else "fill_result.xlsx")
        self.sheet1_name = tk.StringVar()
        self.sheet2_name = tk.StringVar()
        self.cols1 = []
        self.cols2 = []
        self.sheets1 = []
        self.sheets2 = []
        
        # 多列匹配支持：改為列表存儲多組匹配列對
        self.match_pairs = []  # 格式：[(col1_1, col2_1), (col1_2, col2_2), ...]
        self.fill_pairs = []   # 格式：[(col1_1, col2_1), (col1_2, col2_2), ...]
        
        # 臨時選擇變量（支持多選）
        self.match1_var = tk.StringVar()
        self.match2_var = tk.StringVar()
        self.fill1_var = tk.StringVar()
        self.fill2_var = tk.StringVar()
        self.fill1_selected = ""
        self.fill2_selected = ""
        
        # 創建界面
        self.create_widgets()
        # 初始化文本
        self.update_all_texts()

    def create_widgets(self):
        # ========== 語言切換控件 ==========
        lang_frame = ttk.Frame(self.root, padding=5)
        lang_frame.pack(fill="x", padx=10, pady=5, anchor="e")
        self.lang_combo = ttk.Combobox(
            lang_frame, 
            textvariable=self.current_lang,
            values=["Zh", "en"],
            state="readonly",
            width=10
        )
        self.lang_combo.grid(row=0, column=0, padx=5)
        self.lang_label = ttk.Label(lang_frame, text=self.lang_dict[self.current_lang.get()]["lang_select"])
        self.lang_label.grid(row=0, column=1, padx=5)

        # ========== 1. 文件選擇區域 ==========
        self.file_frame = ttk.LabelFrame(self.root, padding=15)
        self.file_frame.pack(fill="x", padx=10, pady=8)

        # 表1選擇
        self.file1_label = ttk.Label(self.file_frame, font=("Arial", 10))
        self.file1_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(self.file_frame, textvariable=self.file1_path, width=45).grid(row=0, column=1, padx=5, pady=5)
        self.browse1_btn = ttk.Button(self.file_frame, command=lambda: self.load_file(True), width=8)
        self.browse1_btn.grid(row=0, column=2, padx=5, pady=5)
        
        self.sheet1_label = ttk.Label(self.file_frame, font=("Arial", 10))
        self.sheet1_label.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        self.sheet1_combo = ttk.Combobox(self.file_frame, textvariable=self.sheet1_name, width=12, state="readonly")
        self.sheet1_combo.grid(row=0, column=4, padx=5, pady=5)
        self.load_cols1_btn = ttk.Button(self.file_frame, command=lambda: self.load_column(True), width=10)
        self.load_cols1_btn.grid(row=0, column=5, padx=5, pady=5)

        # 表2選擇
        self.file2_label = ttk.Label(self.file_frame, font=("Arial", 10))
        self.file2_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(self.file_frame, textvariable=self.file2_path, width=45).grid(row=1, column=1, padx=5, pady=5)
        self.browse2_btn = ttk.Button(self.file_frame, command=lambda: self.load_file(False), width=8)
        self.browse2_btn.grid(row=1, column=2, padx=5, pady=5)
        
        self.sheet2_label = ttk.Label(self.file_frame, font=("Arial", 10))
        self.sheet2_label.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        self.sheet2_combo = ttk.Combobox(self.file_frame, textvariable=self.sheet2_name, width=12, state="readonly")
        self.sheet2_combo.grid(row=1, column=4, padx=5, pady=5)
        self.load_cols2_btn = ttk.Button(self.file_frame, command=lambda: self.load_column(False), width=10)
        self.load_cols2_btn.grid(row=1, column=5, padx=5, pady=5)

        # 輸出文件
        self.output_label = ttk.Label(self.file_frame, font=("Arial", 10))
        self.output_label.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(self.file_frame, textvariable=self.output_path, width=45).grid(row=2, column=1, padx=5, pady=5)
        self.save_path_btn = ttk.Button(self.file_frame, command=self.select_output, width=8)
        self.save_path_btn.grid(row=2, column=2, padx=5, pady=5)

        # ========== 2. 匹配列配置（支持多列） ==========
        self.match_frame = ttk.LabelFrame(self.root, padding=15)
        self.match_frame.pack(fill="x", padx=10, pady=8)

        # 表1匹配列（支持多選）
        self.match1_label = ttk.Label(self.match_frame, font=("Arial", 10))
        self.match1_label.grid(row=0, column=0, sticky="w", padx=5, pady=8)
        self.match1_combo = ttk.Combobox(
            self.match_frame,
            textvariable=self.match1_var,
            width=25,
            state="readonly"
        )
        self.match1_combo.grid(row=0, column=1, padx=5, pady=8)
        self.confirm1_btn = ttk.Button(self.match_frame, command=lambda: self.confirm_col("match1"), width=10)
        self.confirm1_btn.grid(row=0, column=2, padx=5, pady=8)

        # 表2匹配列（支持多選）
        self.match2_label = ttk.Label(self.match_frame, font=("Arial", 10))
        self.match2_label.grid(row=0, column=3, sticky="w", padx=5, pady=8)
        self.match2_combo = ttk.Combobox(
            self.match_frame,
            textvariable=self.match2_var,
            width=25,
            state="readonly"
        )
        self.match2_combo.grid(row=0, column=4, padx=5, pady=8)
        self.confirm2_btn = ttk.Button(self.match_frame, command=lambda: self.confirm_col("match2"), width=10)
        self.confirm2_btn.grid(row=0, column=5, padx=5, pady=8)

        # 匹配列操作按鈕
        self.add_match_btn = ttk.Button(self.match_frame, command=self.add_match_pair, width=12)
        self.add_match_btn.grid(row=0, column=6, padx=5, pady=8)
        self.remove_match_btn = ttk.Button(self.match_frame, command=self.remove_selected_match, width=12)
        self.remove_match_btn.grid(row=0, column=7, padx=5, pady=8)
        self.clear_match_btn = ttk.Button(self.match_frame, command=self.clear_all_match, width=12)
        self.clear_match_btn.grid(row=0, column=8, padx=5, pady=8)

        # 已選匹配列顯示（列表框支持多選）
        self.selected_match_label = ttk.Label(self.match_frame, font=("Arial", 10))
        self.selected_match_label.grid(row=1, column=0, sticky="w", padx=5, pady=8, columnspan=2)
        self.match_listbox = tk.Listbox(
            self.match_frame,
            width=80,
            height=4,
            selectmode=tk.EXTENDED,
            font=("Arial", 9)
        )
        self.match_listbox.grid(row=1, column=2, padx=5, pady=8, columnspan=8)
        # 匹配列滾動條
        match_scroll = ttk.Scrollbar(self.match_frame, orient="vertical", command=self.match_listbox.yview)
        match_scroll.grid(row=1, column=10, sticky="ns", pady=8)
        self.match_listbox.configure(yscrollcommand=match_scroll.set)

        # ========== 3. 填充列配置（支持多列） ==========
        self.fill_frame = ttk.LabelFrame(self.root, padding=15)
        self.fill_frame.pack(fill="x", padx=10, pady=8)

        # 表1填充列
        self.fill1_label = ttk.Label(self.fill_frame, font=("Arial", 10))
        self.fill1_label.grid(row=0, column=0, sticky="w", padx=5, pady=8)
        self.fill1_combo = ttk.Combobox(
            self.fill_frame,
            textvariable=self.fill1_var,
            width=25,
            state="readonly"
        )
        self.fill1_combo.grid(row=0, column=1, padx=5, pady=8)
        self.confirm_fill1_btn = ttk.Button(self.fill_frame, command=lambda: self.confirm_col("fill1"), width=10)
        self.confirm_fill1_btn.grid(row=0, column=2, padx=5, pady=8)

        # 表2填充列
        self.fill2_label = ttk.Label(self.fill_frame, font=("Arial", 10))
        self.fill2_label.grid(row=0, column=3, sticky="w", padx=5, pady=8)
        self.fill2_combo = ttk.Combobox(
            self.fill_frame,
            textvariable=self.fill2_var,
            width=25,
            state="readonly"
        )
        self.fill2_combo.grid(row=0, column=4, padx=5, pady=8)
        self.confirm_fill2_btn = ttk.Button(self.fill_frame, command=lambda: self.confirm_col("fill2"), width=10)
        self.confirm_fill2_btn.grid(row=0, column=5, padx=5, pady=8)

        # 填充列操作按鈕
        self.add_fill_btn = ttk.Button(self.fill_frame, command=self.add_fill_pair, width=12)
        self.add_fill_btn.grid(row=0, column=6, padx=5, pady=8)
        self.remove_fill_btn = ttk.Button(self.fill_frame, command=self.remove_selected_fill, width=12)
        self.remove_fill_btn.grid(row=0, column=7, padx=5, pady=8)
        self.clear_fill_btn = ttk.Button(self.fill_frame, command=self.clear_all_fill, width=12)
        self.clear_fill_btn.grid(row=0, column=8, padx=5, pady=8)

        # 已選填充列顯示（列表框支持多選）
        self.selected_fill_label = ttk.Label(self.fill_frame, font=("Arial", 10))
        self.selected_fill_label.grid(row=1, column=0, sticky="w", padx=5, pady=8, columnspan=2)
        self.fill_listbox = tk.Listbox(
            self.fill_frame,
            width=80,
            height=4,
            selectmode=tk.EXTENDED,
            font=("Arial", 9)
        )
        self.fill_listbox.grid(row=1, column=2, padx=5, pady=8, columnspan=8)
        # 填充列滾動條
        fill_scroll = ttk.Scrollbar(self.fill_frame, orient="vertical", command=self.fill_listbox.yview)
        fill_scroll.grid(row=1, column=10, sticky="ns", pady=8)
        self.fill_listbox.configure(yscrollcommand=fill_scroll.set)

        # ========== 4. 執行區域 ==========
        frame_run = ttk.Frame(self.root, padding=15)
        frame_run.pack(fill="x", padx=10, pady=10)
        
        self.preview_btn = ttk.Button(frame_run, command=self.preview_data, width=15, style="Accent.TButton")
        self.preview_btn.pack(side="left", padx=5)
        self.run_btn = ttk.Button(frame_run, command=self.run_fill, width=15, style="Success.TButton")
        self.run_btn.pack(side="left", padx=5)
        self.reset_btn = ttk.Button(frame_run, command=self.reset_all, width=15)
        self.reset_btn.pack(side="left", padx=5)

        # 樣式優化
        style = ttk.Style()
        style.configure("Accent.TButton", foreground="blue")
        style.configure("Success.TButton", foreground="green")

    def update_all_texts(self, *args):
        """更新所有界面文本（語言切換時調用）"""
        lang = self.current_lang.get()
        # 更新窗口標題
        self.root.title(self.lang_dict[lang]["title"])
        
        # 更新文件選擇區域
        self.file_frame.configure(text=self.lang_dict[lang]["file_frame"])
        self.file1_label.configure(text=self.lang_dict[lang]["file1_label"])
        self.file2_label.configure(text=self.lang_dict[lang]["file2_label"])
        self.browse1_btn.configure(text=self.lang_dict[lang]["browse_btn"])
        self.browse2_btn.configure(text=self.lang_dict[lang]["browse_btn"])
        self.sheet1_label.configure(text=self.lang_dict[lang]["sheet_label"])
        self.sheet2_label.configure(text=self.lang_dict[lang]["sheet_label"])
        self.load_cols1_btn.configure(text=self.lang_dict[lang]["load_cols_btn"])
        self.load_cols2_btn.configure(text=self.lang_dict[lang]["load_cols_btn"])
        self.output_label.configure(text=self.lang_dict[lang]["output_label"])
        self.save_path_btn.configure(text=self.lang_dict[lang]["save_path_btn"])
        
        # 更新匹配列區域
        self.match_frame.configure(text=self.lang_dict[lang]["match_frame"])
        self.match1_label.configure(text=self.lang_dict[lang]["match1_label"])
        self.match2_label.configure(text=self.lang_dict[lang]["match2_label"])
        self.confirm1_btn.configure(text=self.lang_dict[lang]["confirm_btn"])
        self.confirm2_btn.configure(text=self.lang_dict[lang]["confirm_btn"])
        self.add_match_btn.configure(text=self.lang_dict[lang]["add_match_btn"])
        self.remove_match_btn.configure(text=self.lang_dict[lang]["remove_match_btn"])
        self.clear_match_btn.configure(text=self.lang_dict[lang]["clear_match_btn"])
        self.selected_match_label.configure(text=self.lang_dict[lang]["selected_match"])
        
        # 更新填充列區域
        self.fill_frame.configure(text=self.lang_dict[lang]["fill_frame"])
        self.fill1_label.configure(text=self.lang_dict[lang]["fill1_label"])
        self.fill2_label.configure(text=self.lang_dict[lang]["fill2_label"])
        self.confirm_fill1_btn.configure(text=self.lang_dict[lang]["confirm_btn"])
        self.confirm_fill2_btn.configure(text=self.lang_dict[lang]["confirm_btn"])
        self.add_fill_btn.configure(text=self.lang_dict[lang]["add_fill_btn"])
        self.remove_fill_btn.configure(text=self.lang_dict[lang]["remove_fill_btn"])
        self.clear_fill_btn.configure(text=self.lang_dict[lang]["clear_fill_btn"])
        self.selected_fill_label.configure(text=self.lang_dict[lang]["selected_fill"])
        
        # 更新執行按鈕
        self.preview_btn.configure(text=self.lang_dict[lang]["preview_btn"])
        self.run_btn.configure(text=self.lang_dict[lang]["run_btn"])
        self.reset_btn.configure(text=self.lang_dict[lang]["reset_btn"])
        
        # 更新語言選擇標籤
        self.lang_label.configure(text=self.lang_dict[lang]["lang_select"])
        
        # 更新輸出文件名默認值
        if self.output_path.get() == "填充結果.xlsx" or self.output_path.get() == "fill_result.xlsx":
            self.output_path.set("填充結果.xlsx" if lang == "Zh" else "fill_result.xlsx")

    def load_file(self, is_file1):
        """加載Excel文件並獲取工作表名"""
        lang = self.current_lang.get()
        file_types = [("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")] if lang == "Zh" else [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        title = f"選擇{self.lang_dict[lang]['table1'] if is_file1 else self.lang_dict[lang]['table2']}文件" if lang == "Zh" else f"Select {self.lang_dict[lang]['table1'] if is_file1 else self.lang_dict[lang]['table2']} File"
        
        file_path = filedialog.askopenfilename(title=title, filetypes=file_types)
        if not file_path:
            return
        
        try:
            excel = pd.ExcelFile(file_path, engine="openpyxl" if file_path.endswith(".xlsx") else "xlrd")
            sheets = excel.sheet_names
            
            if is_file1:
                self.file1_path.set(file_path)
                self.sheet1_combo['values'] = sheets
                self.sheet1_combo.set(sheets[0] if sheets else "")
                self.sheets1 = sheets
            else:
                self.file2_path.set(file_path)
                self.sheet2_combo['values'] = sheets
                self.sheet2_combo.set(sheets[0] if sheets else "")
                self.sheets2 = sheets
            
            success_text = self.lang_dict[lang]["success_load_file"].format(
                self.lang_dict[lang]["table1"] if is_file1 else self.lang_dict[lang]["table2"],
                ", ".join(sheets)
            )
            messagebox.showinfo("Success" if lang == "en" else "成功", success_text)
        except Exception as e:
            error_text = self.lang_dict[lang]["error_load_file"].format(str(e))
            messagebox.showerror("Error" if lang == "en" else "錯誤", error_text)

    def load_column(self, is_file1):
        """加載列名到下拉框"""
        lang = self.current_lang.get()
        try:
            file_path = self.file1_path.get() if is_file1 else self.file2_path.get()
            sheet_name = self.sheet1_name.get() if is_file1 else self.sheet2_name.get()
            
            if not file_path:
                messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_file"])
                return
            if not sheet_name:
                messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_match"])
                return
            
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl" if file_path.endswith(".xlsx") else "xlrd")
            cols = list(df.columns)
            
            if is_file1:
                self.cols1 = cols
                self.match1_combo['values'] = cols
                self.fill1_combo['values'] = cols
            else:
                self.cols2 = cols
                self.match2_combo['values'] = cols
                self.fill2_combo['values'] = cols
            
            success_text = self.lang_dict[lang]["success_load_cols"].format(
                self.lang_dict[lang]["table1"] if is_file1 else self.lang_dict[lang]["table2"],
                ", ".join(cols)
            )
            messagebox.showinfo("Success" if lang == "en" else "成功", success_text)
        except Exception as e:
            error_text = self.lang_dict[lang]["error_load_cols"].format(str(e))
            messagebox.showerror("Error" if lang == "en" else "錯誤", error_text)

    def confirm_col(self, col_type):
        """確認列選擇"""
        lang = self.current_lang.get()
        if col_type == "match1":
            selected = self.match1_var.get()
            if not selected:
                messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col1"])
                return
            # 臨時存儲選中的匹配列1
            self.temp_match1 = selected
            confirm_text = self.lang_dict[lang]["confirm_match1"] + selected
            messagebox.showinfo("Confirm" if lang == "en" else "確認", confirm_text)
        
        elif col_type == "match2":
            selected = self.match2_var.get()
            if not selected:
                messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col1"])
                return
            # 臨時存儲選中的匹配列2
            self.temp_match2 = selected
            confirm_text = self.lang_dict[lang]["confirm_match2"] + selected
            messagebox.showinfo("Confirm" if lang == "en" else "確認", confirm_text)
        
        elif col_type == "fill1":
            selected = self.fill1_var.get()
            if not selected:
                messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col1"])
                return
            self.fill1_selected = selected
            confirm_text = self.lang_dict[lang]["confirm_fill1"] + selected
            messagebox.showinfo("Confirm" if lang == "en" else "確認", confirm_text)
        
        elif col_type == "fill2":
            selected = self.fill2_var.get()
            if not selected:
                messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col1"])
                return
            self.fill2_selected = selected
            confirm_text = self.lang_dict[lang]["confirm_fill2"] + selected
            messagebox.showinfo("Confirm" if lang == "en" else "確認", confirm_text)

    def add_match_pair(self):
        """添加匹配列對（支持多列）"""
        lang = self.current_lang.get()
        try:
            # 檢查是否已選擇匹配列對
            col1 = self.temp_match1
            col2 = self.temp_match2
        except AttributeError:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col1"])
            return
        
        if not col1 or not col2:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col1"])
            return
        
        # 檢查是否已存在該匹配對
        new_pair = (col1, col2)
        if new_pair in self.match_pairs:
            messagebox.showinfo("Info" if lang == "en" else "提示", f"該匹配列對已存在：{col1} → {col2}")
            return
        
        # 添加到匹配列列表
        self.match_pairs.append(new_pair)
        # 更新列表框顯示
        self.match_listbox.insert(tk.END, f"{col1} → {col2}")
        
        success_text = self.lang_dict[lang]["add_match_success"].format(col1, col2)
        messagebox.showinfo("Success" if lang == "en" else "成功", success_text)

    def remove_selected_match(self):
        """移除選中的匹配列對"""
        lang = self.current_lang.get()
        selected_indices = self.match_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_selected_match"])
            return
        
        # 倒序刪除避免索引錯亂
        count = 0
        for idx in sorted(selected_indices, reverse=True):
            # 從列表和列表框中刪除
            del self.match_pairs[idx]
            self.match_listbox.delete(idx)
            count += 1
        
        success_text = self.lang_dict[lang]["remove_match_success"].format(count)
        messagebox.showinfo("Success" if lang == "en" else "成功", success_text)

    def clear_all_match(self):
        """清空所有匹配列對"""
        lang = self.current_lang.get()
        self.match_pairs.clear()
        self.match_listbox.delete(0, tk.END)
        messagebox.showinfo("Success" if lang == "en" else "成功", self.lang_dict[lang]["clear_match_success"])

    def add_fill_pair(self):
        """添加填充列對"""
        lang = self.current_lang.get()
        if not self.fill1_selected:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col2"])
            return
        if not self.fill2_selected:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_col3"])
            return
        
        new_pair = (self.fill1_selected, self.fill2_selected)
        if new_pair in self.fill_pairs:
            messagebox.showinfo("Info" if lang == "en" else "提示", f"該填充列對已存在：{self.fill1_selected} → {self.fill2_selected}")
            return
        
        # 添加到填充列列表
        self.fill_pairs.append(new_pair)
        # 更新列表框顯示
        self.fill_listbox.insert(tk.END, f"{self.fill1_selected} → {self.fill2_selected}")
        
        success_text = self.lang_dict[lang]["add_fill_success"].format(self.fill1_selected, self.fill2_selected)
        messagebox.showinfo("Success" if lang == "en" else "成功", success_text)

    def remove_selected_fill(self):
        """移除選中的填充列對"""
        lang = self.current_lang.get()
        selected_indices = self.fill_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_selected_fill"])
            return
        
        # 倒序刪除避免索引錯亂
        count = 0
        for idx in sorted(selected_indices, reverse=True):
            del self.fill_pairs[idx]
            self.fill_listbox.delete(idx)
            count += 1
        
        success_text = self.lang_dict[lang]["remove_fill_success"].format(count)
        messagebox.showinfo("Success" if lang == "en" else "成功", success_text)

    def clear_all_fill(self):
        """清空所有填充列對"""
        lang = self.current_lang.get()
        self.fill_pairs.clear()
        self.fill_listbox.delete(0, tk.END)
        messagebox.showinfo("Success" if lang == "en" else "成功", self.lang_dict[lang]["clear_fill_success"])

    def select_output(self):
        """選擇輸出文件位置"""
        lang = self.current_lang.get()
        file_path = filedialog.asksaveasfilename(
            title=self.lang_dict[lang]["select_output_title"],
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")] if lang == "Zh" else [("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)

    def preview_data(self):
        """預覽數據（支持多列匹配）"""
        lang = self.current_lang.get()
        if not self.match_pairs:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_match"])
            return
        if not self.fill_pairs:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_fill"])
            return
        
        try:
            df1 = pd.read_excel(self.file1_path.get(), sheet_name=self.sheet1_name.get())
            df2 = pd.read_excel(self.file2_path.get(), sheet_name=self.sheet2_name.get())
            
            # 構建多列匹配映射
            match_map = {k: v for k, v in self.match_pairs}
            fill_map = {k: v for k, v in self.fill_pairs}
            
            # 重命名匹配列
            df1_rename = df1.rename(columns=match_map)
            # 保留匹配列和填充列
            keep_cols = list(match_map.values()) + list(fill_map.keys())
            df1_filter = df1_rename[keep_cols].drop_duplicates()
            # 重命名填充列
            df1_filter = df1_filter.rename(columns=fill_map)
            
            # 多列合併（on參數支持列表）
            preview_df = pd.merge(
                df2.head(10), 
                df1_filter, 
                on=list(match_map.values()), 
                how='left'
            )
            
            win = tk.Toplevel(self.root)
            win.title(self.lang_dict[lang]["preview_title"])
            win.geometry("850x450")
            
            text = tk.Text(win, wrap=tk.NONE, font=("Consolas", 9))
            text.insert(tk.END, preview_df.to_string(index=False))
            text.pack(fill="both", expand=True, padx=5, pady=5)
            
            x_scroll = ttk.Scrollbar(win, orient="horizontal", command=text.xview)
            x_scroll.pack(fill="x", side="bottom")
            text.configure(xscrollcommand=x_scroll.set)
            
            y_scroll = ttk.Scrollbar(win, orient="vertical", command=text.yview)
            y_scroll.pack(fill="y", side="right")
            text.configure(yscrollcommand=y_scroll.set)
            
        except Exception as e:
            error_text = self.lang_dict[lang]["error_preview"].format(str(e))
            messagebox.showerror("Error" if lang == "en" else "錯誤", error_text)

    def run_fill(self):
        """執行填充（支持多列匹配）"""
        lang = self.current_lang.get()
        if not self.file1_path.get() or not self.file2_path.get():
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_file"])
            return
        if not self.match_pairs:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_match"])
            return
        if not self.fill_pairs:
            messagebox.showwarning("Warning" if lang == "en" else "提示", self.lang_dict[lang]["warn_no_fill"])
            return
        
        try:
            df1 = pd.read_excel(self.file1_path.get(), sheet_name=self.sheet1_name.get())
            df2 = pd.read_excel(self.file2_path.get(), sheet_name=self.sheet2_name.get())
            result = df2.copy()
            
            # 構建多列匹配和填充映射
            match_map = {k: v for k, v in self.match_pairs}
            fill_map = {k: v for k, v in self.fill_pairs}
            
            # 數據處理
            df1_rename = df1.rename(columns=match_map)
            keep_cols = list(match_map.values()) + list(fill_map.keys())
            df1_filter = df1_rename[keep_cols].drop_duplicates()
            df1_filter = df1_filter.rename(columns=fill_map)
            
            # 多列合併
            result = pd.merge(
                result, 
                df1_filter, 
                on=list(match_map.values()), 
                how='left'
            )
            
            # 處理重複列
            for target_col in fill_map.values():
                if target_col in df2.columns:
                    result[target_col] = result[target_col + '_x'].fillna(result[target_col + '_y'])
                    result = result.drop(columns=[target_col + '_x', target_col + '_y'])
            
            # 保存結果
            result.to_excel(self.output_path.get(), index=False, engine="openpyxl")
            
            # 格式化顯示結果
            match_cols_text = ", ".join([f"{k}→{v}" for k, v in self.match_pairs])
            fill_cols_text = ", ".join([f"{k}→{v}" for k, v in self.fill_pairs])
            success_text = self.lang_dict[lang]["fill_result"].format(
                self.output_path.get(),
                len(df2),
                len(result),
                match_cols_text,
                fill_cols_text
            )
            messagebox.showinfo("Success" if lang == "en" else "成功", f"{self.lang_dict[lang]['success_fill']}\n{success_text}")
        except Exception as e:
            error_text = self.lang_dict[lang]["error_fill"].format(str(e))
            messagebox.showerror("Error" if lang == "en" else "錯誤", error_text)

    def reset_all(self):
        """重置所有配置"""
        lang = self.current_lang.get()
        # 清空文件和工作表
        self.file1_path.set("")
        self.file2_path.set("")
        self.output_path.set("填充結果.xlsx" if lang == "Zh" else "fill_result.xlsx")
        self.sheet1_name.set("")
        self.sheet2_name.set("")
        
        # 清空下拉框
        self.sheet1_combo['values'] = []
        self.sheet2_combo['values'] = []
        self.match1_combo['values'] = []
        self.match2_combo['values'] = []
        self.fill1_combo['values'] = []
        self.fill2_combo['values'] = []
        
        # 清空選擇變量
        self.match_pairs.clear()
        self.fill_pairs.clear()
        self.match1_var.set("")
        self.match2_var.set("")
        self.fill1_var.set("")
        self.fill2_var.set("")
        self.fill1_selected = ""
        self.fill2_selected = ""
        
        # 清空列表框
        self.match_listbox.delete(0, tk.END)
        self.fill_listbox.delete(0, tk.END)
        
        # 清空臨時變量
        if hasattr(self, 'temp_match1'):
            del self.temp_match1
        if hasattr(self, 'temp_match2'):
            del self.temp_match2
        
        messagebox.showinfo("Success" if lang == "en" else "成功", self.lang_dict[lang]["reset_success"])

if __name__ == "__main__":
    root = tk.Tk()
    root.option_add('*Font', 'Arial 10')
    app = ExcelFillerGUI(root)
    root.mainloop()
