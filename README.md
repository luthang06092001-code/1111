#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🎓 Trình Ôn Thi - Quiz Master
Phần mềm ôn thi thông minh với giao diện hiện đại
"""

import sys
import os
import subprocess

# ===== CÀI ĐẶT THƯ VIỆN TỰ ĐỘNG =====
def install_requirements():
    """Cài đặt tất cả thư viện cần thiết"""
    required_packages = {
        'pandas': 'pandas',
        'openpyxl': 'openpyxl',
        'docx': 'python-docx'
    }
    
    print("📦 Kiểm tra và cài đặt thư viện...")
    
    for module_name, package_name in required_packages.items():
        try:
            __import__(module_name)
            print(f"✅ {package_name} đã cài đặt")
        except ImportError:
            print(f"📥 Đang cài đặt {package_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", package_name])
            print(f"✅ {package_name} cài đặt thành công")

# Cài đặt thư viện
try:
    install_requirements()
except Exception as e:
    print(f"⚠️ Lỗi cài đặt: {e}")
    print("❌ Vui lòng cài đặt Python với quyền admin")
    input("Nhấn Enter để thoát...")
    sys.exit(1)

# ===== IMPORT CÁC THƯ VIỆN =====
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from docx import Document
import random
from datetime import datetime
import json
import shutil
import copy

print("✅ Tất cả thư viện sẵn sàng")
print("🚀 Khởi động giao diện...")

# ===== MAIN APPLICATION =====
class QuizApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎓 Trình Ôn Thi - Quiz Master 🎓")
        self.root.geometry("1200x900")
        self.root.config(bg="#f5f5f5")
        
        self.question_library_dir = "Question_Library"
        self.exam_dir = "Exam_Sets"
        
        if not os.path.exists(self.question_library_dir):
            os.makedirs(self.question_library_dir)
        if not os.path.exists(self.exam_dir):
            os.makedirs(self.exam_dir)
        
        self.questions = []
        self.shuffled_questions = []
        self.current_idx = 0
        self.map_btns = []
        self.after_id = None
        self.delay_time = 1000
        self.answered = set()
        self.correct_answers = {}
        self.start_time = None
        self.elapsed_time = 0
        self.quiz_in_progress = False
        self.shuffle_mode = "none"
        self.show_library = False
        self.show_questions = False
        self.loaded_exams = {}
        
        # ===== REVIEW MODE =====
        self.is_review_mode = False
        self.wrong_questions = []
        self.review_idx = 0  # ✅ Đặt tên rõ ràng hơn

        self.setup_ui()

    def setup_ui(self):
        # ===== HEADER =====
        header = tk.Frame(self.root, bg="#2196f3", height=60)
        header.pack(fill="x", padx=0, pady=0)
        header.pack_propagate(False)
        
        title_label = tk.Label(header, text="🎓 TRÌNH ÔN THI THÔNG MINH 🎓", 
                              font=("Arial", 20, "bold"), bg="#2196f3", fg="white")
        title_label.pack(pady=10)
        
        # ===== MAIN CONTAINER =====
        self.main_container = tk.Frame(self.root, bg="#f5f5f5")
        self.main_container.pack(fill="both", expand=True)
        
        # ===== SIDE PANEL =====
        self.side_panel = tk.Frame(self.main_container, bg="#ffffff", width=0)
        self.side_panel.pack(side="left", fill="both", padx=0, pady=0)
        self.side_panel.pack_propagate(False)
        
        # ===== LIBRARY FRAME =====
        self.library_frame = tk.Frame(self.side_panel, bg="#ffffff", width=180)
        self.library_frame.pack(fill="both", expand=True, padx=3, pady=3)
        self.library_frame.pack_propagate(False)
        
        library_title = tk.Label(self.library_frame, text="📚 THƯ VIỆN", 
                                font=("Arial", 9, "bold"), bg="#ffffff", fg="#2196f3")
        library_title.pack(pady=5, padx=3)
        
        btn_add_file = tk.Button(self.library_frame, text="➕ Thêm", 
                                command=self.add_file_to_library,
                                font=("Arial", 8, "bold"), bg="#ff6f00", fg="white",
                                padx=6, pady=4, relief="raised", bd=1, width=14)
        btn_add_file.pack(pady=2, padx=3)
        
        btn_refresh = tk.Button(self.library_frame, text="🔄 Làm mới", 
                               command=self.refresh_library,
                               font=("Arial", 8, "bold"), bg="#4caf50", fg="white",
                               padx=6, pady=4, relief="raised", bd=1, width=14)
        btn_refresh.pack(pady=1, padx=3)
        
        scrollbar = ttk.Scrollbar(self.library_frame)
        scrollbar.pack(side="right", fill="y", padx=1, pady=3)
        
        self.library_listbox = tk.Listbox(self.library_frame, bg="#fafafa", fg="#333333",
                                         font=("Arial", 7), yscrollcommand=scrollbar.set,
                                         relief="solid", bd=1, highlightthickness=0)
        self.library_listbox.pack(side="left", fill="both", expand=True, padx=2, pady=3)
        scrollbar.config(command=self.library_listbox.yview)
        self.library_listbox.bind('<<ListboxSelect>>', self.on_file_selected)
        
        # ===== QUESTIONS FRAME =====
        self.questions_frame_container = tk.Frame(self.side_panel, bg="#ffffff", width=180)
        self.questions_frame_container.pack(fill="both", expand=True, padx=3, pady=3)
        self.questions_frame_container.pack_propagate(False)
        
        q_title = tk.Label(self.questions_frame_container, text="📋 CÂUHỎI", 
                          font=("Arial", 9, "bold"), bg="#ffffff", fg="#2196f3")
        q_title.pack(pady=5, padx=3)
        
        scrollbar_q = ttk.Scrollbar(self.questions_frame_container)
        scrollbar_q.pack(side="right", fill="y", padx=1, pady=3)
        
        self.question_frame = tk.Canvas(self.questions_frame_container, bg="#ffffff", 
                                        yscrollcommand=scrollbar_q.set, highlightthickness=0)
        self.question_frame.pack(fill="both", expand=True, padx=2, pady=3)
        scrollbar_q.config(command=self.question_frame.yview)
        
        self.inner_frame = tk.Frame(self.question_frame, bg="#ffffff")
        self.question_frame.create_window((0, 0), window=self.inner_frame, anchor="nw")
        
        def on_configure(event):
            self.question_frame.configure(scrollregion=self.question_frame.bbox("all"))
        self.inner_frame.bind("<Configure>", on_configure)
        
        # ===== RIGHT PANEL =====
        right_panel = tk.Frame(self.main_container, bg="#f5f5f5")
        right_panel.pack(side="right", fill="both", expand=True, padx=0, pady=0)
        
        # ===== TOOLBAR =====
        toolbar = tk.Frame(right_panel, bg="#ffffff", pady=6)
        toolbar.pack(fill="x", padx=5, pady=5)
        
        self.btn_toggle_lib = tk.Button(toolbar, text="📚 Thư viện", command=self.toggle_library, 
                 font=("Arial", 8, "bold"), bg="#2196f3", fg="white", padx=8, pady=6, relief="raised", bd=1)
        self.btn_toggle_lib.pack(side="left", padx=2)
        
        self.btn_toggle_q = tk.Button(toolbar, text="📋 Câu hỏi", command=self.toggle_questions, 
                 font=("Arial", 8, "bold"), bg="#2196f3", fg="white", padx=8, pady=6, relief="raised", bd=1)
        self.btn_toggle_q.pack(side="left", padx=2)
        
        tk.Button(toolbar, text="📁 Nạp Word", command=self.load_word_from_system, 
                 font=("Arial", 8, "bold"), bg="#ff6f00", fg="white", padx=8, pady=6, relief="raised", bd=1).pack(side="left", padx=2)
        tk.Button(toolbar, text="📊 Nạp Excel", command=self.load_excel_from_system, 
                 font=("Arial", 8, "bold"), bg="#ff6f00", fg="white", padx=8, pady=6, relief="raised", bd=1).pack(side="left", padx=2)
        
        tk.Button(toolbar, text="🎲 Tạo Đề", command=self.create_exam_window, 
                 font=("Arial", 8, "bold"), bg="#9c27b0", fg="white", padx=8, pady=6, relief="raised", bd=1).pack(side="left", padx=2)
        
        tk.Button(toolbar, text="🔄 Làm lại", command=self.reset_quiz, 
                 font=("Arial", 8, "bold"), bg="#4caf50", fg="white", padx=8, pady=6, relief="raised", bd=1).pack(side="left", padx=2)
        
        tk.Label(toolbar, text="", bg="#ffffff").pack(side="left", expand=True)
        
        self.lbl_current_file = tk.Label(toolbar, text="📄 Chưa chọn", 
                                        font=("Arial", 8, "bold"), bg="#ffffff", fg="#f57c00")
        self.lbl_current_file.pack(side="right", padx=10)
        
        # ===== CONTROL FRAME =====
        control_frame = tk.Frame(right_panel, bg="#ffffff", pady=6)
        control_frame.pack(fill="x", padx=5, pady=0)
        
        time_frame = tk.Frame(control_frame, bg="#ffffff")
        time_frame.pack(fill="x", padx=10, pady=2)
        
        tk.Label(time_frame, text="⏱️ Thời gian:", bg="#ffffff", fg="#333333", font=("Arial", 8, "bold")).pack(side="left", padx=3)
        
        self.time_slider = ttk.Scale(time_frame, from_=1, to=10, orient="horizontal", 
                                      command=self.update_delay_time, length=100)
        self.time_slider.set(1)
        self.time_slider.pack(side="left", padx=3)
        
        self.lbl_time = tk.Label(time_frame, text="1s", bg="#ffffff", fg="#333333", font=("Arial", 8, "bold"), width=3)
        self.lbl_time.pack(side="left", padx=3)
        
        shuffle_frame = tk.Frame(control_frame, bg="#ffffff")
        shuffle_frame.pack(fill="x", padx=10, pady=2)
        
        tk.Label(shuffle_frame, text="🎲 Xáo:", bg="#ffffff", fg="#333333", font=("Arial", 8, "bold")).pack(side="left", padx=3)
        
        self.shuffle_var = tk.StringVar(value="none")
        self.shuffle_menu = ttk.Combobox(shuffle_frame, textvariable=self.shuffle_var,
                                        values=["Không đảo", "Đảo câu", "Đảo câu&đáp", "Đảo đáp"],
                                        state="readonly", font=("Arial", 8), width=14)
        self.shuffle_menu.pack(side="left", padx=3, fill="x", expand=True)
        self.shuffle_menu.bind("<<ComboboxSelected>>", self.on_shuffle_changed)
        
        # ===== INFO FRAME =====
        info_frame = tk.Frame(right_panel, bg="#ffffff", height=50)
        info_frame.pack(fill="x", padx=5, pady=0)
        info_frame.pack_propagate(False)
        
        top_info = tk.Frame(info_frame, bg="#ffffff")
        top_info.pack(fill="x", padx=10, pady=3)
        
        self.lbl_progress = tk.Label(top_info, text="📍 Câu 0/0", bg="#ffffff", fg="#2196f3", font=("Arial", 11, "bold"))
        self.lbl_progress.pack(side="left", padx=10)
        
        self.lbl_mode = tk.Label(top_info, text="", bg="#ffffff", fg="#d32f2f", font=("Arial", 9, "bold"))
        self.lbl_mode.pack(side="right", padx=10)
        
        bottom_info = tk.Frame(info_frame, bg="#ffffff")
        bottom_info.pack(fill="x", padx=10, pady=3)
        
        self.lbl_stats = tk.Label(bottom_info, text="✓: 0 | ✗: 0", bg="#ffffff", fg="#4caf50", font=("Arial", 9, "bold"))
        self.lbl_stats.pack(side="left", padx=10)
        
        self.lbl_clock = tk.Label(bottom_info, text="⏱️ 00:00", bg="#ffffff", fg="#f57c00", font=("Arial", 10, "bold"))
        self.lbl_clock.pack(side="right", padx=10)
        
        # ===== CONTENT CONTAINER =====
        content_container = tk.Frame(right_panel, bg="#f5f5f5")
        content_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # ===== QUESTION DISPLAY =====
        q_container = tk.Frame(content_container, bg="#ffffff", height=80)
        q_container.pack(fill="both", expand=False, padx=0, pady=0)
        q_container.pack_propagate(False)
        
        q_scrollbar = ttk.Scrollbar(q_container)
        q_scrollbar.pack(side="right", fill="y")
        
        self.q_canvas = tk.Canvas(q_container, bg="#ffffff", yscrollcommand=q_scrollbar.set, highlightthickness=0)
        self.q_canvas.pack(side="left", fill="both", expand=True)
        q_scrollbar.config(command=self.q_canvas.yview)
        
        self.q_text_frame = tk.Frame(self.q_canvas, bg="#ffffff")
        self.q_canvas.create_window((0, 0), window=self.q_text_frame, anchor="nw")
        
        def on_q_configure(event):
            self.q_canvas.configure(scrollregion=self.q_canvas.bbox("all"))
        self.q_text_frame.bind("<Configure>", on_q_configure)
        
        self.lbl_q = tk.Label(self.q_text_frame, text="📝 Chưa có dữ liệu\n\nHãy chọn file!", 
                             font=("Arial", 11, "bold"), wraplength=1000, justify="left", 
                             fg="#1565c0", bg="#ffffff")
        self.lbl_q.pack(pady=5, padx=10)
        
        # ===== ANSWER DISPLAY =====
        ans_container = tk.Frame(content_container, bg="#f5f5f5")
        ans_container.pack(fill="both", expand=True, padx=0, pady=0)
        
        ans_scrollbar = ttk.Scrollbar(ans_container)
        ans_scrollbar.pack(side="right", fill="y")
        
        self.ans_canvas = tk.Canvas(ans_container, bg="#f5f5f5", yscrollcommand=ans_scrollbar.set, highlightthickness=0)
        self.ans_canvas.pack(side="left", fill="both", expand=True)
        ans_scrollbar.config(command=self.ans_canvas.yview)
        
        self.ans_frame = tk.Frame(self.ans_canvas, bg="#f5f5f5")
        self.ans_canvas.create_window((0, 0), window=self.ans_frame, anchor="nw")
        
        def on_ans_configure(event):
            self.ans_canvas.configure(scrollregion=self.ans_canvas.bbox("all"))
        self.ans_frame.bind("<Configure>", on_ans_configure)
        
        self.ans_btns = []
        for i in range(4):
            btn = tk.Button(self.ans_frame, text="", command=lambda i=i: self.handle_answer(i), 
                            font=("Arial", 10, "bold"), state="disabled", bg="#2196f3", fg="white", 
                            relief="raised", bd=2, activebackground="#1976d2", activeforeground="white",
                            wraplength=1000, justify="left")
            btn.pack(pady=4, fill="both", expand=True, padx=5)
            self.ans_btns.append(btn)
        
        self.refresh_library()

    def create_exam_window(self):
        """Tạo cửa sổ để tạo đề thi"""
        exam_window = tk.Toplevel(self.root)
        exam_window.title("🎲 Tạo Đề Thi")
        exam_window.geometry("500x400")
        exam_window.config(bg="#ffffff")
        
        title = tk.Label(exam_window, text="🎲 TẠO ĐỀ THI TỪ NGÂN HÀNG CÂU HỎI", 
                        font=("Arial", 12, "bold"), bg="#ffffff", fg="#2196f3")
        title.pack(pady=15)
        
        frame_files = tk.Frame(exam_window, bg="#ffffff")
        frame_files.pack(fill="both", expand=True, padx=15, pady=10)
        
        tk.Label(frame_files, text="📋 Chọn các bộ đề (với số câu):", 
                font=("Arial", 10, "bold"), bg="#ffffff", fg="#333333").pack(anchor="w", pady=(0, 10))
        
        scrollbar_files = ttk.Scrollbar(frame_files)
        scrollbar_files.pack(side="right", fill="y")
        
        self.file_canvas = tk.Canvas(frame_files, bg="#fafafa", highlightthickness=1, 
                                     highlightbackground="#ddd", yscrollcommand=scrollbar_files.set, height=150)
        self.file_canvas.pack(side="left", fill="both", expand=True)
        scrollbar_files.config(command=self.file_canvas.yview)
        
        self.file_frame = tk.Frame(self.file_canvas, bg="#fafafa")
        self.file_canvas.create_window((0, 0), window=self.file_frame, anchor="nw")
        
        def on_file_configure(event):
            self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all"))
        self.file_frame.bind("<Configure>", on_file_configure)
        
        files = [f for f in os.listdir(self.question_library_dir) 
                if f.endswith(('.docx', '.xlsx', '.xls'))]
        
        self.file_entries = {}
        for file in sorted(files):
            file_frame = tk.Frame(self.file_frame, bg="#fafafa")
            file_frame.pack(fill="x", padx=5, pady=3)
            
            var = tk.BooleanVar()
            tk.Checkbutton(file_frame, text=file, variable=var, bg="#fafafa", 
                          font=("Arial", 9)).pack(side="left", padx=5)
            
            tk.Label(file_frame, text="Số câu:", bg="#fafafa", font=("Arial", 8)).pack(side="left", padx=(10, 5))
            
            spinbox = tk.Spinbox(file_frame, from_=1, to=100, width=5, font=("Arial", 9))
            spinbox.pack(side="left", padx=5)
            spinbox.delete(0, tk.END)
            spinbox.insert(0, "10")
            
            self.file_entries[file] = (var, spinbox)
        
        btn_frame = tk.Frame(exam_window, bg="#ffffff")
        btn_frame.pack(fill="x", padx=15, pady=15)
        
        def create_exam():
            selected_files = {file: (var.get(), int(spinbox.get())) 
                            for file, (var, spinbox) in self.file_entries.items() if var.get()}
            
            if not selected_files:
                messagebox.showwarning("⚠️ Cảnh báo", "Hãy chọn ít nhất 1 bộ đề!")
                return
            
            try:
                all_questions = []
                
                for file, (_, num_questions) in selected_files.items():
                    filepath = os.path.join(self.question_library_dir, file)
                    questions = self.load_questions_from_file(filepath)
                    
                    num_to_select = min(int(num_questions), len(questions))
                    selected = random.sample(questions, num_to_select)
                    all_questions.extend(selected)
                
                random.shuffle(all_questions)
                for q in all_questions:
                    random.shuffle(q['a'])
                
                self.questions = all_questions
                self.shuffled_questions = copy.deepcopy(self.questions)
                self.answered = set()
                self.correct_answers = {}
                self.start_time = None
                self.elapsed_time = 0
                self.quiz_in_progress = False
                self.is_review_mode = False
                self.wrong_questions = []
                self.review_idx = 0
                self.current_idx = 0
                self.lbl_clock.config(text="⏱️ 00:00")
                
                messagebox.showinfo("✅ Thành công", 
                                  f"✅ Tạo đề thi thành công!\n\n"
                                  f"📊 Tổng {len(all_questions)} câu")
                
                exam_window.destroy()
                self.render_question()
                self.create_question_buttons()
                self.update_clock()
                
            except Exception as e:
                messagebox.showerror("❌ Lỗi", f"Lỗi tạo đề thi:\n{e}")
        
        tk.Button(btn_frame, text="✅ Tạo Đề", command=create_exam, 
                 font=("Arial", 10, "bold"), bg="#4caf50", fg="white", 
                 padx=20, pady=8, relief="raised", bd=2).pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="❌ Đóng", command=exam_window.destroy, 
                 font=("Arial", 10, "bold"), bg="#f57c00", fg="white", 
                 padx=20, pady=8, relief="raised", bd=2).pack(side="left", padx=5)

    def load_questions_from_file(self, filepath):
        """Tải câu hỏi từ file Word/Excel"""
        try:
            raw = []
            
            if filepath.endswith('.docx'):
                doc = Document(filepath)
                for p in doc.paragraphs:
                    text = p.text.strip()
                    if len(text) > 1:
                        raw.append((text[0], text[1:].strip()))
            else:
                df = pd.read_excel(filepath, header=None)
                for row in df.values:
                    if len(row) >= 2:
                        raw.append((str(row[0]), str(row[1])))
            
            questions = []
            curr = None
            for tag, txt in raw:
                tag = str(tag).strip()
                if tag == '?':
                    if curr: questions.append(curr)
                    curr = {'q': txt, 'a': []}
                elif tag in ['!', '#'] and curr:
                    curr['a'].append({'t': txt, 'is': (tag == '!')})
            if curr: questions.append(curr)
            
            return questions
        except Exception as e:
            messagebox.showerror("❌ Lỗi", f"Lỗi nạp file:\n{e}")
            return []

    def toggle_library(self):
        self.show_library = not self.show_library
        if self.show_library:
            self.side_panel.config(width=200)
            self.library_frame.pack(fill="both", expand=True, padx=3, pady=3)
            self.btn_toggle_lib.config(relief="sunken", bg="#1976d2")
        else:
            self.library_frame.pack_forget()
            self.side_panel.config(width=0)
            self.btn_toggle_lib.config(relief="raised", bg="#2196f3")

    def toggle_questions(self):
        self.show_questions = not self.show_questions
        if self.show_questions:
            self.side_panel.config(width=200)
            self.questions_frame_container.pack(fill="both", expand=True, padx=3, pady=3)
            self.btn_toggle_q.config(relief="sunken", bg="#1976d2")
        else:
            self.questions_frame_container.pack_forget()
            if not self.show_library:
                self.side_panel.config(width=0)
            self.btn_toggle_q.config(relief="raised", bg="#2196f3")

    def on_shuffle_changed(self, event=None):
        shuffle_map = {"Không đảo": "none", "Đảo câu": "questions", "Đảo câu&đáp": "all", "Đảo đáp": "answers"}
        self.shuffle_mode = shuffle_map.get(self.shuffle_var.get(), "none")
        
        if self.questions:
            self.apply_shuffle()
            self.current_idx = 0
            self.answered = set()
            self.correct_answers = {}
            self.is_review_mode = False
            self.wrong_questions = []
            self.review_idx = 0
            self.render_question()
            self.create_question_buttons()

    def apply_shuffle(self):
        self.shuffled_questions = copy.deepcopy(self.questions)
        
        if self.shuffle_mode == "questions" or self.shuffle_mode == "all":
            random.shuffle(self.shuffled_questions)
        
        if self.shuffle_mode == "all" or self.shuffle_mode == "answers":
            for q in self.shuffled_questions:
                random.shuffle(q['a'])

    def refresh_library(self):
        self.library_listbox.delete(0, tk.END)
        if not os.path.exists(self.question_library_dir):
            os.makedirs(self.question_library_dir)
        
        files = [f for f in os.listdir(self.question_library_dir) if f.endswith(('.docx', '.xlsx', '.xls'))]
        
        if not files:
            self.library_listbox.insert(tk.END, "Trống")
        else:
            for file in sorted(files):
                self.library_listbox.insert(tk.END, file)

    def on_file_selected(self, event):
        selection = self.library_listbox.curselection()
        if not selection:
            return
        
        filename = self.library_listbox.get(selection[0])
        if filename == "Trống":
            return
        
        filepath = os.path.join(self.question_library_dir, filename)
        self.current_file = filename
        self.lbl_current_file.config(text=f"📄 {filename[:20]}...")
        
        if filename.endswith('.docx'):
            self.load_word_from_path(filepath)
        elif filename.endswith(('.xlsx', '.xls')):
            self.load_excel_from_path(filepath)

    def add_file_to_library(self):
        filetypes = [("Word files", "*.docx"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        file = filedialog.askopenfilename(filetypes=filetypes)
        
        if file:
            try:
                filename = os.path.basename(file)
                dest_path = os.path.join(self.question_library_dir, filename)
                shutil.copy2(file, dest_path)
                messagebox.showinfo("✅ Thành công", f"Đã thêm '{filename}'!")
                self.refresh_library()
            except Exception as e:
                messagebox.showerror("❌ Lỗi", f"Không thể thêm tập tin:\n{e}")

    def load_word_from_system(self):
        file = filedialog.askopenfilename(filetypes=[("Word files", "*.docx"), ("All files", "*.*")])
        if file:
            self.current_file = os.path.basename(file)
            self.lbl_current_file.config(text=f"📄 {self.current_file[:20]}...")
            self.load_word_from_path(file)

    def load_excel_from_system(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if file:
            self.current_file = os.path.basename(file)
            self.lbl_current_file.config(text=f"📄 {self.current_file[:20]}...")
            self.load_excel_from_path(file)

    def load_word_from_path(self, filepath):
        try:
            doc = Document(filepath)
            raw = []
            for p in doc.paragraphs:
                text = p.text.strip()
                if len(text) > 1:
                    raw.append((text[0], text[1:].strip()))
            self.parse_raw(raw)
            messagebox.showinfo("✅ OK", f"Nạp {len(self.questions)} câu!")
        except Exception as e:
            messagebox.showerror("❌ Lỗi", f"Lỗi nạp Word:\n{e}")

    def load_excel_from_path(self, filepath):
        try:
            df = pd.read_excel(filepath, header=None)
            raw = []
            for row in df.values:
                if len(row) >= 2:
                    raw.append((str(row[0]), str(row[1])))
            self.parse_raw(raw)
            messagebox.showinfo("✅ OK", f"Nạp {len(self.questions)} câu!")
        except Exception as e:
            messagebox.showerror("❌ Lỗi", f"Lỗi nạp Excel:\n{e}")

    def parse_raw(self, data):
        self.questions = []
        self.shuffled_questions = []
        self.answered = set()
        self.correct_answers = {}
        self.start_time = None
        self.elapsed_time = 0
        self.quiz_in_progress = False
        self.is_review_mode = False
        self.wrong_questions = []
        self.review_idx = 0
        self.lbl_clock.config(text="⏱️ 00:00")
        
        curr = None
        for row in data:
            tag, txt = str(row[0]).strip(), str(row[1]).strip()
            if tag == '?':
                if curr: self.questions.append(curr)
                curr = {'q': txt, 'a': []}
            elif tag in ['!', '#'] and curr:
                curr['a'].append({'t': txt, 'is': (tag == '!')})
        if curr: self.questions.append(curr)
        
        if len(self.questions) == 0:
            messagebox.showerror("❌ Lỗi", "Không có câu hỏi!")
            return
        
        self.apply_shuffle()
        
        self.current_idx = 0
        self.render_question()
        self.create_question_buttons()
        self.update_clock()

    def create_question_buttons(self):
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        
        self.map_btns = []
        question_list = self.shuffled_questions if self.shuffled_questions else self.questions
        
        row_frame = None
        col_count = 0
        
        for i in range(len(question_list)):
            if col_count == 0:
                row_frame = tk.Frame(self.inner_frame, bg="#ffffff")
                row_frame.pack(fill="x", padx=2, pady=1)
            
            btn = tk.Button(row_frame, text=f"{i+1}", width=4, 
                            command=lambda idx=i: self.jump_to_question(idx),
                            font=("Arial", 8, "bold"), relief="raised", bd=1,
                            bg="#e3f2fd", fg="#1976d2")
            btn.pack(side="left", padx=1, pady=1, expand=True, fill="both")
            self.map_btns.append(btn)
            
            col_count += 1
            if col_count == 3:
                col_count = 0
        
        self.update_question_buttons_style()

    def update_question_buttons_style(self):
        for i, btn in enumerate(self.map_btns):
            if self.is_review_mode:
                # Mode ôn lại
                if i == self.review_idx:
                    btn.config(bg="#1976d2", fg="white", relief="sunken")
                else:
                    btn.config(bg="#e3f2fd", fg="#1976d2", relief="solid")
            else:
                # Mode thường
                if i == self.current_idx:
                    btn.config(bg="#1976d2", fg="white", relief="sunken")
                elif i in self.answered:
                    if self.correct_answers.get(i, False):
                        btn.config(bg="#388e3c", fg="white", relief="solid")
                    else:
                        btn.config(bg="#d32f2f", fg="white", relief="solid")
                else:
                    btn.config(bg="#e3f2fd", fg="#1976d2", relief="solid")

    def jump_to_question(self, idx):
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        
        if self.is_review_mode:
            self.review_idx = idx
        else:
            self.current_idx = idx
        
        self.render_question()

    def render_question(self):
        """✅ FIX: Lấy câu hỏi đúng theo mode"""
        if not self.questions:
            return
        
        question_list = self.shuffled_questions if self.shuffled_questions else self.questions
        
        if self.is_review_mode:
            # ✅ Mode ôn lại: Lấy từ danh sách sai
            if self.review_idx >= len(self.wrong_questions):
                return
            
            actual_question_idx = self.wrong_questions[self.review_idx]
            q = question_list[actual_question_idx]
            
            self.lbl_progress.config(text=f"📍 Ôn {self.review_idx + 1}/{len(self.wrong_questions)}")
            self.lbl_mode.config(text="🔧 ÔN LẠI")
        else:
            # ✅ Mode thường: Lấy từ current_idx
            if self.current_idx >= len(question_list):
                return
            
            q = question_list[self.current_idx]
            self.lbl_progress.config(text=f"📍 Câu {self.current_idx + 1}/{len(question_list)}")
            self.lbl_mode.config(text="")
        
        # ✅ Cập nhật nội dung câu hỏi
        self.lbl_q.config(text=q['q'])
        
        # ✅ Cập nhật nội dung đáp án
        for i, btn in enumerate(self.ans_btns):
            if i < len(q['a']):
                answer_text = q['a'][i]['t']
                btn.config(text=answer_text, state="normal", bg="#2196f3", fg="white")
            else:
                btn.config(text="", state="disabled", bg="#f5f5f5")
        
        self.update_question_buttons_style()
        self.update_stats_display()

    def update_stats_display(self):
        if self.is_review_mode:
            self.lbl_stats.config(text=f"✓ {self.review_idx + 1}/{len(self.wrong_questions)}")
        else:
            correct = sum(1 for idx in self.answered if self.correct_answers.get(idx, False))
            wrong = len(self.answered) - correct
            self.lbl_stats.config(text=f"✓: {correct} | ✗: {wrong}")

    def update_delay_time(self, value):
        seconds = int(float(value))
        self.delay_time = seconds * 1000
        self.lbl_time.config(text=f"{seconds}s")

    def update_clock(self):
        if self.quiz_in_progress and self.start_time:
            elapsed = int((datetime.now() - self.start_time).total_seconds())
            minutes = elapsed // 60
            seconds = elapsed % 60
            self.lbl_clock.config(text=f"⏱️ {minutes:02d}:{seconds:02d}")
            self.elapsed_time = elapsed
            self.timer_id = self.root.after(1000, self.update_clock)
        elif not self.quiz_in_progress:
            self.timer_id = self.root.after(1000, self.update_clock)

    def handle_answer(self, idx):
        """✅ Xử lý trả lời"""
        if not self.questions:
            return
        
        question_list = self.shuffled_questions if self.shuffled_questions else self.questions
        
        if self.is_review_mode:
            # ✅ Mode ôn lại
            actual_question_idx = self.wrong_questions[self.review_idx]
            q = question_list[actual_question_idx]
            correct = next(i for i, a in enumerate(q['a']) if a['is'])
            
            for btn in self.ans_btns:
                btn.config(state="disabled")
            
            if idx == correct:
                self.ans_btns[idx].config(bg="#4caf50")
            else:
                self.ans_btns[idx].config(bg="#e74c3c")
                self.ans_btns[correct].config(bg="#4caf50")
            
            self.update_stats_display()
            self.after_id = self.root.after(self.delay_time, self.next_review_q)
        else:
            # ✅ Mode thường
            if self.current_idx in self.answered:
                return
            
            if self.start_time is None:
                self.start_time = datetime.now()
                self.quiz_in_progress = True
                self.update_clock()
            
            q = question_list[self.current_idx]
            correct = next(i for i, a in enumerate(q['a']) if a['is'])
            
            is_correct = (idx == correct)
            self.answered.add(self.current_idx)
            self.correct_answers[self.current_idx] = is_correct
            
            for btn in self.ans_btns:
                btn.config(state="disabled")
            
            if is_correct:
                self.ans_btns[idx].config(bg="#4caf50")
            else:
                self.ans_btns[idx].config(bg="#e74c3c")
                self.ans_btns[correct].config(bg="#4caf50")
            
            self.update_stats_display()
            
            if len(self.answered) == len(question_list):
                self.quiz_in_progress = False
                self.root.after(1500, self.show_final_result)
            else:
                self.after_id = self.root.after(self.delay_time, self.next_q)

    def next_q(self):
        self.after_id = None
        question_list = self.shuffled_questions if self.shuffled_questions else self.questions
        
        if self.current_idx < len(question_list) - 1:
            self.current_idx += 1
            self.render_question()
        else:
            self.show_final_result()

    def next_review_q(self):
        """✅ Câu tiếp theo trong mode ôn lại"""
        self.after_id = None
        self.review_idx += 1
        
        if self.review_idx < len(self.wrong_questions):
            self.render_question()
        else:
            self.show_review_complete()

    def show_final_result(self):
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        
        question_list = self.shuffled_questions if self.shuffled_questions else self.questions
        total_questions = len(question_list)
        correct_count = sum(1 for idx in self.answered if self.correct_answers.get(idx, False))
        wrong_count = len(self.answered) - correct_count
        score_percent = (correct_count / total_questions * 100) if total_questions > 0 else 0
        
        self.wrong_questions = [idx for idx in range(len(question_list)) if not self.correct_answers.get(idx, False)]
        
        minutes = self.elapsed_time // 60
        seconds = self.elapsed_time % 60
        time_str = f"{minutes:02d}:{seconds:02d}"
        
        if score_percent >= 90:
            rating = "🌟 XUẤT SẮC"
        elif score_percent >= 80:
            rating = "⭐ TỐT"
        elif score_percent >= 70:
            rating = "👍 KHÁC"
        elif score_percent >= 60:
            rating = "😐 CẦN GẮNG"
        else:
            rating = "❌ CHƯA ĐỦ"
        
        result_text = f"""
╔════════════════════════════════════╗
║    🎯 KẾT QUẢ CUỐI CÙNG 🎯       ║
╠════════════════════════════════════╣
║  {rating:<32}║
║                                    ║
║  Tổng: {total_questions:<26}║
║  ✓ Đúng: {correct_count:<24}║
║  ✗ Sai: {wrong_count:<25}║
║  Điểm: {score_percent:>6.1f}%{' '*23}║
║  ⏱️ {time_str:<31}║
╚════════════════════════════════════╝
        """
        
        result_window = tk.Toplevel(self.root)
        result_window.title("🎉 HOÀN THÀNH")
        result_window.geometry("450x400")
        result_window.resizable(False, False)
        result_window.config(bg="#ffffff")
        
        result_label = tk.Label(result_window, text=result_text, font=("Courier", 10, "bold"), 
                               justify="left", bg="#ffffff", fg="#1565c0")
        result_label.pack(pady=20, padx=20)
        
        btn_frame = tk.Frame(result_window, bg="#ffffff")
        btn_frame.pack(pady=20)
        
        def review_wrong():
            result_window.destroy()
            self.start_review_mode()
        
        def close_window():
            result_window.destroy()
        
        if wrong_count > 0:
            tk.Button(btn_frame, text=f"🔧 Ôn {wrong_count} câu sai", command=review_wrong, 
                     font=("Arial", 9, "bold"), bg="#d32f2f", fg="white", 
                     padx=15, pady=8, relief="raised", bd=2).pack(side="left", padx=8)
        
        tk.Button(btn_frame, text="✓ Đóng", command=close_window, 
                 font=("Arial", 9, "bold"), bg="#2196f3", fg="white", 
                 padx=15, pady=8, relief="raised", bd=2).pack(side="left", padx=8)
        
        self.lbl_q.config(text=f"🎉 HOÀN THÀNH 🎉\n\n{rating} - {score_percent:.1f}%", 
                         fg="#388e3c")

    def start_review_mode(self):
        """✅ Bắt đầu mode ôn lại"""
        self.is_review_mode = True
        self.review_idx = 0
        
        self.render_question()
        self.create_question_buttons()

    def show_review_complete(self):
        """✅ Hoàn thành ôn lại"""
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        
        messagebox.showinfo("✅ OK", f"Ôn xong {len(self.wrong_questions)} câu!")
        self.lbl_q.config(text="✅ Hoàn thành!", fg="#388e3c")
        self.lbl_mode.config(text="")

    def reset_quiz(self):
        if self.questions:
            self.answered = set()
            self.correct_answers = {}
            self.start_time = None
            self.elapsed_time = 0
            self.current_idx = 0
            self.quiz_in_progress = False
            self.is_review_mode = False
            self.wrong_questions = []
            self.review_idx = 0
            self.lbl_clock.config(text="⏱️ 00:00")
            
            if self.after_id:
                self.root.after_cancel(self.after_id)
                self.after_id = None
            
            self.apply_shuffle()
            
            self.render_question()
            self.create_question_buttons()
            messagebox.showinfo("✅ OK", "Reset xong!")
        else:
            messagebox.showwarning("⚠️ Cảnh báo", "Chưa có dữ liệu!")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = QuizApp(root)
        print("✅ Giao diện sẵn sàng!")
        root.mainloop()
    except Exception as e:
        print(f"LỖI: {e}")
        import traceback
        traceback.print_exc()
        input("Nhấn Enter để thoát...")
