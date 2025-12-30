import sys
import os
import time
import itertools
import random
import re
import pandas as pd
from datetime import datetime

# PyQt6 Libraries
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QTableWidget, 
                             QTableWidgetItem, QFileDialog, QProgressBar, QGroupBox, 
                             QHeaderView, QMessageBox, QFrame, QMenuBar, QListWidget)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QSize
from PyQt6.QtGui import QAction, QFont, QPixmap, QColor, QBrush, QIcon

# Selenium Libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# ================= 1. CORE LOGIC (4-LAYER ANALYSIS) =================

class TextProcessor:
    def __init__(self):
        # Layer 1: Phonetic (Homophones)
        self.phonetic_groups = [
            ['س', 'ص', 'ث'], 
            ['ز', 'ذ', 'ض', 'ظ'], 
            ['ت', 'ط'], 
            ['ه', 'ح'], 
            ['غ', 'ق']
        ]
        
        # Layer 2: Visual Similarity (Common mistakes/confusions)
        self.visual_groups = [
            ['ک', 'گ'], 
            ['ب', 'پ'], 
            ['ج', 'چ'], 
            ['ی', 'ئ']
        ]
        
        # Combined Substitution Groups for Permutation
        self.all_subs = self.phonetic_groups + self.visual_groups

        # Layer 3: Transliteration Map (Persian to Fingilish)
        self.f_to_e_map = {
            'ا': 'a', 'آ': 'a', 'ب': 'b', 'پ': 'p', 'ت': 't', 'ث': 's',
            'ج': 'j', 'چ': 'ch', 'ح': 'h', 'خ': 'kh', 'د': 'd', 'ذ': 'z',
            'ر': 'r', 'ز': 'z', 'ژ': 'zh', 'س': 's', 'ش': 'sh', 'ص': 's',
            'ض': 'z', 'ط': 't', 'ظ': 'z', 'ع': 'a', 'غ': 'gh', 'ف': 'f',
            'ق': 'gh', 'ک': 'k', 'گ': 'g', 'ل': 'l', 'م': 'm', 'ن': 'n',
            'و': 'v', 'ه': 'h', 'ی': 'y', ' ': ' '
        }

        # Layer 4: Descriptive Terms (Stop Words)
        self.stop_words = [
            "شرکت", "گروه", "صنعت", "صنعتی", "گستر", "گسترش", "پرداز", "پردازش", 
            "فرا", "ایمن", "سازان", "سازه", "سیستم", "پارس", "نوین", "مهر", 
            "تک", "برتر", "طلایی", "سبز", "جنوب", "شمال", "شرق", "غرب", "مرکزی",
            "ایرانیان", "توسعه", "فناوری", "خدمات", "مهندسی", "بازرگانی", "بین الملل",
            "پخش", "تولیدی", "آریا", "کیان", "پویا"
        ]

    def extract_core_root(self, text):
        """Layer 4: Strips descriptive terms to find the root brand name."""
        words = text.split()
        # Filter out stop words, but ensure at least one word remains
        core_words = [w for w in words if w not in self.stop_words]
        
        if not core_words: # If all words were stop words (e.g. "شرکت توسعه")
            return text # Return original
        
        return " ".join(core_words)

    def generate_permutations(self, text):
        """Generates variants based on Phonetic and Visual layers."""
        chars = list(text)
        options = []
        for char in chars:
            found = False
            for group in self.all_subs:
                if char in group:
                    options.append(group)
                    found = True
                    break
            if not found: options.append([char])
        
        # Cartesian product to generate all combinations
        raw_variants = list(set([''.join(p) for p in itertools.product(*options)]))
        return raw_variants

    def analyze_name(self, raw_input):
        """
        Main Analysis Engine executing all 4 layers.
        Input: "شرکت تک نان جنوب (Tak Nan)"
        """
        analysis_report = {}
        variants_to_search = set()

        # 0. Pre-processing: Extract Translation if provided in ()
        # Example: "سیب (Apple)" -> persian="سیب", translation="Apple"
        translation = ""
        persian_part = raw_input
        match = re.search(r'\((.*?)\)', raw_input)
        if match:
            translation = match.group(1)
            persian_part = raw_input.replace(f"({translation})", "").strip()
            variants_to_search.add(translation) # Add explicit translation
            analysis_report['Layer 3 (Translation)'] = translation

        # 1. Clean Input
        clean_name = re.sub(r'[^\w\s]', '', persian_part).strip()
        variants_to_search.add(clean_name)

        # 2. Layer 4: Extract Core Root
        core_root = self.extract_core_root(clean_name)
        analysis_report['Layer 4 (Core Root)'] = core_root
        
        if core_root != clean_name:
            variants_to_search.add(core_root)

        # 3. Layer 1 & 2: Homophones & Visuals (Applied to Core Root)
        # We limit permutations to avoid thousands of queries for long names
        if len(core_root) < 15: 
            perms = self.generate_permutations(core_root)
            # Limit strictly to avoid blocking (max 10 most relevant variations)
            # Or just add them all if count is reasonable
            if len(perms) < 32:
                variants_to_search.update(perms)
            else:
                variants_to_search.update(perms[:15]) # Take first 15 variations
        
        # 4. Layer 3: Fingilish (Applied to Core Root)
        fingilish = "".join([self.f_to_e_map.get(c, c) for c in core_root])
        variants_to_search.add(fingilish)
        analysis_report['Layer 3 (Fingilish)'] = fingilish

        return list(variants_to_search), analysis_report

class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    captcha_signal = pyqtSignal(bytes)
    result_signal = pyqtSignal(dict)
    finished_signal = pyqtSignal()
    progress_signal = pyqtSignal(int, int)
    status_signal = pyqtSignal(str)

    def __init__(self, file_path, classes):
        super().__init__()
        self.file_path = file_path
        self.classes = classes
        self.processor = TextProcessor()
        self.driver = None
        self.is_running = True
        self.is_paused = False
        self.captcha_code = None
        self.waiting_for_captcha = False

    def setup_driver(self):
        opts = webdriver.ChromeOptions()
        opts.add_argument('--start-maximized')
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.driver = webdriver.Chrome(options=opts)
        self.wait = WebDriverWait(self.driver, 15)

    def receive_captcha(self, code):
        self.captcha_code = code
        self.waiting_for_captcha = False

    def pause_check(self):
        while self.is_paused:
            time.sleep(0.5)

    def get_captcha_image(self):
        try:
            element = self.driver.find_element(By.ID, 'imgCaptcha')
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
            time.sleep(0.5)
            png = element.screenshot_as_png
            self.captcha_signal.emit(png)
            self.log_signal.emit("⌨️ لطفا کد کپچا را وارد کنید...")
            self.status_signal.emit("منتظر ورودی کاربر...")
            
            self.waiting_for_captcha = True
            while self.waiting_for_captcha and self.is_running:
                time.sleep(0.5)
            
            self.status_signal.emit("در حال تحلیل...")
            return self.captcha_code
        except Exception as e:
            self.log_signal.emit(f"خطا در دریافت تصویر کپچا: {e}")
            return "00000"

    def run(self):
        self.setup_driver()
        self.driver.get("https://ipm.ssaa.ir/Search-Trademark")
        
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                names = [l.strip() for l in f if l.strip()]
        except Exception as e:
            self.log_signal.emit(f"خطا در خواندن فایل: {e}")
            return

        try:
            time.sleep(1)
            btns = self.driver.find_elements(By.XPATH, "//button[contains(text(), 'متوجه شدم')]")
            for b in btns: 
                if b.is_displayed(): b.click()
        except: pass

        current_captcha = self.get_captcha_image()
        total = len(names)
        
        for idx, name in enumerate(names):
            if not self.is_running: break
            self.pause_check()
            self.progress_signal.emit(idx+1, total)
            
            # --- 4-LAYER ANALYSIS ---
            self.log_signal.emit(f"⚙️ آنالیز ۴ لایه روی: {name}")
            variants, report = self.processor.analyze_name(name)
            self.log_signal.emit(f"   > ریشه اصلی: {report.get('Layer 4 (Core Root)', '-')}")
            self.log_signal.emit(f"   > تعداد کل جایگشت‌ها: {len(variants)}")

            found_any_conflict = False # To track overall status for this name

            for variant in variants:
                if not self.is_running: break
                self.pause_check()
                
                success = False
                while not success and self.is_running:
                    try:
                        self.driver.find_element(By.ID, "ItemTitle").clear()
                        self.driver.find_element(By.ID, "ItemTitle").send_keys(variant)
                        
                        cls_inp = self.driver.find_element(By.ID, "SignProductId")
                        if cls_inp.get_attribute('value') != self.classes:
                            cls_inp.clear(); cls_inp.send_keys(self.classes)

                        self.driver.find_element(By.ID, "txtCaptcha").clear()
                        self.driver.find_element(By.ID, "txtCaptcha").send_keys(current_captcha)
                        
                        self.driver.find_element(By.ID, "LogIn").click()

                        # Check Alert
                        try:
                            WebDriverWait(self.driver, 1.5).until(EC.alert_is_present())
                            alert = self.driver.switch_to.alert
                            err = alert.text
                            alert.accept()
                            
                            if "کد امنیتی" in err or "اشتباه" in err:
                                self.log_signal.emit("❌ کپچا منقضی شد. دریافت مجدد...")
                                img = self.driver.find_element(By.ID, "imgCaptcha")
                                self.driver.execute_script("arguments[0].click();", img)
                                time.sleep(1.5)
                                current_captcha = self.get_captcha_image()
                                continue
                            else:
                                self.log_signal.emit(f"⚠️ پیام سایت: {err}")
                                break 
                        except TimeoutException:
                            pass

                        # Process Results
                        time.sleep(2)
                        res_box = self.driver.find_element(By.CSS_SELECTOR, ".result")
                        
                        if "رکوردی یافت نشد" in res_box.text:
                            # Clean pass
                            pass 
                        else:
                            items = self.scrape_all_pages(name, variant)
                            if items > 0:
                                found_any_conflict = True

                        success = True
                        time.sleep(random.uniform(1.5, 3.5))

                    except Exception as e:
                        self.log_signal.emit(f"⛔ خطای شبکه: {str(e)[:50]}")
                        try: self.driver.refresh()
                        except: pass
                        time.sleep(3)
                        current_captcha = self.get_captcha_image()
            
            # Final Report for this Name
            if not found_any_conflict:
                self.result_signal.emit({
                    "Search Term": name, "Variant": "مجموع بررسی‌ها", 
                    "Status": "آزاد (تأیید شده)", 
                    "Brand": "---", "Reg No": "-", "Owner": "بدون سابقه تعارض", "Goods": "-"
                })

        self.driver.quit()
        self.finished_signal.emit()

    def scrape_all_pages(self, search_term, variant):
        page = 1
        total_items = 0
        while self.is_running:
            links = self.driver.find_elements(By.CSS_SELECTOR, ".result > a")
            main_window = self.driver.current_window_handle
            
            if not links: break
            total_items += len(links)

            for i in range(len(links)):
                if not self.is_running: return total_items
                try:
                    links = self.driver.find_elements(By.CSS_SELECTOR, ".result > a")
                    if i >= len(links): break
                    item = links[i]
                    
                    try: title = item.find_element(By.TAG_NAME, "h2").text.strip()
                    except: title = "Unknown"

                    self.driver.execute_script("arguments[0].click();", item)
                    time.sleep(3)

                    if len(self.driver.window_handles) > 1:
                        new_win = [h for h in self.driver.window_handles if h != main_window][0]
                        self.driver.switch_to.window(new_win)
                        self.extract_data(search_term, variant, title)
                        self.driver.close()
                        self.driver.switch_to.window(main_window)
                    else:
                        try:
                            self.wait.until(EC.visibility_of_element_located((By.ID, "C_Spec")))
                            self.extract_data(search_term, variant, title, is_modal=True)
                            webdriver.ActionChains(self.driver).send_keys('\ue00c').perform() 
                            time.sleep(1)
                        except: pass
                except: pass

            try:
                next_btns = self.driver.find_elements(By.XPATH, "//div[contains(@onclick, \"goto('next')\")]")
                if not next_btns:
                     next_btns = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'pager_but')][contains(@onclick, 'next')]")
                
                if next_btns and next_btns[0].is_displayed():
                    self.driver.execute_script("arguments[0].click();", next_btns[0])
                    time.sleep(4)
                    page += 1
                else:
                    break
            except: break
        
        return total_items

    def extract_data(self, st, var, title, is_modal=False):
        try:
            xpath = "//div[@id='C_Spec']" if is_modal else ""
            def gv(lbl):
                return self.driver.find_element(By.XPATH, f"{xpath}//td[contains(text(), '{lbl}')]/following-sibling::td").text.strip()
            
            data = {
                "Search Term": st, "Variant": var, "Status": "دارای تعارض",
                "Brand": title, "Reg No": gv("شماره ثبت"),
                "Owner": gv("نام مالک"), "Goods": gv("کالاها")[:100]
            }
            self.result_signal.emit(data)
        except:
            self.result_signal.emit({"Search Term": st, "Variant": var, "Status": "خطا", "Brand": title, "Owner": "خطا در خواندن", "Reg No": "-", "Goods": "-"})

    def stop(self): self.is_running = False
    def toggle_pause(self): self.is_paused = not self.is_paused

# ================= 2. GUI APPLICATION =================

class FinalApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("سامانه هوش تجاری آنالیز برند (نسخه 6.0)")
        self.resize(1280, 900)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.results = []
        self.setup_ui()
        
    def setup_ui(self):
        font = QFont("Vazirmatn", 10)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        QApplication.setFont(font)

        self.setStyleSheet("""
            QMainWindow { background-color: #f0f2f5; color: #212529; }
            QGroupBox { background-color: white; border-radius: 8px; border: 1px solid #dee2e6; margin-top: 10px; font-weight: bold; padding-top: 20px; }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top right; right: 10px; padding: 0 5px; color: #0d6efd; }
            QPushButton { background-color: #0d6efd; color: white; border-radius: 5px; padding: 8px 20px; font-weight: bold; border: none; }
            QPushButton:hover { background-color: #0b5ed7; }
            QPushButton#stopBtn { background-color: #dc3545; }
            QPushButton#exportBtn { background-color: #198754; }
            QLineEdit, QListWidget { border: 1px solid #ced4da; border-radius: 4px; padding: 6px; background: #fff; }
            QTableWidget { border: 1px solid #dee2e6; background: white; gridline-color: #f1f3f5; }
            QHeaderView::section { background-color: #e9ecef; padding: 8px; border: none; font-weight: bold; color: #495057; }
            QLabel#captchaBox { border: 2px dashed #adb5bd; border-radius: 6px; background: #f8f9fa; }
            QLabel#designer { color: #6c757d; font-size: 11px; padding: 5px; }
        """)

        # Menu
        menubar = self.menuBar()
        file_menu = menubar.addMenu("پرونده")
        load_action = QAction("بارگذاری فایل اسامی", self)
        load_action.triggered.connect(self.browse_file)
        file_menu.addAction(load_action)
        export_action = QAction("خروجی اکسل", self)
        export_action.triggered.connect(self.export_data)
        file_menu.addAction(export_action)
        
        help_menu = menubar.addMenu("راهنما")
        about_action = QAction("درباره طراح", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # Top
        top_layout = QHBoxLayout()
        conf_group = QGroupBox("تنظیمات آنالیز ۴ لایه")
        conf_layout = QVBoxLayout()
        file_row = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("مسیر فایل Name.txt ...")
        self.path_edit.setReadOnly(True)
        btn_browse = QPushButton("📂 انتخاب فایل")
        btn_browse.clicked.connect(self.browse_file)
        file_row.addWidget(btn_browse)
        file_row.addWidget(self.path_edit)
        class_row = QHBoxLayout()
        class_row.addWidget(QLabel("کد طبقات:"))
        self.class_edit = QLineEdit("5,31")
        class_row.addWidget(self.class_edit)
        conf_layout.addLayout(file_row)
        conf_layout.addLayout(class_row)
        conf_group.setLayout(conf_layout)
        
        cap_group = QGroupBox("پنل امنیتی")
        cap_layout = QHBoxLayout()
        self.lbl_captcha_img = QLabel("تصویر کپچا")
        self.lbl_captcha_img.setObjectName("captchaBox")
        self.lbl_captcha_img.setFixedSize(180, 70)
        self.lbl_captcha_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        cap_input_col = QVBoxLayout()
        self.txt_captcha = QLineEdit()
        self.txt_captcha.setPlaceholderText("کد را وارد کنید...")
        self.txt_captcha.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.txt_captcha.setStyleSheet("font-size: 16px; font-weight: bold; letter-spacing: 2px;")
        self.txt_captcha.returnPressed.connect(self.send_captcha)
        self.btn_send_cap = QPushButton("ارسال")
        self.btn_send_cap.clicked.connect(self.send_captcha)
        self.btn_send_cap.setEnabled(False)
        cap_input_col.addWidget(self.txt_captcha)
        cap_input_col.addWidget(self.btn_send_cap)
        cap_layout.addWidget(self.lbl_captcha_img)
        cap_layout.addLayout(cap_input_col)
        cap_group.setLayout(cap_layout)

        top_layout.addWidget(conf_group, 60)
        top_layout.addWidget(cap_group, 40)
        layout.addLayout(top_layout)

        # Controls
        ctrl_layout = QHBoxLayout()
        self.btn_start = QPushButton("شروع آنالیز و جستجو")
        self.btn_start.clicked.connect(self.start_process)
        self.btn_pause = QPushButton("توقف موقت")
        self.btn_pause.setObjectName("pauseBtn")
        self.btn_pause.clicked.connect(self.pause_process)
        self.btn_pause.setEnabled(False)
        self.btn_stop = QPushButton("پایان عملیات")
        self.btn_stop.setObjectName("stopBtn")
        self.btn_stop.clicked.connect(self.stop_process)
        self.btn_stop.setEnabled(False)
        self.lbl_status = QLabel("وضعیت: آماده")
        self.lbl_status.setStyleSheet("color: #6c757d; font-weight: bold; margin-right: 10px;")
        ctrl_layout.addWidget(self.btn_start)
        ctrl_layout.addWidget(self.btn_pause)
        ctrl_layout.addWidget(self.btn_stop)
        ctrl_layout.addStretch()
        ctrl_layout.addWidget(self.lbl_status)
        layout.addLayout(ctrl_layout)

        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        self.progress.setFormat("پیشرفت کل: %p%")
        layout.addWidget(self.progress)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        headers = ["نام اصلی", "واریانت بررسی شده", "وضعیت تعارض", "برند پیدا شده", "شماره ثبت", "مالک", "کالاها"]
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        layout.addWidget(self.table)

        # Summary Lists
        summary_group = QGroupBox("داشبورد مدیریتی نتایج")
        summary_layout = QHBoxLayout()
        
        approved_layout = QVBoxLayout()
        approved_label = QLabel("✅ لیست آزاد (تأیید شده برای ثبت):")
        approved_label.setStyleSheet("color: #198754; font-weight: bold;")
        self.list_approved = QListWidget()
        self.list_approved.setStyleSheet("border: 1px solid #198754; background: #f0fff4;")
        approved_layout.addWidget(approved_label)
        approved_layout.addWidget(self.list_approved)
        
        rejected_layout = QVBoxLayout()
        rejected_label = QLabel("❌ لیست دارای تعارض (رد شده):")
        rejected_label.setStyleSheet("color: #dc3545; font-weight: bold;")
        self.list_rejected = QListWidget()
        self.list_rejected.setStyleSheet("border: 1px solid #dc3545; background: #fff5f5;")
        rejected_layout.addWidget(rejected_label)
        rejected_layout.addWidget(self.list_rejected)
        
        summary_layout.addLayout(approved_layout)
        summary_layout.addLayout(rejected_layout)
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        # Footer
        footer_layout = QHBoxLayout()
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMaximumHeight(60)
        self.log_box.setStyleSheet("background: #f8f9fa; font-size: 10px;")
        btn_export = QPushButton("ذخیره گزارش کامل")
        btn_export.setObjectName("exportBtn")
        btn_export.setFixedSize(120, 60)
        btn_export.clicked.connect(self.export_data)
        footer_layout.addWidget(self.log_box)
        footer_layout.addWidget(btn_export)
        layout.addLayout(footer_layout)

        lbl_designer = QLabel("طراحی و توسعه: هادی علایی | مدیر واحد فناوری اطلاعات")
        lbl_designer.setObjectName("designer")
        lbl_designer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_designer)

    def browse_file(self):
        f, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل متنی", "", "Text Files (*.txt)")
        if f: self.path_edit.setText(f)

    def show_about(self):
        QMessageBox.information(self, "درباره طراح", "طراح: هادی علایی\nنسخه: 6.0 (BI Edition)")

    def log(self, msg):
        t = datetime.now().strftime("%H:%M:%S")
        self.log_box.append(f"[{t}] {msg}")
        self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())

    def update_captcha_img(self, img_data):
        pixmap = QPixmap()
        pixmap.loadFromData(img_data)
        self.lbl_captcha_img.setPixmap(pixmap.scaled(160, 60, Qt.AspectRatioMode.KeepAspectRatio))
        self.txt_captcha.setEnabled(True)
        self.btn_send_cap.setEnabled(True)
        self.txt_captcha.setFocus()
        QApplication.alert(self)

    def send_captcha(self):
        code = self.txt_captcha.text()
        if code:
            self.worker.receive_captcha(code)
            self.txt_captcha.clear()
            self.txt_captcha.setEnabled(False)
            self.btn_send_cap.setEnabled(False)
            self.lbl_captcha_img.setText("کد ارسال شد...")

    def add_result(self, data):
        self.results.append(data)
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        items = [
            data["Search Term"], data["Variant"], data["Status"], 
            data["Brand"], data["Reg No"], data["Owner"], data["Goods"]
        ]
        
        status = data["Status"]
        
        # Logic for Summary Lists & Colors
        if "آزاد" in status:
            color = QColor("#d1e7dd")
            text_color = QColor("#0f5132")
            # Only add to approved list if it's the final summary row for the name
            if data["Variant"] == "مجموع بررسی‌ها":
                self.list_approved.addItem(f"✅ {data['Search Term']}")
        elif "تعارض" in status or "مشابه" in status:
            color = QColor("#f8d7da")
            text_color = QColor("#842029")
            self.list_rejected.addItem(f"⚠️ {data['Search Term']} -> {data['Brand']} ({data['Variant']})")
        else:
            color = QColor("#fff3cd")
            text_color = QColor("#664d03")

        for col, text in enumerate(items):
            item = QTableWidgetItem(str(text))
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if col == 2:
                item.setBackground(QBrush(color))
                item.setForeground(QBrush(text_color))
                item.setFont(QFont("Vazirmatn", 10, QFont.Weight.Bold))
            self.table.setItem(row, col, item)
        
        self.table.scrollToBottom()
        self.list_approved.scrollToBottom()
        self.list_rejected.scrollToBottom()

    def start_process(self):
        if not self.path_edit.text():
            QMessageBox.warning(self, "خطا", "لطفا فایل اسامی (Name.txt) را انتخاب کنید.")
            return
        self.btn_start.setEnabled(False)
        self.btn_pause.setEnabled(True)
        self.btn_stop.setEnabled(True)
        self.table.setRowCount(0)
        self.list_approved.clear()
        self.list_rejected.clear()
        self.results = []
        
        self.worker = WorkerThread(self.path_edit.text(), self.class_edit.text())
        self.worker.log_signal.connect(self.log)
        self.worker.captcha_signal.connect(self.update_captcha_img)
        self.worker.result_signal.connect(self.add_result)
        self.worker.progress_signal.connect(lambda c, t: (self.progress.setMaximum(t), self.progress.setValue(c)))
        self.worker.status_signal.connect(self.lbl_status.setText)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.start()

    def pause_process(self):
        if self.worker.is_paused:
            self.worker.toggle_pause()
            self.btn_pause.setText("توقف موقت")
            self.lbl_status.setText("وضعیت: در حال اجرا")
        else:
            self.worker.toggle_pause()
            self.btn_pause.setText("ادامه عملیات")
            self.lbl_status.setText("وضعیت: متوقف شده")

    def stop_process(self):
        if self.worker:
            self.worker.stop()
            self.lbl_status.setText("وضعیت: در حال توقف...")

    def process_finished(self):
        self.btn_start.setEnabled(True)
        self.btn_pause.setEnabled(False)
        self.btn_stop.setEnabled(False)
        self.lbl_status.setText("وضعیت: پایان یافته")
        QMessageBox.information(self, "پایان", "آنالیز جامع به پایان رسید.")

    def export_data(self):
        if not self.results: return
        default_name = f"BI_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "ذخیره اکسل", default_name, "Excel Files (*.xlsx)")
        if path:
            try:
                clean_results = []
                ILLEGAL_CHARACTERS_RE = r'[\000-\010]|[\013-\014]|[\016-\037]'
                for row in self.results:
                    clean_row = {}
                    for k, v in row.items():
                        if isinstance(v, str): clean_row[k] = re.sub(ILLEGAL_CHARACTERS_RE, "", v)
                        else: clean_row[k] = v
                    clean_results.append(clean_row)
                pd.DataFrame(clean_results).to_excel(path, index=False)
                QMessageBox.information(self, "موفق", f"فایل ذخیره شد:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در ذخیره‌سازی: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FinalApp()
    window.show()
    sys.exit(app.exec())
