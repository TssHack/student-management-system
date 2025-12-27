import sys
import sqlite3
import hashlib
import os
import datetime
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
                             QHeaderView, QFileDialog, QMessageBox, QGroupBox, QFormLayout, 
                             QDoubleSpinBox, QStatusBar, QMenuBar, QMenu, QAction, QDialog,
                             QDialogButtonBox, QProgressBar, QSizePolicy, QFrame, QSplitter, QGridLayout) # QGridLayout اضافه شد
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtGui import QFont, QPixmap, QPainter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib

matplotlib.use('Qt5Agg')
matplotlib.rcParams['font.family'] = 'B Nazanin'
matplotlib.rcParams['axes.unicode_minus'] = False

class StudentManagementSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.create_database()
        self.load_students()
        self.update_dashboard()
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
    
    def initUI(self):
        self.setWindowTitle("مدیریت دانشجویان")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 600)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setFont(QFont("B Nazanin", 10))
        
        self.create_menu()
        self.create_toolbar()
        
        central_widget = QWidget()
        central_widget.setLayoutDirection(Qt.RightToLeft)
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        
        splitter = QSplitter(Qt.Vertical)
        splitter.setLayoutDirection(Qt.RightToLeft)
        main_layout.addWidget(splitter)
        
        top_widget = QWidget()
        top_widget.setLayoutDirection(Qt.RightToLeft)
        top_layout = QVBoxLayout(top_widget)
        
        self.create_dashboard(top_layout)
        self.create_search_section(top_layout)
        
        splitter.addWidget(top_widget)
        
        table_widget = QWidget()
        table_widget.setLayoutDirection(Qt.RightToLeft)
        table_layout = QVBoxLayout(table_widget)
        self.create_student_table(table_layout)
        
        splitter.addWidget(table_widget)
        splitter.setSizes([200, 600])
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.setLayoutDirection(Qt.RightToLeft)
        self.status_bar.showMessage("آماده به کار")
        
        self.time_label = QLabel()
        self.time_label.setLayoutDirection(Qt.RightToLeft)
        self.status_bar.addPermanentWidget(self.time_label)
        self.update_time()
    
    def create_menu(self):
        menubar = self.menuBar()
        menubar.setLayoutDirection(Qt.RightToLeft)
        
        file_menu = menubar.addMenu("فایل")
        file_menu.setLayoutDirection(Qt.RightToLeft)
        
        export_excel_action = QAction("خروجی اکسل", self)
        export_excel_action.triggered.connect(self.export_to_excel)
        file_menu.addAction(export_excel_action)
        
        import_excel_action = QAction("وارد کردن از اکسل", self)
        import_excel_action.triggered.connect(self.import_from_excel)
        file_menu.addAction(import_excel_action)
        
        file_menu.addSeparator()
        
        backup_action = QAction("پشتیبان‌گیری", self)
        backup_action.triggered.connect(self.backup_database)
        file_menu.addAction(backup_action)
        
        restore_action = QAction("بازیابی پشتیبان", self)
        restore_action.triggered.connect(self.restore_database)
        file_menu.addAction(restore_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("خروج", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        edit_menu = menubar.addMenu("ویرایش")
        edit_menu.setLayoutDirection(Qt.RightToLeft)
        
        add_student_action = QAction("افزودن دانشجو", self)
        add_student_action.triggered.connect(self.add_student_dialog)
        edit_menu.addAction(add_student_action)
        
        edit_student_action = QAction("ویرایش دانشجو", self)
        edit_student_action.triggered.connect(self.edit_student_dialog)
        edit_menu.addAction(edit_student_action)
        
        delete_student_action = QAction("حذف دانشجو", self)
        delete_student_action.triggered.connect(self.delete_student)
        edit_menu.addAction(delete_student_action)
        
        tools_menu = menubar.addMenu("ابزارها")
        tools_menu.setLayoutDirection(Qt.RightToLeft)
        
        rank_action = QAction("رتبه‌بندی دانشجویان", self)
        rank_action.triggered.connect(self.rank_students)
        tools_menu.addAction(rank_action)
        
        chart_action = QAction("نمودار آماری", self)
        chart_action.triggered.connect(self.show_statistics_chart)
        tools_menu.addAction(chart_action)
        
        report_action = QAction("چاپ گزارش", self)
        report_action.triggered.connect(self.print_report)
        tools_menu.addAction(report_action)
        
        help_menu = menubar.addMenu("راهنما")
        help_menu.setLayoutDirection(Qt.RightToLeft)
        
        about_action = QAction("درباره برنامه", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def create_toolbar(self):
        toolbar = self.addToolBar("نوار ابزار")
        toolbar.setMovable(False)
        toolbar.setLayoutDirection(Qt.RightToLeft)
        toolbar.setIconSize(QSize(24, 24))
        
        add_btn = QPushButton("➕ افزودن دانشجو")
        add_btn.clicked.connect(self.add_student_dialog)
        toolbar.addWidget(add_btn)
        
        edit_btn = QPushButton("✏️ ویرایش")
        edit_btn.clicked.connect(self.edit_student_dialog)
        toolbar.addWidget(edit_btn)
        
        delete_btn = QPushButton("🗑️ حذف")
        delete_btn.clicked.connect(self.delete_student)
        toolbar.addWidget(delete_btn)
        
        toolbar.addSeparator()
        
        rank_btn = QPushButton("🏆 رتبه‌بندی")
        rank_btn.clicked.connect(self.rank_students)
        toolbar.addWidget(rank_btn)
        
        excel_btn = QPushButton("📊 خروجی اکسل")
        excel_btn.clicked.connect(self.export_to_excel)
        toolbar.addWidget(excel_btn)
        
        toolbar.addSeparator()
        
        title_label = QLabel("مدیریت دانشجویان")
        title_label.setFont(QFont("B Nazanin", 12, QFont.Bold))
        toolbar.addWidget(title_label)
        
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        toolbar.addWidget(spacer)
    
    def create_dashboard(self, parent_layout):
        dashboard_group = QGroupBox("داشبورد آماری")
        dashboard_group.setLayoutDirection(Qt.RightToLeft)
        dashboard_layout = QHBoxLayout()
        
        self.students_count_card = self.create_stat_card("تعداد دانشجویان", "0", "#3498db")
        dashboard_layout.addWidget(self.students_count_card)
        
        self.avg_grade_card = self.create_stat_card("میانگین کل", "0.00", "#9b59b6")
        dashboard_layout.addWidget(self.avg_grade_card)
        
        self.max_grade_card = self.create_stat_card("بالاترین معدل", "0.00", "#e74c3c")
        dashboard_layout.addWidget(self.max_grade_card)
        
        self.excellent_card = self.create_stat_card("دانشجویان ممتاز", "0", "#f39c12")
        dashboard_layout.addWidget(self.excellent_card)
        
        dashboard_group.setLayout(dashboard_layout)
        parent_layout.addWidget(dashboard_group)
    
    def create_stat_card(self, title, value, color):
        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)
        card.setFrameShadow(QFrame.Raised)
        card.setMinimumWidth(180)
        card.setLayoutDirection(Qt.RightToLeft)
        card.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border: 1px solid #e0e0e0;
            }}
        """)
        
        layout = QVBoxLayout()
        
        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        title_label.setStyleSheet("font-weight: bold; color: #555;")
        layout.addWidget(title_label)
        
        value_label = QLabel(value)
        value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        value_label.setStyleSheet(f"font-size: 24px; font-weight: bold; color: {color};")
        layout.addWidget(value_label)
        
        card.setLayout(layout)
        card.value_label = value_label
        
        return card
    
    def create_search_section(self, parent_layout):
        search_group = QGroupBox("جستجوی پیشرفته")
        search_group.setLayoutDirection(Qt.RightToLeft)
        search_layout = QGridLayout()
        
        search_layout.addWidget(QLabel("نام:"), 0, 0, 1, 1, Qt.AlignRight)
        self.search_name_edit = QLineEdit()
        self.search_name_edit.setPlaceholderText("نام دانشجو")
        self.search_name_edit.setLayoutDirection(Qt.RightToLeft)
        search_layout.addWidget(self.search_name_edit, 0, 1, 1, 1)
        
        search_layout.addWidget(QLabel("نام خانوادگی:"), 0, 2, 1, 1, Qt.AlignRight)
        self.search_lastname_edit = QLineEdit()
        self.search_lastname_edit.setPlaceholderText("نام خانوادگی دانشجو")
        self.search_lastname_edit.setLayoutDirection(Qt.RightToLeft)
        search_layout.addWidget(self.search_lastname_edit, 0, 3, 1, 1)
        
        search_layout.addWidget(QLabel("شماره دانشجویی:"), 0, 4, 1, 1, Qt.AlignRight)
        self.search_id_edit = QLineEdit()
        self.search_id_edit.setPlaceholderText("شماره دانشجویی")
        self.search_id_edit.setLayoutDirection(Qt.RightToLeft)
        search_layout.addWidget(self.search_id_edit, 0, 5, 1, 1)
        
        search_layout.addWidget(QLabel("محدوده معدل:"), 1, 0, 1, 1, Qt.AlignRight)
        self.search_min_avg_edit = QLineEdit()
        self.search_min_avg_edit.setPlaceholderText("حداقل")
        self.search_min_avg_edit.setLayoutDirection(Qt.RightToLeft)
        search_layout.addWidget(self.search_min_avg_edit, 1, 1, 1, 1)
        
        search_layout.addWidget(QLabel("تا"), 1, 2, 1, 1, Qt.AlignCenter)
        self.search_max_avg_edit = QLineEdit()
        self.search_max_avg_edit.setPlaceholderText("حداکثر")
        self.search_max_avg_edit.setLayoutDirection(Qt.RightToLeft)
        search_layout.addWidget(self.search_max_avg_edit, 1, 3, 1, 1)
        
        search_btn = QPushButton("🔍 جستجو")
        search_btn.clicked.connect(self.advanced_search)
        search_btn.setStyleSheet("background-color: #3498db; color: white; font-weight: bold;")
        search_layout.addWidget(search_btn, 1, 4, 1, 1)
        
        reset_btn = QPushButton("↺ نمایش همه")
        reset_btn.clicked.connect(self.load_students)
        search_layout.addWidget(reset_btn, 1, 5, 1, 1)
        
        search_group.setLayout(search_layout)
        parent_layout.addWidget(search_group)
    
    def create_student_table(self, parent_layout):
        table_group = QGroupBox("لیست دانشجویان")
        table_group.setLayoutDirection(Qt.RightToLeft)
        table_layout = QVBoxLayout()
        
        self.student_table = QTableWidget()
        self.student_table.setColumnCount(8)
        self.student_table.setHorizontalHeaderLabels([
            "ردیف", "نام", "نام خانوادگی", "شماره دانشجویی", 
            "نمره میانترم", "نمره پایان‌ترم", "معدل", "رتبه"
        ])
        
        self.student_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.student_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.student_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.student_table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)
        self.student_table.setAlternatingRowColors(True)
        self.student_table.setSortingEnabled(True)
        self.student_table.setLayoutDirection(Qt.RightToLeft)
        self.student_table.doubleClicked.connect(self.show_student_details)
        
        table_layout.addWidget(self.student_table)
        table_group.setLayout(table_layout)
        parent_layout.addWidget(table_group)
    
    def create_database(self):
        try:
            self.conn = sqlite3.connect('students.db')
            self.cursor = self.conn.cursor()
            
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    first_name TEXT NOT NULL,
                    last_name TEXT NOT NULL,
                    student_id TEXT UNIQUE NOT NULL,
                    midterm REAL,
                    final REAL,
                    average REAL,
                    rank INTEGER,
                    registration_date TEXT,
                    photo_path TEXT
                )
            ''')
            
            self.cursor.execute("PRAGMA table_info(students)")
            columns = [column[1] for column in self.cursor.fetchall()]
            
            if 'registration_date' not in columns:
                self.cursor.execute("ALTER TABLE students ADD COLUMN registration_date TEXT")
            
            if 'photo_path' not in columns:
                self.cursor.execute("ALTER TABLE students ADD COLUMN photo_path TEXT")
            
            self.conn.commit()
            
        except Exception as e:
            QMessageBox.critical(self, "خطا در پایگاه داده", f"خطا در ایجاد پایگاه داده: {str(e)}")
            sys.exit(1)
    
    def hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()
    
    def load_students(self):
        try:
            self.cursor.execute("SELECT * FROM students ORDER BY id")
            students = self.cursor.fetchall()
            
            self.student_table.setRowCount(0)
            
            for row_idx, student in enumerate(students):
                self.student_table.insertRow(row_idx)
                cols_to_show = [0, 1, 2, 3, 4, 5, 6, 7]
                
                for table_col_idx, db_col_idx in enumerate(cols_to_show):
                    data = student[db_col_idx] if db_col_idx < len(student) else None
                    if data is None:
                        data = "-"
                    item = QTableWidgetItem(str(data))
                    item.setTextAlignment(Qt.AlignCenter)
                    self.student_table.setItem(row_idx, table_col_idx, item)
            
            self.update_dashboard()
            self.status_bar.showMessage(f"تعداد {len(students)} دانشجو بارگذاری شد")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در بارگذاری دانشجویان: {str(e)}")

    def update_dashboard(self):
        try:
            self.cursor.execute("SELECT COUNT(*) FROM students")
            total_students = self.cursor.fetchone()[0]
            
            self.cursor.execute("SELECT AVG(average) FROM students WHERE average IS NOT NULL")
            avg_result = self.cursor.fetchone()[0]
            avg_grade = avg_result if avg_result else 0
            
            self.cursor.execute("SELECT MAX(average) FROM students WHERE average IS NOT NULL")
            max_result = self.cursor.fetchone()[0]
            max_grade = max_result if max_result else 0
            
            self.cursor.execute("SELECT COUNT(*) FROM students WHERE average >= 17")
            excellent_count = self.cursor.fetchone()[0]
            
            self.students_count_card.value_label.setText(str(total_students))
            self.avg_grade_card.value_label.setText(f"{avg_grade:.2f}")
            self.max_grade_card.value_label.setText(f"{max_grade:.2f}")
            self.excellent_card.value_label.setText(str(excellent_count))
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در به‌روزرسانی داشبورد: {str(e)}")

    def advanced_search(self):
        try:
            query = "SELECT * FROM students WHERE 1=1"
            params = []
            
            if self.search_name_edit.text().strip():
                query += " AND first_name LIKE ?"
                params.append(f"%{self.search_name_edit.text().strip()}%")
            
            if self.search_lastname_edit.text().strip():
                query += " AND last_name LIKE ?"
                params.append(f"%{self.search_lastname_edit.text().strip()}%")
            
            if self.search_id_edit.text().strip():
                query += " AND student_id LIKE ?"
                params.append(f"%{self.search_id_edit.text().strip()}%")
            
            if self.search_min_avg_edit.text().strip():
                try:
                    min_avg = float(self.search_min_avg_edit.text())
                    query += " AND average >= ?"
                    params.append(min_avg)
                except ValueError:
                    pass
            
            if self.search_max_avg_edit.text().strip():
                try:
                    max_avg = float(self.search_max_avg_edit.text())
                    query += " AND average <= ?"
                    params.append(max_avg)
                except ValueError:
                    pass
            
            self.cursor.execute(query, params)
            students = self.cursor.fetchall()
            
            self.student_table.setRowCount(0)
            
            for row_idx, student in enumerate(students):
                self.student_table.insertRow(row_idx)
                cols_to_show = [0, 1, 2, 3, 4, 5, 6, 7]
                for table_col_idx, db_col_idx in enumerate(cols_to_show):
                    data = student[db_col_idx] if db_col_idx < len(student) else None
                    if data is None:
                        data = "-"
                    item = QTableWidgetItem(str(data))
                    item.setTextAlignment(Qt.AlignCenter)
                    self.student_table.setItem(row_idx, table_col_idx, item)
            
            self.status_bar.showMessage(f"نتایج جستجو: {len(students)} دانشجو یافت شد")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در جستجو: {str(e)}")
    
    def add_student_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("افزودن دانشجوی جدید")
        dialog.setMinimumWidth(450)
        dialog.setLayoutDirection(Qt.RightToLeft)
        
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        
        first_name_edit = QLineEdit()
        first_name_edit.setPlaceholderText("نام دانشجو")
        form_layout.addRow("نام:", first_name_edit)
        
        last_name_edit = QLineEdit()
        last_name_edit.setPlaceholderText("نام خانوادگی دانشجو")
        form_layout.addRow("نام خانوادگی:", last_name_edit)
        
        student_id_edit = QLineEdit()
        student_id_edit.setPlaceholderText("شماره دانشجویی")
        form_layout.addRow("شماره دانشجویی:", student_id_edit)
        
        midterm_edit = QDoubleSpinBox()
        midterm_edit.setRange(0, 20)
        midterm_edit.setSingleStep(0.5)
        midterm_edit.setDecimals(1)
        midterm_edit.setValue(0)
        form_layout.addRow("نمره میانترم:", midterm_edit)
        
        final_edit = QDoubleSpinBox()
        final_edit.setRange(0, 20)
        final_edit.setSingleStep(0.5)
        final_edit.setDecimals(1)
        final_edit.setValue(0)
        form_layout.addRow("نمره پایان‌ترم:", final_edit)

        average_edit = QDoubleSpinBox()
        average_edit.setRange(0, 20)
        average_edit.setSingleStep(0.01)
        average_edit.setDecimals(2)
        form_layout.addRow("معدل:", average_edit)
        
        def auto_calc_average():
            mid = midterm_edit.value()
            fin = final_edit.value()
            avg = (mid * 0.3) + (fin * 0.7)
            average_edit.setValue(avg)

        midterm_edit.valueChanged.connect(auto_calc_average)
        final_edit.valueChanged.connect(auto_calc_average)
        
        photo_layout = QHBoxLayout()
        self.photo_label = QLabel()
        self.photo_label.setFixedSize(100, 100)
        self.photo_label.setFrameShape(QFrame.Box)
        self.photo_label.setAlignment(Qt.AlignCenter)
        self.show_default_photo()
        photo_layout.addWidget(self.photo_label)
        
        self.photo_path = ""
        photo_btn = QPushButton("انتخاب عکس")
        photo_btn.clicked.connect(lambda: self.select_photo(self.photo_label))
        photo_layout.addWidget(photo_btn)
        
        form_layout.addRow("عکس دانشجو:", photo_layout)
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            try:
                first_name = first_name_edit.text().strip()
                last_name = last_name_edit.text().strip()
                student_id = student_id_edit.text().strip()
                midterm = midterm_edit.value() if midterm_edit.value() > 0 else None
                final = final_edit.value() if final_edit.value() > 0 else None
                average = average_edit.value() if average_edit.value() > 0 else None
                
                if not first_name or not last_name or not student_id:
                    QMessageBox.warning(self, "خطا", "لطفاً تمام فیلدهای ضروری را پر کنید")
                    return
                
                registration_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                self.cursor.execute(
                    "INSERT INTO students (first_name, last_name, student_id, midterm, final, average, registration_date, photo_path) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (first_name, last_name, student_id, midterm, final, average, registration_date, self.photo_path)
                )
                self.conn.commit()
                
                QMessageBox.information(self, "موفقیت", "دانشجو با موفقیت اضافه شد")
                self.load_students()
                
            except ValueError:
                QMessageBox.warning(self, "خطا", "لطفاً نمرات را به صورت عددی وارد کنید")
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "خطا", "شماره دانشجویی تکراری است")
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در افزودن دانشجو: {str(e)}")
    
    def show_default_photo(self):
        pixmap = QPixmap(100, 100)
        pixmap.fill(Qt.white)
        painter = QPainter(pixmap)
        painter.setPen(Qt.black)
        painter.drawText(pixmap.rect(), Qt.AlignCenter, "عکس دانشجو")
        painter.end()
        self.photo_label.setPixmap(pixmap)
    
    def select_photo(self, label):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب عکس دانشجو", "", "Image files (*.jpg *.jpeg *.png *.bmp)")
        if file_path:
            try:
                pixmap = QPixmap(file_path)
                if not pixmap.isNull():
                    pixmap = pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    label.setPixmap(pixmap)
                    self.photo_path = file_path
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در بارگذاری عکس: {str(e)}")
    
    def edit_student_dialog(self):
        selected_items = self.student_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "هشدار", "لطفاً یک دانشجو را انتخاب کنید")
            return
        
        row = selected_items[0].row()
        student_id = self.student_table.item(row, 3).text()
        
        try:
            self.cursor.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
            student = self.cursor.fetchone()
            
            if not student:
                QMessageBox.critical(self, "خطا", "دانشجو یافت نشد")
                return
            
            if len(student) < 10:
                self.cursor.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
                student = self.cursor.fetchone()
                if not student or len(student) < 10:
                    QMessageBox.critical(self, "خطا", "داده‌های دانشجو ناقص هستند")
                    return
            
            dialog = QDialog(self)
            dialog.setWindowTitle("ویرایش اطلاعات دانشجو")
            dialog.setMinimumWidth(450)
            dialog.setLayoutDirection(Qt.RightToLeft)
            
            layout = QVBoxLayout()
            form_layout = QFormLayout()
            
            first_name_edit = QLineEdit(student[1])
            form_layout.addRow("نام:", first_name_edit)
            
            last_name_edit = QLineEdit(student[2])
            form_layout.addRow("نام خانوادگی:", last_name_edit)
            
            student_id_edit = QLineEdit(student[3])
            student_id_edit.setReadOnly(True)
            form_layout.addRow("شماره دانشجویی:", student_id_edit)
            
            midterm_edit = QDoubleSpinBox()
            midterm_edit.setRange(0, 20)
            midterm_edit.setSingleStep(0.5)
            midterm_edit.setDecimals(1)
            midterm_edit.setValue(student[4] if student[4] is not None else 0)
            form_layout.addRow("نمره میانترم:", midterm_edit)
            
            final_edit = QDoubleSpinBox()
            final_edit.setRange(0, 20)
            final_edit.setSingleStep(0.5)
            final_edit.setDecimals(1)
            final_edit.setValue(student[5] if student[5] is not None else 0)
            form_layout.addRow("نمره پایان‌ترم:", final_edit)
            
            average_edit = QDoubleSpinBox()
            average_edit.setRange(0, 20)
            average_edit.setSingleStep(0.01)
            average_edit.setDecimals(2)
            average_edit.setValue(student[6] if student[6] is not None else 0)
            form_layout.addRow("معدل:", average_edit)

            def auto_calc_average_edit():
                mid = midterm_edit.value()
                fin = final_edit.value()
                avg = (mid * 0.3) + (fin * 0.7)
                average_edit.setValue(avg)

            midterm_edit.valueChanged.connect(auto_calc_average_edit)
            final_edit.valueChanged.connect(auto_calc_average_edit)
            
            photo_layout = QHBoxLayout()
            self.photo_label = QLabel()
            self.photo_label.setFixedSize(100, 100)
            self.photo_label.setFrameShape(QFrame.Box)
            self.photo_label.setAlignment(Qt.AlignCenter)
            
            if len(student) > 9 and student[9]:
                pixmap = QPixmap(student[9])
                if not pixmap.isNull():
                    pixmap = pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    self.photo_label.setPixmap(pixmap)
                    self.photo_path = student[9]
                else:
                    self.show_default_photo()
                    self.photo_path = ""
            else:
                self.show_default_photo()
                self.photo_path = ""
            
            photo_layout.addWidget(self.photo_label)
            
            photo_btn = QPushButton("تغییر عکس")
            photo_btn.clicked.connect(lambda: self.select_photo(self.photo_label))
            photo_layout.addWidget(photo_btn)
            
            remove_photo_btn = QPushButton("حذف عکس")
            remove_photo_btn.clicked.connect(lambda: self.remove_photo(self.photo_label))
            photo_layout.addWidget(remove_photo_btn)
            
            form_layout.addRow("عکس دانشجو:", photo_layout)
            layout.addLayout(form_layout)
            
            buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addWidget(buttons)
            
            dialog.setLayout(layout)
            
            if dialog.exec_() == QDialog.Accepted:
                try:
                    first_name = first_name_edit.text().strip()
                    last_name = last_name_edit.text().strip()
                    midterm = midterm_edit.value() if midterm_edit.value() > 0 else None
                    final = final_edit.value() if final_edit.value() > 0 else None
                    average = average_edit.value() if average_edit.value() > 0 else None
                    
                    if not first_name or not last_name:
                        QMessageBox.warning(self, "خطا", "لطفاً تمام فیلدهای ضروری را پر کنید")
                        return
                    
                    self.cursor.execute(
                        "UPDATE students SET first_name=?, last_name=?, midterm=?, final=?, average=?, photo_path=? WHERE student_id=?",
                        (first_name, last_name, midterm, final, average, self.photo_path, student_id)
                    )
                    self.conn.commit()
                    
                    QMessageBox.information(self, "موفقیت", "اطلاعات دانشجو با موفقیت به‌روزرسانی شد")
                    self.load_students()
                    
                except ValueError:
                    QMessageBox.warning(self, "خطا", "لطفاً نمرات را به صورت عددی وارد کنید")
                except Exception as e:
                    QMessageBox.critical(self, "خطا", f"خطا در ویرایش دانشجو: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ویرایش دانشجو: {str(e)}")
    
    def remove_photo(self, label):
        self.show_default_photo()
        self.photo_path = ""
    
    def delete_student(self):
        selected_items = self.student_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "هشدار", "لطفاً یک دانشجو را انتخاب کنید")
            return
        
        row = selected_items[0].row()
        student_id = self.student_table.item(row, 3).text()
        student_name = f"{self.student_table.item(row, 1).text()} {self.student_table.item(row, 2).text()}"
        
        reply = QMessageBox.question(
            self, "تایید حذف", 
            f"آیا از حذف دانشجو {student_name} اطمینان دارید؟\nاین عمل غیرقابل بازگشت است.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                self.cursor.execute("DELETE FROM students WHERE student_id=?", (student_id,))
                self.conn.commit()
                QMessageBox.information(self, "موفقیت", "دانشجو با موفقیت حذف شد")
                self.load_students()
            except Exception as e:
                QMessageBox.critical(self, "خطا", f"خطا در حذف دانشجو: {str(e)}")
    
    def rank_students(self):
        try:
            self.cursor.execute("SELECT * FROM students WHERE average IS NOT NULL ORDER BY average DESC")
            students = self.cursor.fetchall()
            
            if not students:
                QMessageBox.information(self, "اطلاع", "هیچ دانشجویی با معدل معتبر یافت نشد")
                return
            
            for rank, student in enumerate(students, start=1):
                student_id = student[3]
                self.cursor.execute("UPDATE students SET rank=? WHERE student_id=?", (rank, student_id))
            self.conn.commit()
            
            QMessageBox.information(self, "موفقیت", "رتبه‌بندی دانشجوها با موفقیت انجام شد")
            self.load_students()
            self.update_dashboard()
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در رتبه‌بندی: {str(e)}")
    
    def show_student_details(self, index):
        row = index.row()
        student_id = self.student_table.item(row, 3).text()
        
        try:
            self.cursor.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
            student = self.cursor.fetchone()
            
            if not student:
                return
            
            if len(student) < 10:
                self.cursor.execute("SELECT * FROM students WHERE student_id=?", (student_id,))
                student = self.cursor.fetchone()
                if not student or len(student) < 10:
                    QMessageBox.critical(self, "خطا", "داده‌های دانشجو ناقص هستند")
                    return
            
            dialog = QDialog(self)
            dialog.setWindowTitle(f"جزئیات دانشجو: {student[1]} {student[2]}")
            dialog.setMinimumWidth(500)
            dialog.setLayoutDirection(Qt.RightToLeft)
            
            layout = QVBoxLayout()
            
            info_group = QGroupBox("اطلاعات دانشجو")
            info_group.setLayoutDirection(Qt.RightToLeft)
            info_layout = QFormLayout()
            
            info_layout.addRow("نام:", QLabel(student[1]))
            info_layout.addRow("نام خانوادگی:", QLabel(student[2]))
            info_layout.addRow("شماره دانشجویی:", QLabel(student[3]))
            info_layout.addRow("نمره میانترم:", QLabel(str(student[4]) if student[4] is not None else "ثبت نشده"))
            info_layout.addRow("نمره پایان‌ترم:", QLabel(str(student[5]) if student[5] is not None else "ثبت نشده"))
            info_layout.addRow("معدل:", QLabel(str(student[6]) if student[6] is not None else "محاسبه نشده"))
            info_layout.addRow("رتبه:", QLabel(str(student[7]) if student[7] is not None else "رتبه‌بندی نشده"))
            info_layout.addRow("تاریخ ثبت:", QLabel(student[8] if len(student) > 8 else "ثبت نشده"))
            
            info_group.setLayout(info_layout)
            layout.addWidget(info_group)
            
            photo_group = QGroupBox("عکس دانشجو")
            photo_group.setLayoutDirection(Qt.RightToLeft)
            photo_layout = QVBoxLayout()
            
            photo_label = QLabel()
            photo_label.setAlignment(Qt.AlignCenter)
            
            if len(student) > 9 and student[9]:
                pixmap = QPixmap(student[9])
                if not pixmap.isNull():
                    pixmap = pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    photo_label.setPixmap(pixmap)
                else:
                    photo_label.setText("عکسی برای این دانشجو ثبت نشده است")
            else:
                photo_label.setText("عکسی برای این دانشجو ثبت نشده است")
            
            photo_layout.addWidget(photo_label)
            photo_group.setLayout(photo_layout)
            layout.addWidget(photo_group)
            
            close_btn = QPushButton("بستن")
            close_btn.clicked.connect(dialog.accept)
            layout.addWidget(close_btn)
            
            dialog.setLayout(layout)
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در نمایش جزئیات دانشجو: {str(e)}")
    
    def export_to_excel(self):
        try:
            self.cursor.execute("SELECT * FROM students ORDER BY id")
            students = self.cursor.fetchall()
            
            if not students:
                QMessageBox.information(self, "اطلاع", "هیچ دانشجویی برای خروجی وجود ندارد")
                return
            
            if students and len(students[0]) < 10:
                self.create_database()
                self.cursor.execute("SELECT * FROM students ORDER BY id")
                students = self.cursor.fetchall()
            
            if len(students[0]) == 10:
                df = pd.DataFrame(students, columns=[
                    "ردیف", "نام", "نام خانوادگی", "شماره دانشجویی", 
                    "نمره میانترم", "نمره پایان‌ترم", "معدل", "رتبه",
                    "تاریخ ثبت", "مسیر عکس"
                ])
                df = df.drop(columns=["مسیر عکس"])
            else:
                df = pd.DataFrame(students, columns=[
                    "ردیف", "نام", "نام خانوادگی", "شماره دانشجویی", 
                    "نمره میانترم", "نمره پایان‌ترم", "معدل", "رتبه"
                ])
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path, _ = QFileDialog.getSaveFileName(
                self, "ذخیره فایل اکسل", f"students_list_{timestamp}.xlsx", "Excel files (*.xlsx)"
            )
            
            if not file_path:
                return
            
            progress_dialog = QDialog(self)
            progress_dialog.setWindowTitle("در حال ایجاد فایل اکسل")
            progress_dialog.setMinimumWidth(300)
            progress_dialog.setLayoutDirection(Qt.RightToLeft)
            
            layout = QVBoxLayout()
            progress_label = QLabel("در حال ایجاد فایل اکسل...")
            layout.addWidget(progress_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            layout.addWidget(progress_bar)
            progress_dialog.setLayout(layout)
            progress_dialog.show()
            
            df.to_excel(file_path, index=False, engine='openpyxl')
            progress_dialog.close()
            
            QMessageBox.information(self, "موفقیت", f"فایل اکسل با موفقیت ایجاد شد:\n{file_path}")
            
            reply = QMessageBox.question(self, "باز کردن فایل", "آیا مایلید فایل اکسل باز شود؟", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                os.startfile(file_path)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ایجاد فایل اکسل: {str(e)}")
    
    def import_from_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل", "", "Excel files (*.xlsx)")
        if not file_path:
            return
        
        try:
            progress_dialog = QDialog(self)
            progress_dialog.setWindowTitle("در حال وارد کردن داده‌ها")
            progress_dialog.setMinimumWidth(300)
            progress_dialog.setLayoutDirection(Qt.RightToLeft)
            layout = QVBoxLayout()
            progress_label = QLabel("در حال خواندن فایل اکسل...")
            layout.addWidget(progress_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            layout.addWidget(progress_bar)
            progress_dialog.setLayout(layout)
            progress_dialog.show()
            
            df = pd.read_excel(file_path)
            
            required_columns = ["نام", "نام خانوادگی", "شماره دانشجویی"]
            for col in required_columns:
                if col not in df.columns:
                    progress_dialog.close()
                    QMessageBox.critical(self, "خطا", f"ستون '{col}' در فایل اکسل یافت نشد")
                    return
            
            success_count = 0
            error_count = 0
            errors = []
            
            self.create_database()
            
            for index, row in df.iterrows():
                try:
                    first_name = str(row["نام"]).strip()
                    last_name = str(row["نام خانوادگی"]).strip()
                    student_id = str(row["شماره دانشجویی"]).strip()
                    
                    if not first_name or not last_name or not student_id:
                        errors.append(f"ردیف {index+1}: اطلاعات ناقص")
                        error_count += 1
                        continue
                    
                    midterm = None
                    final = None
                    
                    if "نمره میانترم" in row and pd.notna(row["نمره میانترم"]):
                        midterm = float(row["نمره میانترم"])
                    if "نمره پایان‌ترم" in row and pd.notna(row["نمره پایان‌ترم"]):
                        final = float(row["نمره پایان‌ترم"])
                        
                    average = None
                    if midterm is not None and final is not None:
                        average = (midterm * 0.3) + (final * 0.7)
                    
                    registration_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    self.cursor.execute(
                        "INSERT INTO students (first_name, last_name, student_id, midterm, final, average, registration_date) VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (first_name, last_name, student_id, midterm, final, average, registration_date)
                    )
                    self.conn.commit()
                    success_count += 1
                except sqlite3.IntegrityError:
                    errors.append(f"ردیف {index+1}: شماره دانشجویی تکراری")
                    error_count += 1
                except Exception as e:
                    errors.append(f"ردیف {index+1}: خطا - {str(e)}")
                    error_count += 1
            
            progress_dialog.close()
            
            result_msg = f"وارد کردن اطلاعات از اکسل تکمیل شد:\n\n✅ موفق: {success_count}\n❌ خطا: {error_count}\n"
            if errors:
                result_msg += "\nخطاهای رخ داده:\n" + "\n".join(errors[:10])
            
            QMessageBox.information(self, "نتیجه وارد کردن", result_msg)
            self.load_students()
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در خواندن فایل اکسل: {str(e)}")
    
    def backup_database(self):
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره پشتیبان", f"students_backup_{timestamp}.db", "Database files (*.db)")
            if not file_path:
                return
            
            progress_dialog = QDialog(self)
            progress_dialog.setWindowTitle("در حال پشتیبان‌گیری")
            progress_dialog.setMinimumWidth(300)
            progress_dialog.setLayoutDirection(Qt.RightToLeft)
            layout = QVBoxLayout()
            progress_label = QLabel("در حال پشتیبان‌گیری...")
            layout.addWidget(progress_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            layout.addWidget(progress_bar)
            progress_dialog.setLayout(layout)
            progress_dialog.show()
            
            with open('students.db', 'rb') as src, open(file_path, 'wb') as dst:
                dst.write(src.read())
            
            progress_dialog.close()
            QMessageBox.information(self, "موفقیت", f"پشتیبان‌گیری با موفقیت انجام شد:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در پشتیبان‌گیری: {str(e)}")
    
    def restore_database(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل پشتیبان", "", "Database files (*.db)")
        if not file_path:
            return
        
        reply = QMessageBox.question(self, "تایید بازیابی", "آیا از بازیابی پشتیبان اطمینان دارید؟\nتمامی داده‌های فعلی جایگزین خواهند شد.", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return
        
        try:
            progress_dialog = QDialog(self)
            progress_dialog.setWindowTitle("در حال بازیابی پشتیبان")
            progress_dialog.setMinimumWidth(300)
            progress_dialog.setLayoutDirection(Qt.RightToLeft)
            layout = QVBoxLayout()
            progress_label = QLabel("در حال بازیابی...")
            layout.addWidget(progress_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            layout.addWidget(progress_bar)
            progress_dialog.setLayout(layout)
            progress_dialog.show()
            
            self.conn.close()
            with open(file_path, 'rb') as src, open('students.db', 'wb') as dst:
                dst.write(src.read())
            
            self.conn = sqlite3.connect('students.db')
            self.cursor = self.conn.cursor()
            
            progress_dialog.close()
            QMessageBox.information(self, "موفقیت", "بازیابی پشتیبان با موفقیت انجام شد")
            self.load_students()
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در بازیابی پشتیبان: {str(e)}")
            try:
                self.conn = sqlite3.connect('students.db')
                self.cursor = self.conn.cursor()
            except:
                pass
    
    def show_statistics_chart(self):
        try:
            self.cursor.execute("SELECT average FROM students WHERE average IS NOT NULL")
            averages = [row[0] for row in self.cursor.fetchall()]
            
            if not averages:
                QMessageBox.information(self, "اطلاع", "هیچ داده‌ای برای نمایش نمودار وجود ندارد")
                return
            
            dialog = QDialog(self)
            dialog.setWindowTitle("نمودار آماری معدل دانشجویان")
            dialog.setMinimumSize(800, 600)
            dialog.setLayoutDirection(Qt.RightToLeft)
            
            layout = QVBoxLayout()
            
            figure = Figure(figsize=(10, 6), dpi=100)
            canvas = FigureCanvas(figure)
            layout.addWidget(canvas)
            
            ax = figure.add_subplot(111)
            ax.hist(averages, bins=20, color='#3498db', edgecolor='black', alpha=0.7)
            ax.set_title('توزیع معدل دانشجویان', fontname='B Nazanin', fontsize=16)
            ax.set_xlabel('معدل', fontname='B Nazanin', fontsize=12)
            ax.set_ylabel('تعداد دانشجویان', fontname='B Nazanin', fontsize=12)
            ax.grid(True, linestyle='--', alpha=0.7)
            
            close_btn = QPushButton("بستن")
            close_btn.clicked.connect(dialog.accept)
            layout.addWidget(close_btn)
            
            dialog.setLayout(layout)
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در نمایش نمودار: {str(e)}")
    
    def print_report(self):
        try:
            self.cursor.execute("SELECT * FROM students ORDER BY id")
            students = self.cursor.fetchall()
            
            if not students:
                QMessageBox.information(self, "اطلاع", "هیچ دانشجویی برای چاپ وجود ندارد")
                return
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره گزارش", f"students_report_{timestamp}.txt", "Text files (*.txt)")
            if not file_path:
                return
            
            progress_dialog = QDialog(self)
            progress_dialog.setWindowTitle("در حال ایجاد گزارش")
            progress_dialog.setMinimumWidth(300)
            progress_dialog.setLayoutDirection(Qt.RightToLeft)
            layout = QVBoxLayout()
            progress_label = QLabel("در حال ایجاد گزارش...")
            layout.addWidget(progress_label)
            progress_bar = QProgressBar()
            progress_bar.setRange(0, 0)
            layout.addWidget(progress_bar)
            progress_dialog.setLayout(layout)
            progress_dialog.show()
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("="*50 + "\n")
                f.write("گزارش دانشجویان".center(50) + "\n")
                f.write("="*50 + "\n\n")
                f.write(f"تاریخ گزارش: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write(f"{'ردیف':<5}{'نام':<15}{'نام خانوادگی':<20}{'شماره دانشجویی':<15}{'میانترم':<10}{'پایان‌ترم':<10}{'معدل':<10}{'رتبه':<10}\n")
                f.write("-"*95 + "\n")
                
                for student in students:
                    f.write(f"{student[0]:<5}{student[1]:<15}{student[2]:<20}{student[3]:<15}{str(student[4]) if student[4] is not None else '-':<10}{str(student[5]) if student[5] is not None else '-':<10}{str(student[6]) if student[6] is not None else '-':<10}{str(student[7]) if student[7] is not None else '-':<10}\n")
                
                f.write("\n" + "="*50 + "\n")
                f.write("آمار پایانی".center(50) + "\n")
                f.write("="*50 + "\n\n")
                
                total_students = len(students)
                avg_list = [s[6] for s in students if s[6] is not None]
                avg_grade = sum(avg_list) / len(avg_list) if avg_list else 0
                max_grade = max(avg_list) if avg_list else 0
                excellent_count = len([s for s in students if s[6] is not None and s[6] >= 17])
                
                f.write(f"تعداد کل دانشجویان: {total_students}\n")
                f.write(f"میانگین معدل: {avg_grade:.2f}\n")
                f.write(f"بالاترین معدل: {max_grade:.2f}\n")
                f.write(f"تعداد دانشجویان ممتاز: {excellent_count}\n")
            
            progress_dialog.close()
            QMessageBox.information(self, "موفقیت", f"گزارش با موفقیت ایجاد شد:\n{file_path}")
            
            reply = QMessageBox.question(self, "باز کردن فایل", "آیا مایلید فایل گزارش باز شود؟", QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                os.startfile(file_path)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ایجاد گزارش: {str(e)}")
    
    def show_about(self):
        about_text = """
        <h2>مدیریت دانشجویان</h2>
        <p>نسخه: 2.0.0</p>
        <p>یک برنامه ساده و کاربردی برای مدیریت اطلاعات دانشجویان</p>
        """
        msg = QMessageBox(self)
        msg.setWindowTitle("درباره برنامه")
        msg.setTextFormat(Qt.RichText)
        msg.setLayoutDirection(Qt.RightToLeft)
        msg.setText(about_text)
        msg.exec_()
    
    def update_time(self):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_label.setText(now)
    
    def closeEvent(self, event):
        try:
            if hasattr(self, 'conn') and self.conn:
                self.conn.close()
        except:
            pass
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(QFont("B Nazanin", 10))
    
    window = StudentManagementSystem()
    window.show()
    sys.exit(app.exec_())
