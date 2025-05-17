import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox
)
import sys


class ExcelMerger(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("دمج ملفات جدول الإجازات")
        self.setGeometry(100, 100, 400, 200)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label_folder = QLabel("لم يتم اختيار مجلد")
        self.label_output = QLabel("لم يتم تحديد اسم الملف")

        self.btn_select_folder = QPushButton("اختر مجلد ملفات Excel")
        self.btn_select_folder.clicked.connect(self.select_folder)

        self.btn_select_output = QPushButton("اختر اسم ملف Excel الناتج")
        self.btn_select_output.clicked.connect(self.select_output_file)

        self.btn_merge = QPushButton("دمج الملفات")
        self.btn_merge.clicked.connect(self.merge_excels)

        layout.addWidget(self.label_folder)
        layout.addWidget(self.btn_select_folder)
        layout.addWidget(self.label_output)
        layout.addWidget(self.btn_select_output)
        layout.addWidget(self.btn_merge)

        self.setLayout(layout)

        self.folder_path = ""
        self.output_file_path = ""
        self.columns_order = [
            "الرقم الاحصائي", "الرتبة", "الاسم", "المديرية",
            "مدة الاجازة", "الدولة", "المادة القانونية",
            "سبب الاجازة", "رقم الكتاب", "تاريخ الكتاب"
        ]

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "اختر المجلد")
        if folder:
            self.folder_path = folder
            self.label_folder.setText(f"المجلد المختار: {folder}")

    def select_output_file(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "اختر اسم الملف", "", "Excel Files (*.xlsx)")
        if file_path:
            self.output_file_path = file_path
            self.label_output.setText(f"اسم الملف الناتج: {file_path}")

    def merge_excels(self):
        if not self.folder_path or not self.output_file_path:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار المجلد واسم الملف أولاً")
            return

        all_data = []
        for file in os.listdir(self.folder_path):
            if file.endswith('.xlsx'):
                file_path = os.path.join(self.folder_path, file)
                try:
                    df = pd.read_excel(file_path)

                    # تنظيف أسماء الأعمدة
                    df.columns = [col.strip() for col in df.columns]

                    # ترتيب الأعمدة حسب النسق
                    df = df.reindex(columns=self.columns_order)

                    all_data.append(df)
                except Exception as e:
                    QMessageBox.warning(self, "خطأ", f"حدث خطأ أثناء قراءة الملف: {file}\n{str(e)}")
                    return

        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            try:
                merged_df.to_excel(self.output_file_path, index=False)
                QMessageBox.information(self, "نجاح", "تم دمج الملفات وإعادة ترتيب الأعمدة بنجاح")
            except Exception as e:
                QMessageBox.critical(self, "خطأ", f"فشل في حفظ الملف\n{str(e)}")
        else:
            QMessageBox.warning(self, "تحذير", "لا توجد ملفات Excel صالحة في المجلد")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelMerger()
    window.show()
    sys.exit(app.exec_())
