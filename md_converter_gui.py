import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Markdown 批量转换工具")
        self.setGeometry(500, 300, 400, 250)
        self.docs_dir = ""

        layout = QVBoxLayout()

        self.label = QLabel("请选择包含 Markdown 文件的文件夹：")
        layout.addWidget(self.label)

        self.path_label = QLabel("📂 当前未选择目录")
        self.path_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.path_label)

        btn_select = QPushButton("选择文件夹")
        btn_select.clicked.connect(self.select_folder)
        layout.addWidget(btn_select)

        self.btn_word = QPushButton("转换为 Word（.docx）")
        self.btn_word.clicked.connect(self.convert_to_word)
        self.btn_word.setEnabled(False)
        layout.addWidget(self.btn_word)

        self.btn_pdf = QPushButton("转换为 PDF")
        self.btn_pdf.clicked.connect(self.convert_to_pdf)
        self.btn_pdf.setEnabled(False)
        layout.addWidget(self.btn_pdf)

        self.setLayout(layout)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择 Markdown 文件夹")
        if folder:
            self.docs_dir = folder
            self.path_label.setText(f"📁 已选择目录：{folder}")
            self.btn_word.setEnabled(True)
            self.btn_pdf.setEnabled(True)

    def convert_to_word(self):
        if not self.docs_dir:
            QMessageBox.warning(self, "提示", "请先选择文件夹！")
            return
        try:
            script_path = os.path.join(os.path.dirname(__file__), "merge_md_to_docx.py")
            subprocess.run(["python", script_path, self.docs_dir], check=True)
            QMessageBox.information(self, "成功", "✅ Word 文件生成完成！")
        except subprocess.CalledProcessError:
            QMessageBox.critical(self, "错误", "❌ Word 转换失败，请检查日志。")

    def convert_to_pdf(self):
        if not self.docs_dir:
            QMessageBox.warning(self, "提示", "请先选择文件夹！")
            return
        try:
            script_path = os.path.join(os.path.dirname(__file__), "to_pdf.py")
            subprocess.run(["python", script_path, self.docs_dir], check=True)
            QMessageBox.information(self, "成功", "✅ PDF 文件生成完成！")
        except subprocess.CalledProcessError:
            QMessageBox.critical(self, "错误", "❌ PDF 转换失败，请检查日志。")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    sys.exit(app.exec_())
