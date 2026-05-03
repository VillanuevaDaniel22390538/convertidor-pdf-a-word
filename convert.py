import os
import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QLineEdit, QHBoxLayout, QProgressBar
)
from pdf2docx import Converter
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
import pytesseract
from docx import Document
from docx2pdf import convert as docx2pdf_convert

OCR_LANG = "spa"

# -------------------------------
# Funciones de conversión
# -------------------------------

def pdf_has_text(pdf_path):
    reader = PdfReader(pdf_path)
    for page in reader.pages:
        if page.extract_text():
            return True
    return False

def convert_with_pdf2docx(pdf_path, word_path):
    cv = Converter(pdf_path)
    cv.convert(word_path, start=0, end=None)
    cv.close()

def convert_with_ocr(pdf_path, word_path):
    pages = convert_from_path(pdf_path)
    doc = Document()
    for page_num, page in enumerate(pages, start=1):
        text = pytesseract.image_to_string(page, lang=OCR_LANG)
        doc.add_paragraph(f"--- Página {page_num} ---")
        doc.add_paragraph(text if text.strip() else "[Sin texto detectado]")
    doc.save(word_path)

def convert_word_to_pdf(word_path, pdf_path=None):
    # Si no se especifica pdf_path, guarda en la misma carpeta
    docx2pdf_convert(word_path, pdf_path)

# -------------------------------
# Interfaz gráfica
# -------------------------------

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Convert - PDF ↔ Word Converter (Versión 2.0)")
        self.setGeometry(200, 200, 650, 450)

        layout = QVBoxLayout()

        title = QLabel("Convert - PDF ↔ Word")
        title.setStyleSheet("font-size: 22px; font-weight: bold; color: darkblue;")
        layout.addWidget(title)

        # Input folder
        hbox1 = QHBoxLayout()
        self.input_edit = QLineEdit()
        btn_input = QPushButton("Seleccionar carpeta de entrada")
        btn_input.clicked.connect(self.select_input)
        hbox1.addWidget(QLabel("Carpeta de entrada:"))
        hbox1.addWidget(self.input_edit)
        hbox1.addWidget(btn_input)
        layout.addLayout(hbox1)

        # Output folder
        hbox2 = QHBoxLayout()
        self.output_edit = QLineEdit()
        btn_output = QPushButton("Seleccionar carpeta de salida")
        btn_output.clicked.connect(self.select_output)
        hbox2.addWidget(QLabel("Carpeta de salida:"))
        hbox2.addWidget(self.output_edit)
        hbox2.addWidget(btn_output)
        layout.addLayout(hbox2)

        # Progress bar
        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        # Buttons
        hbox3 = QHBoxLayout()
        btn_pdf_to_word = QPushButton("PDF → Word")
        btn_pdf_to_word.setStyleSheet("background-color: green; color: white; font-weight: bold;")
        btn_pdf_to_word.clicked.connect(self.run_pdf_to_word)

        btn_word_to_pdf = QPushButton("Word → PDF")
        btn_word_to_pdf.setStyleSheet("background-color: purple; color: white; font-weight: bold;")
        btn_word_to_pdf.clicked.connect(self.run_word_to_pdf)

        btn_help = QPushButton("Ayuda")
        btn_help.setStyleSheet("background-color: gold;")
        btn_help.clicked.connect(self.show_help)

        btn_about = QPushButton("Acerca de")
        btn_about.setStyleSheet("background-color: steelblue; color: white;")
        btn_about.clicked.connect(self.show_about)

        hbox3.addWidget(btn_pdf_to_word)
        hbox3.addWidget(btn_word_to_pdf)
        hbox3.addWidget(btn_help)
        hbox3.addWidget(btn_about)
        layout.addLayout(hbox3)

        self.setLayout(layout)

    def select_input(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecciona carpeta de entrada")
        if folder:
            self.input_edit.setText(folder)

    def select_output(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecciona carpeta de salida")
        if folder:
            self.output_edit.setText(folder)

    def run_pdf_to_word(self):
        input_folder = self.input_edit.text()
        output_folder = self.output_edit.text()
        if not input_folder or not output_folder:
            QMessageBox.critical(self, "Error", "Debes seleccionar ambas carpetas.")
            return

        os.makedirs(output_folder, exist_ok=True)
        files = [f for f in os.listdir(input_folder) if f.endswith(".pdf")]
        total = len(files)
        self.progress.setValue(0)

        for i, file in enumerate(files, start=1):
            pdf_path = os.path.join(input_folder, file)
            word_path = os.path.join(output_folder, file.replace(".pdf", ".docx"))

            if pdf_has_text(pdf_path):
                convert_with_pdf2docx(pdf_path, word_path)
            else:
                convert_with_ocr(pdf_path, word_path)

            self.progress.setValue(int((i / total) * 100))

        QMessageBox.information(self, "Completado", f"Conversión PDF→Word terminada.\nArchivos guardados en:\n{output_folder}")

    def run_word_to_pdf(self):
        input_folder = self.input_edit.text()
        output_folder = self.output_edit.text()
        if not input_folder or not output_folder:
            QMessageBox.critical(self, "Error", "Debes seleccionar ambas carpetas.")
            return

        os.makedirs(output_folder, exist_ok=True)
        files = [f for f in os.listdir(input_folder) if f.endswith(".docx")]
        total = len(files)
        self.progress.setValue(0)

        for i, file in enumerate(files, start=1):
            word_path = os.path.join(input_folder, file)
            pdf_path = os.path.join(output_folder, file.replace(".docx", ".pdf"))

            convert_word_to_pdf(word_path, pdf_path)
            self.progress.setValue(int((i / total) * 100))

        QMessageBox.information(self, "Completado", f"Conversión Word→PDF terminada.\nArchivos guardados en:\n{output_folder}")

    def show_help(self):
        QMessageBox.information(self, "Ayuda",
            "Instrucciones de uso:\n\n"
            "1. Selecciona la carpeta de entrada (PDFs o DOCX).\n"
            "2. Selecciona la carpeta de salida.\n"
            "3. Usa el botón correspondiente (PDF→Word o Word→PDF).\n\n"
            "El programa detecta automáticamente si el PDF tiene texto o es escaneado."
        )

    def show_about(self):
        QMessageBox.information(self, "Acerca de",
            "Convert - Versión 2.0\n\n"
            "Hecho por Daniel Donaldo Villanueva Canul\n"
            "Este programa convierte múltiples PDFs a Word y también Word a PDF."
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConverterApp()
    window.show()
    sys.exit(app.exec_())
