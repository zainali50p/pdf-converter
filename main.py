# main.py
import sys
import os
import io
from PyQt5 import QtWidgets, QtGui, QtCore
from pdf2docx import Converter
import mammoth
from xhtml2pdf import pisa

APP_TITLE = "File Converter — PDF ⇆ DOCX"

NEON_STYLESHEET = """
QWidget {
    background: #0b0f1a;
    color: #e6f0ff;
    font-family: "Segoe UI", Roboto, "Helvetica Neue", Arial;
    font-size: 11pt;
}
QPushButton {
    background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #00ffa3, stop:1 #7a00ff);
    border-radius: 10px;
    padding: 10px;
    color: white;
    font-weight: 600;
    min-width: 160px;
}
QPushButton:hover {
    box-shadow: 0 0 20px rgba(122,0,255,0.55);
}
QPushButton:pressed {
    padding-left: 12px;
}
QLabel#title {
    font-size: 16pt;
    font-weight: 700;
    color: #00ffa3;
}
QFrame#card {
    background: #071024;
    border-radius: 14px;
    border: 1px solid rgba(0,255,163,0.06);
    padding: 18px;
}
QProgressBar {
    background: #06101a;
    border: 1px solid #12202b;
    border-radius: 8px;
    text-align: center;
    padding: 3px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00ffa3, stop:1 #7a00ff);
    border-radius: 8px;
}
QLineEdit {
    background: #081121;
    border: 1px solid #12202b;
    padding: 8px;
    border-radius: 8px;
    color: #cfeffd;
}
"""

class Worker(QtCore.QThread):
    progress = QtCore.pyqtSignal(int)           # percent
    finished = QtCore.pyqtSignal(bool, str)     # success, message

    def __init__(self, mode, input_path, output_path=None):
        super().__init__()
        self.mode = mode
        self.input_path = input_path
        self.output_path = output_path

    def run(self):
        try:
            if self.mode == "pdf_to_docx":
                self._pdf_to_docx(self.input_path, self.output_path)
            elif self.mode == "docx_to_pdf":
                self._docx_to_pdf(self.input_path, self.output_path)
            else:
                raise RuntimeError("Unknown mode")
            self.finished.emit(True, "Conversion completed successfully.")
        except Exception as e:
            self.finished.emit(False, f"Conversion failed: {str(e)}")

    def _pdf_to_docx(self, pdf_path, docx_path):
        # pdf2docx conversion (chunked progress)
        cv = Converter(pdf_path)
        # pages info - attempt to report progress
        try:
            total_pages = cv.page_count
        except Exception:
            total_pages = None
        def progress_callback(page, pages_total=None):
            if pages_total:
                pct = int(page / pages_total * 100)
            elif total_pages:
                pct = int(page / total_pages * 100)
            else:
                pct = min(99, page % 100)
            self.progress.emit(pct)
        cv.convert(docx_path, start=0, end=None, callback=progress_callback)
        cv.close()
        self.progress.emit(100)

    def _docx_to_pdf(self, docx_path, pdf_path):
        # Use mammoth to convert DOCX -> HTML, then xhtml2pdf to render PDF.
        with open(docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value  # string
            messages = result.messages  # warnings
        # Add a minimal wrapper to ensure correct rendering
        full_html = f"""
        <html>
        <head>
        <meta charset="utf-8" />
        <style>
            body {{ font-family: Arial, Helvetica, sans-serif; margin: 1cm; color: #000; }}
            img {{ max-width: 100%; height: auto; }}
        </style>
        </head>
        <body>
        {html}
        </body>
        </html>
        """
        # Convert html -> pdf using xhtml2pdf (pisa)
        # xhtml2pdf likes bytes/str IO
        with open(pdf_path, "wb") as out_file:
            # pisa status
            self.progress.emit(20)
            pisa_status = pisa.CreatePDF(io.StringIO(full_html), dest=out_file)
            # pisa_status.err indicates error count
            if pisa_status.err:
                raise RuntimeError(f"xhtml2pdf failed with {pisa_status.err} errors.")
        self.progress.emit(100)


class MainWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setMinimumSize(700, 360)
        self.setWindowIcon(QtGui.QIcon())
        self.setup_ui()

    def setup_ui(self):
        self.setStyleSheet(NEON_STYLESHEET)
        layout = QtWidgets.QVBoxLayout(self)
        header = QtWidgets.QLabel("File Converter — PDF ⇆ DOCX")
        header.setObjectName("title")
        header.setAlignment(QtCore.Qt.AlignCenter)

        card = QtWidgets.QFrame()
        card.setObjectName("card")
        card_layout = QtWidgets.QGridLayout(card)

        # PDF -> DOCX controls
        self.pdf_input = QtWidgets.QLineEdit()
        self.pdf_input.setPlaceholderText("Choose PDF file...")
        btn_pdf_browse = QtWidgets.QPushButton("Browse PDF")
        btn_pdf_browse.clicked.connect(self.browse_pdf)
        btn_pdf_convert = QtWidgets.QPushButton("Convert PDF → DOCX")
        btn_pdf_convert.clicked.connect(self.convert_pdf_to_docx)

        # DOCX -> PDF controls
        self.docx_input = QtWidgets.QLineEdit()
        self.docx_input.setPlaceholderText("Choose DOCX file...")
        btn_docx_browse = QtWidgets.QPushButton("Browse DOCX")
        btn_docx_browse.clicked.connect(self.browse_docx)
        btn_docx_convert = QtWidgets.QPushButton("Convert DOCX → PDF")
        btn_docx_convert.clicked.connect(self.convert_docx_to_pdf)

        # Progress and log
        self.progress = QtWidgets.QProgressBar()
        self.progress.setValue(0)
        self.log = QtWidgets.QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(120)

        # Layout placements
        card_layout.addWidget(QtWidgets.QLabel("PDF → DOCX"), 0, 0)
        card_layout.addWidget(self.pdf_input, 0, 1)
        card_layout.addWidget(btn_pdf_browse, 0, 2)
        card_layout.addWidget(btn_pdf_convert, 0, 3)

        card_layout.addWidget(QtWidgets.QLabel("DOCX → PDF"), 1, 0)
        card_layout.addWidget(self.docx_input, 1, 1)
        card_layout.addWidget(btn_docx_browse, 1, 2)
        card_layout.addWidget(btn_docx_convert, 1, 3)

        # Stretch columns
        card_layout.setColumnStretch(1, 1)
        card_layout.setColumnMinimumWidth(3, 160)

        layout.addWidget(header)
        layout.addWidget(card)
        layout.addSpacing(8)
        layout.addWidget(self.progress)
        layout.addWidget(self.log)

    def browse_pdf(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PDF file", "", "PDF Files (*.pdf)")
        if path:
            self.pdf_input.setText(path)

    def browse_docx(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select DOCX file", "", "Word Files (*.docx)")
        if path:
            self.docx_input.setText(path)

    def convert_pdf_to_docx(self):
        pdf_path = self.pdf_input.text().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            QtWidgets.QMessageBox.warning(self, "Missing file", "Select a valid PDF file first.")
            return
        # default output filename
        out_default = os.path.splitext(pdf_path)[0] + ".docx"
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save DOCX as", out_default, "DOCX Files (*.docx)")
        if not save_path:
            return
        self.log.append(f"Starting PDF → DOCX: {os.path.basename(pdf_path)} → {os.path.basename(save_path)}")
        self.start_worker("pdf_to_docx", pdf_path, save_path)

    def convert_docx_to_pdf(self):
        docx_path = self.docx_input.text().strip()
        if not docx_path or not os.path.isfile(docx_path):
            QtWidgets.QMessageBox.warning(self, "Missing file", "Select a valid DOCX file first.")
            return
        out_default = os.path.splitext(docx_path)[0] + ".pdf"
        save_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save PDF as", out_default, "PDF Files (*.pdf)")
        if not save_path:
            return
        self.log.append(f"Starting DOCX → PDF: {os.path.basename(docx_path)} → {os.path.basename(save_path)}")
        self.start_worker("docx_to_pdf", docx_path, save_path)

    def start_worker(self, mode, input_path, output_path):
        self.progress.setValue(0)
        self.worker = Worker(mode, input_path, output_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def on_progress(self, pct):
        self.progress.setValue(pct)

    def on_finished(self, success, message):
        self.progress.setValue(100 if success else 0)
        if success:
            self.log.append("✅ " + message)
            QtWidgets.QMessageBox.information(self, "Done", message)
        else:
            self.log.append("❌ " + message)
            QtWidgets.QMessageBox.critical(self, "Error", message)

def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
