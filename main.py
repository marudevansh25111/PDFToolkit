import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout,QPushButton, QFileDialog, QLabel, QProgressBar, QMessageBox,QListWidget, QHBoxLayout, QSpinBox, QComboBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import PyPDF2
from pdf2docx import Converter
import pikepdf
import os

class PDFWorker(QThread):
    """Worker thread to handle PDF operations in background"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, operation, params):
        super().__init__()
        self.operation = operation
        self.params = params
        
    def run(self):
        try:
            if self.operation == "merge":
                self._merge_pdfs()
            elif self.operation == "split":
                self._split_pdf()
            elif self.operation == "compress":
                self._compress_pdf()
            elif self.operation == "to_word":
                self._pdf_to_word()
                
            self.finished.emit(True, "Operation completed successfully!")
        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")
    
    def _merge_pdfs(self):
        input_files = self.params["input_files"]
        output_file = self.params["output_file"]
        
        merger = PyPDF2.PdfMerger()
        
        total_files = len(input_files)
        for i, file in enumerate(input_files):
            merger.append(file)
            self.progress.emit(int((i + 1) / total_files * 100))
            
        merger.write(output_file)
        merger.close()
    
    def _split_pdf(self):
        input_file = self.params["input_file"]
        output_dir = self.params["output_dir"]
        
        with open(input_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            total_pages = len(reader.pages)
            
            for i in range(total_pages):
                writer = PyPDF2.PdfWriter()
                writer.add_page(reader.pages[i])
                
                output_filename = os.path.join(output_dir, f"page_{i+1}.pdf")
                with open(output_filename, 'wb') as output_file:
                    writer.write(output_file)
                
                self.progress.emit(int((i + 1) / total_pages * 100))
    
    def _compress_pdf(self):
        input_file = self.params["input_file"]
        output_file = self.params["output_file"]
        quality = self.params.get("quality", "medium")
        
        # Quality settings map to different compression levels
        quality_settings = {
            "low": 0.5,      # Higher compression, lower quality
            "medium": 0.75,  # Balanced
            "high": 0.9      # Lower compression, higher quality
        }
        
        # Use pikepdf for compression
        with pikepdf.open(input_file) as pdf:
            # We'll set the progress to 50% for opening
            self.progress.emit(50)
            
            # Save with compression settings
            pdf.save(output_file, 
                    compress_streams=True,
                    preserve_pdfa=False,
                    object_stream_mode=pikepdf.ObjectStreamMode.generate,
                    linearize=False)
            
            self.progress.emit(100)
    
    def _pdf_to_word(self):
        input_file = self.params["input_file"]
        output_file = self.params["output_file"]
        
        # Convert PDF to DOCX
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None)
        cv.close()
        
        self.progress.emit(100)


class PDFToolkitApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Toolkit")
        self.setMinimumSize(800, 600)
        
        # Create tabs for different PDF operations
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Create each operation tab
        self.setup_merge_tab()
        self.setup_split_tab()
        self.setup_compress_tab()
        self.setup_pdf_to_word_tab()
        
        # Initialize worker thread
        self.worker = None
    
    def setup_merge_tab(self):
        merge_tab = QWidget()
        layout = QVBoxLayout()
        
        # File list
        self.merge_file_list = QListWidget()
        layout.addWidget(QLabel("Files to merge:"))
        layout.addWidget(self.merge_file_list)
        
        # Buttons
        button_layout = QHBoxLayout()
        add_button = QPushButton("Add Files")
        add_button.clicked.connect(self.add_files_to_merge)
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self.remove_selected_merge_file)
        
        button_layout.addWidget(add_button)
        button_layout.addWidget(remove_button)
        layout.addLayout(button_layout)
        
        # Output selection
        output_layout = QHBoxLayout()
        self.merge_output_label = QLabel("Output file: Not selected")
        output_button = QPushButton("Select Output")
        output_button.clicked.connect(self.select_merge_output)
        
        output_layout.addWidget(self.merge_output_label)
        output_layout.addWidget(output_button)
        layout.addLayout(output_layout)
        
        # Progress bar and merge button
        self.merge_progress = QProgressBar()
        layout.addWidget(self.merge_progress)
        
        merge_button = QPushButton("Merge PDFs")
        merge_button.clicked.connect(self.merge_pdfs)
        layout.addWidget(merge_button)
        
        merge_tab.setLayout(layout)
        self.tabs.addTab(merge_tab, "Merge PDFs")
    
    def setup_split_tab(self):
        split_tab = QWidget()
        layout = QVBoxLayout()
        
        # File selection
        file_layout = QHBoxLayout()
        self.split_file_label = QLabel("Input file: Not selected")
        select_file_button = QPushButton("Select PDF")
        select_file_button.clicked.connect(self.select_split_input)
        
        file_layout.addWidget(self.split_file_label)
        file_layout.addWidget(select_file_button)
        layout.addLayout(file_layout)
        
        # Output directory
        output_layout = QHBoxLayout()
        self.split_output_label = QLabel("Output directory: Not selected")
        select_output_button = QPushButton("Select Output Directory")
        select_output_button.clicked.connect(self.select_split_output)
        
        output_layout.addWidget(self.split_output_label)
        output_layout.addWidget(select_output_button)
        layout.addLayout(output_layout)
        
        # Progress bar and split button
        self.split_progress = QProgressBar()
        layout.addWidget(self.split_progress)
        
        split_button = QPushButton("Split PDF")
        split_button.clicked.connect(self.split_pdf)
        layout.addWidget(split_button)
        
        split_tab.setLayout(layout)
        self.tabs.addTab(split_tab, "Split PDF")
    
    def setup_compress_tab(self):
        compress_tab = QWidget()
        layout = QVBoxLayout()
        
        # File selection
        file_layout = QHBoxLayout()
        self.compress_file_label = QLabel("Input file: Not selected")
        select_file_button = QPushButton("Select PDF")
        select_file_button.clicked.connect(self.select_compress_input)
        
        file_layout.addWidget(self.compress_file_label)
        file_layout.addWidget(select_file_button)
        layout.addLayout(file_layout)
        
        # Quality selection
        quality_layout = QHBoxLayout()
        quality_layout.addWidget(QLabel("Compression Quality:"))
        self.quality_combo = QComboBox()
        self.quality_combo.addItems(["Low", "Medium", "High"])
        self.quality_combo.setCurrentIndex(1)  # Medium by default
        quality_layout.addWidget(self.quality_combo)
        layout.addLayout(quality_layout)
        
        # Output file
        output_layout = QHBoxLayout()
        self.compress_output_label = QLabel("Output file: Not selected")
        select_output_button = QPushButton("Select Output File")
        select_output_button.clicked.connect(self.select_compress_output)
        
        output_layout.addWidget(self.compress_output_label)
        output_layout.addWidget(select_output_button)
        layout.addLayout(output_layout)
        
        # Progress bar and compress button
        self.compress_progress = QProgressBar()
        layout.addWidget(self.compress_progress)
        
        compress_button = QPushButton("Compress PDF")
        compress_button.clicked.connect(self.compress_pdf)
        layout.addWidget(compress_button)
        
        compress_tab.setLayout(layout)
        self.tabs.addTab(compress_tab, "Compress PDF")
    
    def setup_pdf_to_word_tab(self):
        pdf_to_word_tab = QWidget()
        layout = QVBoxLayout()
        
        # File selection
        file_layout = QHBoxLayout()
        self.pdf_word_file_label = QLabel("Input PDF: Not selected")
        select_file_button = QPushButton("Select PDF")
        select_file_button.clicked.connect(self.select_pdf_to_word_input)
        
        file_layout.addWidget(self.pdf_word_file_label)
        file_layout.addWidget(select_file_button)
        layout.addLayout(file_layout)
        
        # Output file
        output_layout = QHBoxLayout()
        self.pdf_word_output_label = QLabel("Output Word file: Not selected")
        select_output_button = QPushButton("Select Output File")
        select_output_button.clicked.connect(self.select_pdf_to_word_output)
        
        output_layout.addWidget(self.pdf_word_output_label)
        output_layout.addWidget(select_output_button)
        layout.addLayout(output_layout)
        
        # Progress bar and convert button
        self.pdf_word_progress = QProgressBar()
        layout.addWidget(self.pdf_word_progress)
        
        convert_button = QPushButton("Convert to Word")
        convert_button.clicked.connect(self.pdf_to_word)
        layout.addWidget(convert_button)
        
        pdf_to_word_tab.setLayout(layout)
        self.tabs.addTab(pdf_to_word_tab, "PDF to Word")
    
    # Event handlers for Merge tab
    def add_files_to_merge(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF Files", "", "PDF Files (*.pdf)"
        )
        if files:
            self.merge_file_list.addItems(files)
    
    def remove_selected_merge_file(self):
        for item in self.merge_file_list.selectedItems():
            self.merge_file_list.takeItem(self.merge_file_list.row(item))
    
    def select_merge_output(self):
        file, _ = QFileDialog.getSaveFileName(
            self, "Save Merged PDF", "", "PDF Files (*.pdf)"
        )
        if file:
            self.merge_output_label.setText(f"Output file: {file}")
            self.merge_output_file = file
    
    def merge_pdfs(self):
        # Get all files from list
        files = []
        for i in range(self.merge_file_list.count()):
            files.append(self.merge_file_list.item(i).text())
        
        if not files:
            QMessageBox.warning(self, "Warning", "Please add PDF files to merge")
            return
        
        if not hasattr(self, "merge_output_file"):
            QMessageBox.warning(self, "Warning", "Please select an output file")
            return
        
        # Start the worker thread
        params = {
            "input_files": files,
            "output_file": self.merge_output_file
        }
        self.start_worker("merge", params, self.merge_progress)
    
    # Event handlers for Split tab
    def select_split_input(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)"
        )
        if file:
            self.split_file_label.setText(f"Input file: {file}")
            self.split_input_file = file
    
    def select_split_output(self):
        directory = QFileDialog.getExistingDirectory(
            self, "Select Output Directory"
        )
        if directory:
            self.split_output_label.setText(f"Output directory: {directory}")
            self.split_output_dir = directory
    
    def split_pdf(self):
        if not hasattr(self, "split_input_file"):
            QMessageBox.warning(self, "Warning", "Please select an input file")
            return
        
        if not hasattr(self, "split_output_dir"):
            QMessageBox.warning(self, "Warning", "Please select an output directory")
            return
        
        # Start the worker thread
        params = {
            "input_file": self.split_input_file,
            "output_dir": self.split_output_dir
        }
        self.start_worker("split", params, self.split_progress)
    
    # Event handlers for Compress tab
    def select_compress_input(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)"
        )
        if file:
            self.compress_file_label.setText(f"Input file: {file}")
            self.compress_input_file = file
    
    def select_compress_output(self):
        file, _ = QFileDialog.getSaveFileName(
            self, "Save Compressed PDF", "", "PDF Files (*.pdf)"
        )
        if file:
            self.compress_output_label.setText(f"Output file: {file}")
            self.compress_output_file = file
    
    def compress_pdf(self):
        if not hasattr(self, "compress_input_file"):
            QMessageBox.warning(self, "Warning", "Please select an input file")
            return
        
        if not hasattr(self, "compress_output_file"):
            QMessageBox.warning(self, "Warning", "Please select an output file")
            return
        
        # Get quality setting
        quality = self.quality_combo.currentText().lower()
        
        # Start the worker thread
        params = {
            "input_file": self.compress_input_file,
            "output_file": self.compress_output_file,
            "quality": quality
        }
        self.start_worker("compress", params, self.compress_progress)
    
    # Event handlers for PDF to Word tab
    def select_pdf_to_word_input(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)"
        )
        if file:
            self.pdf_word_file_label.setText(f"Input PDF: {file}")
            self.pdf_word_input_file = file
    
    def select_pdf_to_word_output(self):
        file, _ = QFileDialog.getSaveFileName(
            self, "Save Word File", "", "Word Files (*.docx)"
        )
        if file:
            self.pdf_word_output_label.setText(f"Output Word file: {file}")
            self.pdf_word_output_file = file
    
    def pdf_to_word(self):
        if not hasattr(self, "pdf_word_input_file"):
            QMessageBox.warning(self, "Warning", "Please select an input PDF file")
            return
        
        if not hasattr(self, "pdf_word_output_file"):
            QMessageBox.warning(self, "Warning", "Please select an output Word file")
            return
        
        # Start the worker thread
        params = {
            "input_file": self.pdf_word_input_file,
            "output_file": self.pdf_word_output_file
        }
        self.start_worker("to_word", params, self.pdf_word_progress)
    
    # Worker thread management
    def start_worker(self, operation, params, progress_bar):
        # Stop any existing worker
        if self.worker is not None and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
        
        # Create new worker
        self.worker = PDFWorker(operation, params)
        self.worker.progress.connect(progress_bar.setValue)
        self.worker.finished.connect(lambda success, msg: self.on_worker_finished(success, msg))
        
        # Reset progress bar and start worker
        progress_bar.setValue(0)
        self.worker.start()
    
    def on_worker_finished(self, success, message):
        if success:
            QMessageBox.information(self, "Success", message)
        else:
            QMessageBox.critical(self, "Error", message)


def main():
    app = QApplication(sys.argv)
    window = PDFToolkitApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()