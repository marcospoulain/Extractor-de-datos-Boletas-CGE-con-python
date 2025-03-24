import sys
import os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar
from PyQt6.QtGui import QPixmap
from PyQt6.QtCore import Qt
import pdfplumber
import pandas as pd
import re
import subprocess

class PDFExtractorGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_path = ''
        self.output_file = ''
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Extractor boletas CGE')
        self.setGeometry(100, 100, 400, 300)

        layout = QVBoxLayout()

        # Logo layout
        logo_layout = QHBoxLayout()
        cge_logo = QLabel()
        cge_pixmap = QPixmap('programa/path_to_cge_logo.png')  # Reemplaza con la ruta real
        cge_logo.setPixmap(cge_pixmap.scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio))
        logo_layout.addWidget(cge_logo)

        arica_logo = QLabel()
        arica_pixmap = QPixmap('programa/path_to_arica_logo.png')  # Reemplaza con la ruta real
        arica_logo.setPixmap(arica_pixmap.scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio))
        logo_layout.addWidget(arica_logo)

        layout.addLayout(logo_layout)

        # Buttons
        self.select_pdf_btn = QPushButton('Seleccionar PDF', self)
        self.select_pdf_btn.clicked.connect(self.select_pdf)
        layout.addWidget(self.select_pdf_btn)

        self.extract_btn = QPushButton('Extraer Datos', self)
        self.extract_btn.clicked.connect(self.extract_data)
        self.extract_btn.setEnabled(False)
        layout.addWidget(self.extract_btn)

        self.output_label = QLabel('Salida: No generada aún', self)
        layout.addWidget(self.output_label)

        # Progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid white;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #090c9b;
                width: 10px;
                margin: 1px;
            }
        """)
        layout.addWidget(self.progress_bar)

        # New button to open file location
        self.open_location_btn = QPushButton('Abrir Ubicación del Archivo', self)
        self.open_location_btn.clicked.connect(self.open_file_location)
        self.open_location_btn.setEnabled(False)
        layout.addWidget(self.open_location_btn)

        self.setLayout(layout)

    def select_pdf(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo PDF", "", "Archivos PDF (*.pdf)")
        if file_name:
            self.pdf_path = file_name
            self.select_pdf_btn.setText(f'Seleccionado: {os.path.basename(file_name)}')
            self.extract_btn.setEnabled(True)

    def extract_data(self):
        try:
            text = self.extract_text_from_pdf(self.pdf_path)
            sections = re.split(r'(?:N° Cliente|S\.I\.I\.-SANTIAGO ORIENTE|FACTURA ELECTRÓNICA|BOLETA ELECTRÓNICA)\s*\d+', text)[1:]
            
            # Set progress bar maximum
            self.progress_bar.setMaximum(len(sections))
            
            data_list = []
            for i, section in enumerate(sections):
                match = re.search(r'(N° Cliente|S\.I\.I\.-SANTIAGO ORIENTE|FACTURA ELECTRÓNICA|BOLETA ELECTRÓNICA)\s*(\d+)', text)
                if match:
                    sections[i] = match.group(0) + section
                data_list.append(self.process_text(section))
                
                # Update progress bar
                self.progress_bar.setValue(i + 1)
                self.progress_bar.setFormat(f"Loading {int((i + 1) / len(sections) * 100)}%")

            self.output_file = os.path.join(os.path.dirname(self.pdf_path), "datos_extraidos.xlsx")
            self.save_to_excel(data_list, self.output_file)
            
            self.output_label.setText(f'Salida: {self.output_file}')
            self.open_location_btn.setEnabled(True)
            QMessageBox.information(self, "Éxito", f"Datos extraídos y guardados en {self.output_file}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error: {str(e)}")

    def extract_text_from_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join(page.extract_text() for page in pdf.pages)

    def open_file_location(self):
        if self.output_file:
            file_path = os.path.dirname(self.output_file)
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.Popen(["open", file_path])
            else:  # linux variants
                subprocess.Popen(["xdg-open", file_path])

    def safe_search(self, pattern, text, group_index=1, default=""):
        match = re.search(pattern, text)
        return match.group(group_index) if match else default

    def extract_client_numbers(self, text):
        matches = re.findall(r'N° Cliente.*?(\d{7})', text, re.DOTALL)
        if not matches:
            matches = re.findall(r'\b(\d{7})\b', text)
        return matches[0] if matches else ""

    def process_text(self, text):
        data = {}
        
        data['Nº Cliente'] = self.extract_client_numbers(text)
        data['RUT'] = self.safe_search(r'R\.U\.T\.: (\d+\.\d+\.\d+-\d)', text)
        
        data['Nº Factura'] = self.safe_search(r'FACTURA ELECTRÓNICA\s+Nº (\d+)', text)
        if not data['Nº Factura']:
            data['Nº Factura'] = self.safe_search(r'BOLETA ELECTRÓNICA\s+Nº (\d+)', text)
        
        data['Fecha de emisión'] = self.safe_search(r'Fecha de emisión: (\d+ \w+ \d{4})', text)
        data['Total a pagar'] = self.safe_search(r'Total a pagar \$?([\d,.]+)', text)
        data['Fecha de vencimiento'] = self.safe_search(r'Fecha de vencimiento\s+(\d+ \w+ \d{4})', text)
        
        ultimo_pago_fecha = self.safe_search(r'Último Pago: el (\d+ \w+ \d{4})', text)
        ultimo_pago_monto = self.safe_search(r'Último Pago: el \d+ \w+ \d{4} por un monto de \$([\d,.]+)', text)
        data['Último Pago'] = f"{ultimo_pago_fecha} por un monto de ${ultimo_pago_monto}"

        details = re.findall(r'([^\n]+?)\s+\$\s*([\d,.]+)', text)
        for item, amount in details:
            data[item.strip()] = amount.strip()

        consumo_match = re.search(r'Consumo total del mes\s*=\s*([\d,.]+)\s*kWh', text, re.IGNORECASE)
        if consumo_match:
            data['Consumo total del mes (kWh)'] = consumo_match.group(1).replace(',', '.')
        else:
            consumo_match = re.search(r'Consumo total del mes\s*([\d,.]+)\s*kWh', text, re.IGNORECASE)
            if consumo_match:
                data['Consumo total del mes (kWh)'] = consumo_match.group(1).replace(',', '.')
            else:
                data['Consumo total del mes (kWh)'] = "No encontrado"

        return data

    def save_to_excel(self, data_list, excel_path):
        df = pd.DataFrame(data_list)
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Datos Extraídos', index=False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFExtractorGUI()
    ex.show()
    sys.exit(app.exec())