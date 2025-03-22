# project_manager.py
import json
import os
from pathlib import Path
from datetime import datetime
from PyQt6.QtWidgets import (QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
                           QFileDialog, QPushButton, QLineEdit, QLabel, QMessageBox, 
                           QTableWidget, QTableWidgetItem)
from PyQt6.QtCore import Qt
import pandas as pd
import openai
from api.services.excel_postprocess import GeneradorTablaSustQ
from api.services.preprocess import DocumentPreprocessor
from api.services.extract_info import extract_info_from_hds_txt


class ProjectWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SIGASH HDS Manager")
        self.setGeometry(100, 100, 1200, 800)
        
        self.generator = GeneradorTablaSustQ()
        self.doc_processor = DocumentPreprocessor()
        self.client= openai.OpenAI()
        
        self.current_project = None
        self.project_data = {}
        self.setup_ui()

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        
        # Project controls
        project_group = QWidget()
        project_layout = QVBoxLayout()
        
        self.project_name = QLineEdit()
        self.project_name.setPlaceholderText("Project Name")
        project_layout.addWidget(self.project_name)
        
        btn_create = QPushButton("Create Project")
        btn_create.clicked.connect(self.create_project)
        project_layout.addWidget(btn_create)
        
        btn_analyze = QPushButton("Analyze HDS Folder")
        btn_analyze.clicked.connect(self.analyze_folder)
        project_layout.addWidget(btn_analyze)
        
        project_group.setLayout(project_layout)
        layout.addWidget(project_group)

        # Tabs
        self.tabs = QTabWidget()
        self.setup_data_tab()
        self.setup_gei_tab()
        layout.addWidget(self.tabs)
        
        main_widget.setLayout(layout)

    def setup_data_tab(self):
        data_tab = QWidget()
        layout = QVBoxLayout()
        
        self.data_table = QTableWidget()
        layout.addWidget(self.data_table)
        
        btn_group = QWidget()
        btn_layout = QVBoxLayout()
        
        btn_excel = QPushButton("Generate Excel")
        btn_excel.clicked.connect(self.generate_excel)
        btn_layout.addWidget(btn_excel)
        
        btn_save = QPushButton("Save JSON")
        btn_save.clicked.connect(self.save_json)
        btn_layout.addWidget(btn_save)
        
        btn_group.setLayout(btn_layout)
        layout.addWidget(btn_group)
        
        data_tab.setLayout(layout)
        self.tabs.addTab(data_tab, "HDS Data")

    def setup_gei_tab(self):
        gei_tab = QWidget()
        layout = QVBoxLayout()
        
        self.gei_table = QTableWidget()
        layout.addWidget(self.gei_table)
        
        btn_calculate = QPushButton("Calculate GEI")
        btn_calculate.clicked.connect(self.calculate_gei)
        layout.addWidget(btn_calculate)
        
        gei_tab.setLayout(layout)
        self.tabs.addTab(gei_tab, "GEI Calculator")

    def create_project(self):
        name = self.project_name.text()
        if not name:
            QMessageBox.warning(self, "Error", "Please enter a project name")
            return
            
        self.current_project = name
        self.project_data = {
            "name": name,
            "created_date": datetime.now().isoformat(),
            "hds_data": [],
            "gei_calculations": {}
        }
        
        project_dir = Path(f"projects/{name}")
        project_dir.mkdir(parents=True, exist_ok=True)
        self.save_project_state()

    def analyze_folder(self):
        if not self.current_project:
            QMessageBox.warning(self, "Error", "Please create a project first")
            return
            
        folder = QFileDialog.getExistingDirectory(self, "Select HDS Folder")
        if folder:
            # Process PDFs using existing logic
            processed_data = []
            for file in Path(folder).glob("*.pdf"):
                hds_text = self.doc_processor.extract_text(file)
                hds_data = extract_info_from_hds_txt(hds_text,self.client)
                processed_data.append(hds_data.model_dump())
                
            # Flatten data using existing GeneradorTablaSustQ
            self.project_data["hds_data"] = self.generator.flatten_hds_data(processed_data)
            self.update_data_view()
            self.save_project_state()

    def update_data_view(self):
        if not self.project_data.get("hds_data"):
            return
            
        df = pd.DataFrame(self.project_data["hds_data"])
        self.data_table.setRowCount(len(df))
        self.data_table.setColumnCount(len(df.columns))
        self.data_table.setHorizontalHeaderLabels(df.columns)
        
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.data_table.setItem(i, j, item)

    def calculate_gei(self):
        if not self.project_data.get("hds_data"):
            return
            
        gei_data = []
        for substance in self.project_data["hds_data"]:
            if substance.get("Sujeta a GEI"):
                gei_entry = {
                    "nombre": substance["Nombre de la Sustancia Química"],
                    "cas": substance["Número CAS del Componente"],
                    "pcg": substance.get("Potencial de Calentamiento Global", 0)
                }
                gei_data.append(gei_entry)
                
        self.show_gei_calculator(gei_data)

    def generate_excel(self):
        if not self.current_project or not self.project_data.get("hds_data"):
            return
            
        output_file = f"projects/{self.current_project}/output.xlsx"
        self.generator.export_to_excel_with_template(
            self.project_data["hds_data"], 
            output_file
        )
        QMessageBox.information(self, "Success", "Excel file generated successfully")

    def save_project_state(self):
        if not self.current_project:
            return
            
        project_file = Path(f"projects/{self.current_project}/project.json")
        with open(project_file, "w") as f:
            json.dump(self.project_data, f)

    def save_json(self):
        """Save current HDS data to JSON file"""
        if not self.current_project:
            QMessageBox.warning(self, "Error", "Please create a project first")
            return
            
        if not self.project_data.get("hds_data"):
            QMessageBox.warning(self, "Error", "No HDS data to save")
            return
            
        try:
            # Save to project directory
            json_file = Path(f"projects/{self.current_project}/hds_data.json")
            json_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(json_file, "w", encoding="utf-8") as f:
                json.dump(self.project_data["hds_data"], f, ensure_ascii=False, indent=2)
                
            QMessageBox.information(
                self, 
                "Success", 
                f"Data saved to:\n{json_file.absolute()}"
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                "Error",
                f"Failed to save JSON file:\n{str(e)}"
            )

if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = ProjectWindow()
    window.show()
    sys.exit(app.exec())