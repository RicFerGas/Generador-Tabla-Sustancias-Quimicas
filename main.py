# main.py
import sys
import os
import json
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QFileDialog, QMessageBox, QProgressBar
)
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import  Qt,QThread, pyqtSignal
import openai
from hds_processor import HDSProcessor



if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # Running in a PyInstaller bundle
            BASE_DIR= os.path.dirname(os.path.abspath(__file__))
            INTERNAL_BASE_DIR = sys._MEIPASS
else:
            # Running in a normal Python environment
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            INTERNAL_BASE_DIR = BASE_DIR

# Ruta del archivo de configuración incluyendo ruta absoluta
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')

# Get the absolute path to the model directory
model_path = os.path.abspath('./models/xx_ent_wiki_sm/xx_ent_wiki_sm')

def check_tesseract_installed():
    """Check if Tesseract OCR is installed on the system."""
    try:
        subprocess.run(["tesseract", "-v"], check=True)
    except FileNotFoundError:
        print("Tesseract no está instalado. Por favor instálalo usando 'brew install tesseract'.")

# Función para guardar la configuración en un archivo JSON
def save_config(api_key):
    config_data = {
        "api_key": api_key
    }
    with open(CONFIG_FILE, 'w') as config_file:
        json.dump(config_data, config_file)

# Función para cargar la configuración desde el archivo JSON
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as config_file:
            config_data = json.load(config_file)
            return config_data.get("api_key")
    return None

class ConfigWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label_api = QLabel('Clave de API de OpenAI:', self)
        layout.addWidget(self.label_api)

        self.input_api = QLineEdit(self)
        self.input_api.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.input_api)

        self.btn_save = QPushButton('Guardar y Continuar', self)
        self.btn_save.clicked.connect(self.save_config_and_continue)
        layout.addWidget(self.btn_save)

        self.setLayout(layout)
        self.setWindowTitle('Configuración Inicial')
        self.setGeometry(200, 200, 400, 150)

    def save_config_and_continue(self):
        api_key = self.input_api.text()
        if not api_key:
            QMessageBox.warning(self, 'Error', 'Por favor, ingresa tu clave de API de OpenAI.')
            return

        # Guardar la clave API en el archivo de configuración
        save_config(api_key)

        # Cerrar la ventana de configuración y abrir la ventana principal
        self.close()
        self.main_window = MainApp(api_key)
        self.main_window.show()


class WorkerThread(QThread):
    file_processed = pyqtSignal(int, str)  # Signal to emit progress (percentage) and file name
    finished = pyqtSignal(bool)

    def __init__(self, input_dir, output_dir, api_key, project_name):
        super().__init__() 
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.api_key = api_key
        self.project_name = project_name

    def run(self):
        try:
            # Establecer la clave de API de OpenAI
            os.environ["OPENAI_API_KEY"] = self.api_key
            openai.api_key = self.api_key
            client = openai.OpenAI()
            processor= HDSProcessor(client,INTERNAL_BASE_DIR)

            # Definir el callback para el progreso
            def progress_callback(progress, filename):
                self.file_processed.emit(progress, filename)
            
            # Procesar el directorio de entrada
            
            _,_= processor.process_folder(
                self.input_dir,
                self.project_name,
                self.output_dir,
                progress_callback)

            self.finished.emit(True)
        except Exception as e:
            print(e)
            self.finished.emit(False)

class MainApp(QWidget):
    def __init__(self, api_key):
        super().__init__()
        self.api_key = api_key
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label_input_dir = QLabel('Directorio de PDFs de entrada:', self)
        layout.addWidget(self.label_input_dir)

        self.btn_browse_input = QPushButton('Seleccionar Directorio de Entrada', self)
        self.btn_browse_input.clicked.connect(self.browse_input_folder)
        layout.addWidget(self.btn_browse_input)

        self.label_output_dir = QLabel('Directorio de salida:', self)
        layout.addWidget(self.label_output_dir)

        self.btn_browse_output = QPushButton('Seleccionar Directorio de Salida', self)
        self.btn_browse_output.clicked.connect(self.browse_output_folder)
        layout.addWidget(self.btn_browse_output)

        self.label_project = QLabel('Nombre del Proyecto:', self)
        layout.addWidget(self.label_project)

        self.input_project = QLineEdit(self)
        layout.addWidget(self.input_project)

        self.btn_process = QPushButton('Procesar', self)
        self.btn_process.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_process)

        
        # Etiqueta para mostrar la animación de carga
        loading_path= os.path.join(INTERNAL_BASE_DIR, 'data_sets' ,'loading.gif')
        self.loading_label = QLabel(self)
        print(loading_path ,"loading gif seached path")
        self.loading_movie = QMovie(loading_path)
        self.loading_label.setMovie(self.loading_movie)
        self.loading_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.loading_label)
        # self.loading_label.hide()
        # Barra de progreso
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Etiqueta para mostrar el archivo actual
        self.label_current_file = QLabel('Archivo actual: Ninguno', self)
        layout.addWidget(self.label_current_file)

        self.setLayout(layout)
        self.setWindowTitle('Procesador de HDS')
        self.setGeometry(200, 200, 400, 250)

    def browse_input_folder(self):
        self.input_dir = QFileDialog.getExistingDirectory(self, 'Seleccionar Directorio de Entrada')
        if self.input_dir:
            self.label_input_dir.setText(f'Directorio de entrada: {self.input_dir}')

    def browse_output_folder(self):
        self.output_dir = QFileDialog.getExistingDirectory(self, 'Seleccionar Directorio de Salida')
        if self.output_dir:
            self.label_output_dir.setText(f'Directorio de salida: {self.output_dir}')

    def start_processing(self):
        project_name = self.input_project.text()

        if not hasattr(self, 'input_dir') or not self.input_dir:
            QMessageBox.warning(self, 'Error', 'Por favor, selecciona un directorio de entrada.')
            return
        if not hasattr(self, 'output_dir') or not self.output_dir:
            QMessageBox.warning(self, 'Error', 'Por favor, selecciona un directorio de salida.')
            return
        if not project_name:
            QMessageBox.warning(self, 'Error', 'Por favor, ingresa un nombre para el proyecto.')
            return

        # Deshabilitar botones mientras el proceso está en ejecución
        self.btn_browse_input.setEnabled(False)
        self.btn_browse_output.setEnabled(False)
        self.btn_process.setEnabled(False)

        # Mostrar un mensaje de inicio
        QMessageBox.information(self, 'Inicio de Procesamiento', 'El procesamiento ha comenzado.')
        self.loading_label.show()
        self.loading_movie.start()

        # Iniciar el hilo de procesamiento con las rutas de entrada y salida
        self.thread = WorkerThread(self.input_dir, self.output_dir, self.api_key, project_name)
        self.thread.file_processed.connect(self.update_progress)  # Conectar la señal de progreso
        self.thread.finished.connect(self.process_finished)  # Conectar cuando el proceso finaliza
        self.thread.start()

    def update_progress(self, progress, filename):
        self.progress_bar.setValue(progress)
        self.label_current_file.setText(f'Archivo actual: {filename}')

    def process_finished(self, success):
        if success:
            QMessageBox.information(self, 'Éxito', 'El procesamiento ha finalizado correctamente.')
        else:
            QMessageBox.critical(self, 'Error', 'Ha ocurrido un error durante el procesamiento.')

        # Volver a habilitar los botones
        self.btn_browse_input.setEnabled(True)
        self.btn_browse_output.setEnabled(True)
        self.btn_process.setEnabled(True)

        #Detener y ocultar la animación de carga
        self.loading_movie.stop()
        self.loading_label.hide()

        # Resetear la barra de progreso y etiqueta del archivo actual
        self.progress_bar.setValue(0)
        self.label_current_file.setText('Archivo actual:Inicializando el sistema')

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)

    # Cargar la clave API de la configuración si ya existe
    api_key = load_config()
    check_tesseract_installed()

    if not api_key:
        # Si no hay configuración previa, mostrar el menú de configuración
        config_window = ConfigWindow()
        config_window.show()
    else:
        # Si ya tenemos configuración, iniciar la aplicación principal
        main_window = MainApp(api_key)
        main_window.show()

    sys.exit(app.exec_())