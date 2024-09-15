# SIGASH HDS Processor

Este proyecto, desarrollado para SIGASH por **Ricardo Fernandez**, es una aplicación que permite procesar Hojas de Datos de Seguridad (HDS) en formato PDF y generar un archivo Excel y JSON con la información extraída. El procesamiento de los PDF utiliza un **Large Language Model (LLM)**, un modelo de lenguaje a gran escala potenciado por inteligencia artificial (IA).

## ¿Qué es un LLM?

Un **Large Language Model (LLM)** es un modelo de lenguaje entrenado utilizando grandes cantidades de datos textuales, capaz de entender y generar lenguaje natural con alto grado de precisión. Los LLMs, como el que se utiliza en esta aplicación, están impulsados por algoritmos avanzados de inteligencia artificial que pueden comprender el contexto, extraer información y generar respuestas complejas. 

En este caso, la inteligencia artificial se utiliza para analizar automáticamente los contenidos de las Hojas de Datos de Seguridad (HDS) y extraer datos importantes, como los componentes químicos, porcentajes, advertencias, entre otros. Todo esto es hecho de manera automática, facilitando la tarea de análisis y procesamiento de datos, lo que hace que el sistema sea más eficiente y confiable para los usuarios de SIGASH.

## Requisitos del Sistema

- **Sistema Operativo**: Windows, macOS o Linux
- **Python**: Versión 3.9 o superior (Instrucciones para instalar más adelante)

### Nota sobre el Entorno Virtual

El entorno virtual está excluido del repositorio. Se recomienda que cada usuario cree su propio entorno virtual para evitar conflictos de dependencias. Esto también asegura que los usuarios puedan instalar las dependencias correctas según su sistema operativo y configuración.

## Instalación

### Paso 1: Descargar el Proyecto

Si no tienes Git instalado o no sabes cómo usarlo, puedes descargar el proyecto directamente desde el siguiente enlace:

1. Ve a la [página del repositorio del proyecto](https://github.com/RicFerGas/Generador-Tabla-Sustancias-Quimicas).
2. Haz clic en el botón verde que dice **"Code"**.
3. Selecciona **"Download ZIP"** para descargar el proyecto a tu computadora.
4. Descomprime el archivo ZIP en una carpeta de tu elección.

### Paso 2: Instalar Python

Si no tienes Python instalado, sigue estos pasos:

#### En Windows:

1. Ve a la página de descargas de Python: [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Descarga el instalador de **Python 3.9**.
3. Asegúrate de marcar la casilla que dice **"Add Python to PATH"** durante la instalación.
4. Finaliza la instalación.

#### En macOS:

1. Descarga Python desde [https://www.python.org/downloads/](https://www.python.org/downloads/) o instala usando Homebrew:
 ```bash
   brew install python
```
#### En Linux (Ubuntu/Debian):

 ```bash
sudo apt update
sudo apt install python3 python3-pip
```

### Paso 3: Instalar PyInstaller

PyInstaller es la herramienta que utilizaremos para generar el ejecutable del programa. Para instalar PyInstaller, abre una terminal o línea de comandos y ejecuta:

 ```bash
pip install pyinstaller
```

### Paso 4: Instalar las Dependencias del Proyecto

Con Python y PyInstaller instalados, ahora necesitas instalar las dependencias del proyecto. Asegúrate de estar en el directorio donde descargaste el proyecto.

#### Moverse al directorio del proyecto

1. **En Windows**:
   Abre el Explorador de Archivos, navega a la carpeta donde descomprimiste el proyecto, luego haz clic derecho y selecciona **"Abrir ventana de comandos aquí"** o **"Abrir en terminal"**.

   Si no tienes esa opción, abre una ventana de comandos y navega manualmente:
    ```bash
   cd C:\ruta\de\la\carpeta\hds-processor
   ```

2. **En macOS/Linux**:
   Abre una terminal y navega al directorio donde descargaste el proyecto:
    ```bash
   cd /ruta/de/la/carpeta/hds-processor
   ```

#### Instalar las dependencias

Con la terminal en el directorio del proyecto, ejecuta el siguiente comando para instalar las dependencias:

 ```bash
pip install -r requirements.txt
```

### Paso 5: Generar el Ejecutable

Una vez que tengas las dependencias instaladas, puedes generar el ejecutable que no necesitará de Python en el futuro. Para hacerlo, ejecuta el siguiente comando en la terminal:

 ```bash
pyinstaller --onefile --noconsole main.py
```

Este comando generará un archivo ejecutable que se guardará en la carpeta `dist`. Este archivo será un ejecutable independiente que puedes distribuir a otros usuarios sin necesidad de que tengan Python instalado.

### Paso 6: Ejecutar el Programa

Una vez que tengas el ejecutable, puedes simplemente hacer doble clic en él para abrir el programa. Alternativamente, puedes ejecutarlo desde la línea de comandos:

#### En Windows:

 ```bash
cd dist
main.exe
```

#### En macOS/Linux:

 ```bash
cd dist
./main
```

### Paso 7: Configurar la Clave API de OpenAI

Cuando ejecutes el programa por primera vez, te pedirá que ingreses tu clave de API de OpenAI. Esta clave se guardará en un archivo `config.json` para que no tengas que ingresarla nuevamente.

### Entrada:

- El programa requiere un directorio con archivos PDF de Hojas de Datos de Seguridad (HDS) para procesar.

### Salida:

- La aplicación generará dos archivos en el directorio de salida:
  - Un archivo **Excel** con la información estructurada de las HDS.
  - Un archivo **JSON** con la información completa extraída de las HDS.

## Contacto

Si tienes alguna pregunta o necesitas ayuda, por favor contacta con:

- **Ricardo Fernandez** - Autor del proyecto
- **SIGASH**
