# process_hds.py
import os
import json
import openai
from schemas import HDSData
from pydantic import ValidationError
import pdfplumber
import pandas as pd
from dotenv import load_dotenv
from pdf2image import convert_from_path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import pytesseract
import spacy


nlp = spacy.load('xx_ent_wiki_sm')

def validate_extracted_text(text: str, threshold: float = 0.6) -> bool:
    """
    Validates the extracted text from a PDF based on the word count threshold.

    Args:
        text (str): The extracted text from the PDF.
        threshold (float, optional): The threshold for the word count ratio. Defaults to 0.6.

    Returns:
        bool: True if the word count ratio exceeds the threshold, False otherwise.
    """
    # Tokenize and count valid words
    doc = nlp(text)
    word_count = sum(1 for token in doc if token.is_alpha)  # Only count alphabetic words
    total_count = len(doc)
    print(f"word count{word_count}, total count {total_count} rate {word_count / total_count}")

    # If the percentage of valid words exceeds the threshold, consider it a good extract
    if total_count == 0:
        return False

    return (word_count / total_count) >= threshold


def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extracts text from a PDF file.
    Args:
        pdf_path (str): The path to the PDF file.
    Returns:
        str: The extracted text from the PDF file.
    Raises:
            None
     """
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()

    # Si el texto extraído pasa la validación, lo retornamos
    if text and validate_extracted_text(text):
        return text
    else:
        print("PDF sin texto, procesando como imagen")
        # Si no, usamos OCR para extraer el texto
        images = convert_from_path(pdf_path)
        extracted_text = ""
        for image in images:
            extracted_text += pytesseract.image_to_string(image, lang='spa')
        return extracted_text

def prompt_generation(HDS):
    system_prompt = """
    Eres un asistente que extrae información de Hojas de Datos de Seguridad (HDS) para estudios de riesgo por sustancias químicas. 
    Extrae la información relevante de la HDS proporcionada y devuélvela siguiendo el formato especificado, para el caso de las propiedades numericas,
    no pongas 0 si no ecnuentras la propiedad, el valor default es none."""
    user_prompt = f"La HDS es la siguiente:\n\n{HDS}"
    return system_prompt, user_prompt

def get_completion_openai(system_prompt, user_prompt, client,model="gpt-4o-2024-08-06", max_tokens=2500):
    
    completion = client.beta.chat.completions.parse(
    model=model,
    messages=[
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ],
    # Pasamos el modelo Pydantic directamente
    response_format=HDSData,
    max_tokens=max_tokens,
    temperature=0.0,
    top_p=1,
)

# Obtenemos el objeto parseado directamente
    hds_data = completion.choices[0].message.parsed
    return hds_data

def flatten_hds_data(hds_dict):
    sustancias_flat = []

    # Información común para todas las filas de la sustancia
    common_info = {
        'Archivo': hds_dict.get('Archivo'),
        'Nombre de la Sustancia Química': hds_dict.get('nombre_sustancia_quimica'),
        'Idioma de la HDS': hds_dict.get('idioma'),
        'pH de la sustancia': hds_dict.get('ph'),
        'Palabra de Advertencia': hds_dict.get('palabra_advertencia'),
        'Indicaciones de toxicología': hds_dict.get('indicaciones_toxicologia'),
        'Pictogramas': ', '.join(hds_dict.get('pictogramas', [])),
        # Agrega aquí otras propiedades no anidadas o que no sean listas
    }

    # Desanidar propiedades como Temperatura de ebullición, Punto de inflamación, etc.
    propiedades = ['temperatura_ebullicion', 'punto_congelacion', 'densidad', 'punto_inflamacion', 'solubilidad_agua']
    for prop in propiedades:
        propiedad = hds_dict.get(prop)
        if propiedad:
            common_info[f'{prop} (Valor)'] = propiedad.get('valor')
            common_info[f'{prop} (Unidades)'] = propiedad.get('unidades')
        else:
            common_info[f'{prop} (Valor)'] = None
            common_info[f'{prop} (Unidades)'] = None

    # Limites de inflamabilidad
    common_info['Límite inferior de inflamabilidad'] = hds_dict.get('limite_inf_inflamabilidad')
    common_info['Límite superior de inflamabilidad'] = hds_dict.get('limite_sup_inflamabilidad')

    # Manejar componentes
    componentes = hds_dict.get('componentes', [])
    if componentes:
        for componente in componentes:
            row = common_info.copy()
            row['Nombre del Componente'] = componente.get('nombre')
            row['Número CAS del Componente'] = componente.get('numero_cas')
            row['Porcentaje del Componente'] = componente.get('porcentaje')
            sustancias_flat.append(row)
    else:
        # Si no hay componentes, agregar la sustancia sin detalles de componentes
        sustancias_flat.append(common_info)

    # Manejar Indicaciones de peligro H y Consejos de Prudencia P
    indicaciones_h = hds_dict.get('identificaciones_peligro_h', [])
    consejos_p = hds_dict.get('consejos_prudencia_p', [])

    # Concatenar códigos y descripciones
    common_info['Identificaciones de peligro H'] = '; '.join([f"{i['codigo']}: {i['descripcion']}" for i in indicaciones_h])
    common_info['Consejos de Prudencia P'] = '; '.join([f"{i['codigo']}: {i['descripcion']}" for i in consejos_p])

    # Actualizar las filas existentes con las indicaciones
    for row in sustancias_flat:
        row.update({
            'Identificaciones de peligro H': common_info['Identificaciones de peligro H'],
            'Consejos de Prudencia P': common_info['Consejos de Prudencia P']
        })

    return sustancias_flat


def export_to_excel(data, output_file="output.xlsx"):
    # Crear DataFrame de pandas
    df = pd.DataFrame(data)

    # Definir el orden deseado de las columnas
    column_order = [
        'Archivo',                         # 1. Nombre del archivo
        'Nombre de la Sustancia Química',   # 2. Nombre de la sustancia
        'Idioma de la HDS',                 # 3. Idioma
        'Nombre del Componente',            # 4. Componente
        'Número CAS del Componente',        # 5. CAS del componente
        'Porcentaje del Componente',        # 6. Porcentaje
        'Palabra de Advertencia',           # 7. Palabra de advertencia
        'Indicaciones de toxicología',      # 8. Indicaciones de toxicología
        'Pictogramas',                      # 9. Pictogramas
        'Identificaciones de peligro H',    # 10. Identificaciones de peligro H
        'Consejos de Prudencia P',          # 11. Consejos de prudencia P
        'temperatura_ebullicion (Valor)',   # 12. Propiedades: temperatura de ebullición
        'temperatura_ebullicion (Unidades)',
        'punto_congelacion (Valor)',        # 13. Propiedades: punto de congelación
        'punto_congelacion (Unidades)',
        'densidad (Valor)',                 # 14. Propiedades: densidad
        'densidad (Unidades)',
        'punto_inflamacion (Valor)',        # 15. Propiedades: punto de inflamación
        'punto_inflamacion (Unidades)',
        'solubilidad_agua (Valor)',         # 16. Propiedades: solubilidad en agua
        'solubilidad_agua (Unidades)',
        'Límite inferior de inflamabilidad', # 17. Límites de inflamabilidad
        'Límite superior de inflamabilidad'
    ]

    # Reordenar las columnas del DataFrame
    df = df[column_order]

    # Guardar el DataFrame a un archivo Excel temporal
    temp_file = "temp_output.xlsx"
    df.to_excel(temp_file, index=False)

    # Cargar el archivo Excel con openpyxl para personalizarlo
    wb = load_workbook(temp_file)
    ws = wb.active

    # Combinación de celdas para las columnas especificadas solo si la sustancia es la misma
    merge_columns = [
        'Archivo', 'Nombre de la Sustancia Química', 'Idioma de la HDS', 'Palabra de Advertencia',
        'Indicaciones de toxicología', 'Pictogramas', 'Identificaciones de peligro H', 'Consejos de Prudencia P',
        'temperatura_ebullicion (Valor)', 'temperatura_ebullicion (Unidades)', 'punto_congelacion (Valor)',
        'punto_congelacion (Unidades)', 'densidad (Valor)', 'densidad (Unidades)', 'punto_inflamacion (Valor)',
        'punto_inflamacion (Unidades)', 'solubilidad_agua (Valor)', 'solubilidad_agua (Unidades)',
        'Límite inferior de inflamabilidad', 'Límite superior de inflamabilidad'
    ]

    # Combinación de celdas para las columnas especificadas solo si la sustancia es la misma
    current_values = {col: None for col in merge_columns}
    start_rows = {col: 2 for col in merge_columns}  # Asume que comienza en la fila 2 (después del encabezado)
    current_sustancia = None
    start_row_sustancia = 2

    for row in range(2, ws.max_row + 1):
        sustancia_value = ws[f'B{row}'].value  # Columna B para Nombre de la Sustancia Química
        if sustancia_value != current_sustancia:
            # Si la sustancia cambia, combinar las celdas previas
            for col_name in merge_columns:
                if current_values[col_name] is not None:
                    ws.merge_cells(start_row=start_rows[col_name], start_column=df.columns.get_loc(col_name) + 1, 
                                   end_row=row - 1, end_column=df.columns.get_loc(col_name) + 1)
                current_values[col_name] = ws[f'{get_column_letter(df.columns.get_loc(col_name) + 1)}{row}'].value
                start_rows[col_name] = row
            current_sustancia = sustancia_value
            start_row_sustancia = row
        else:
            # Si la sustancia no cambia, verificar las columnas para combinar
            for col_name in merge_columns:
                col_letter = get_column_letter(df.columns.get_loc(col_name) + 1)  # Obtener la letra de la columna
                cell_value = ws[f'{col_letter}{row}'].value

                if cell_value == current_values[col_name]:
                    ws[f'{col_letter}{row}'].value = None  # Vaciar celdas duplicadas
                else:
                    # Si detecta un valor diferente, combinar las celdas anteriores
                    if current_values[col_name] is not None:
                        ws.merge_cells(start_row=start_rows[col_name], start_column=df.columns.get_loc(col_name) + 1, 
                                       end_row=row - 1, end_column=df.columns.get_loc(col_name) + 1)
                    current_values[col_name] = cell_value
                    start_rows[col_name] = row

    # Combinar las celdas del último grupo de la misma sustancia
    for col_name in merge_columns:
        if current_values[col_name] is not None:
            ws.merge_cells(start_row=start_rows[col_name], start_column=df.columns.get_loc(col_name) + 1, 
                           end_row=ws.max_row, end_column=df.columns.get_loc(col_name) + 1)

    # Fijar el ancho de las columnas según lo que esperas ver
    fixed_column_widths = {
        'A': 10,  # Archivo
        'B': 12,  # Nombre de la Sustancia Química
        'C': 5,  # Idioma de la HDS
        'D': 12,  # Nombre del Componente
        'E': 6,  # Número CAS del Componente
        'F': 6,  # Porcentaje del Componente
        'G': 10,  # Palabra de Advertencia
        'H': 25,  # Indicaciones de toxicología
        'I': 20,  # Pictogramas
        'J': 20,  # Identificaciones de peligro H
        'K': 20,  # Consejos de Prudencia P
        'L': 8,  # Temperatura de ebullición (Valor)
        'M': 8,  # Temperatura de ebullición (Unidades)
        'N': 8,  # Punto de congelación (Valor)
        'O': 8,  # Punto de congelación (Unidades)
        'P': 8,  # Densidad (Valor)
        'Q': 8,  # Densidad (Unidades)
        'R': 8,  # Punto de inflamación (Valor)
        'S': 8,  # Punto de inflamación (Unidades)
        'T': 8,  # Solubilidad en agua (Valor)
        'U': 8,  # Solubilidad en agua (Unidades)
        'V': 8,  # Límite inferior de inflamabilidad
        'W': 8,  # Límite superior de inflamabilidad
    }

    for col, width in fixed_column_widths.items():
        ws.column_dimensions[col].width = width

    # Ajustar la altura de las filas para que se vea todo el texto
    for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
        max_height = 1
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='center')  # Ajuste automático del texto y centrado
            if cell.value and isinstance(cell.value, str):
                max_height = max(max_height, len(cell.value.split('\n')))
        ws.row_dimensions[cell.row].height = max_height * 15  # Ajuste según la cantidad de líneas

    # Guardar el archivo final
    wb.save(output_file)


def process_hds_folder(folder_path, project_name, client, excel_output, json_output, progress_callback=None):
    all_sustancias = []
    errores = []
    sustancias_data = []  # Lista para almacenar los datos procesados para Excel

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    total_files = len(pdf_files)

    for index, filename in enumerate(pdf_files):
        file_path = os.path.join(folder_path, filename)

        # Leer el contenido de la HDS
        hds_text = extract_text_from_pdf(file_path)
        # Generar el prompt
        system_prompt, user_prompt = prompt_generation(hds_text)

        try:
            # Obtener el objeto HDSData directamente del LLM
            hds_data = get_completion_openai(system_prompt, user_prompt, client)

            # Convertir el objeto HDSData a dict
            hds_dict = hds_data.dict()
            hds_dict["Archivo"] = filename

            # Postprocesar para Excel
            sustancias_data.extend(flatten_hds_data(hds_dict))

            print(f"Procesado correctamente: {filename}")
            all_sustancias.append(hds_dict)
        except Exception as e:
            print(f"Error al procesar {filename}: {e}")
            errores.append(filename)

        # Calcular el progreso y llamar al callback
        if progress_callback:
            progress_percent = int((index + 1) / total_files * 100)
            progress_callback(progress_percent, filename)

    # Guardar el JSON final con todas las sustancias
    with open(json_output, 'w', encoding='utf-8') as output_file:
        json.dump(all_sustancias, output_file, indent=2, ensure_ascii=False)

    # Generar el Excel formateado con celdas combinadas y alineadas
    export_to_excel(sustancias_data, excel_output)

    print(f"Proceso completado. Los archivos {json_output} y {excel_output} han sido generados.")
    return errores, all_sustancias


# load_dotenv()
# openai.api_key = os.getenv("OPENAI_API_KEY")
# example_path="ejemplo/"
# client = openai.OpenAI()
# process_hds_folder(example_path, "Ejemplo",client)