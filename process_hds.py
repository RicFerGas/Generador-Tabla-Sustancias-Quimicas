# process_hds.py
import os
import sys
import json
import openai
import pandas as pd
from dotenv import load_dotenv
from pdf2image import convert_from_path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import pytesseract
import PyPDF2
import spacy
import fitz  # PyMuPDF
from rapidfuzz import fuzz, process
from schemas import HDSData


# Determine base path
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
model_relative_path = os.path.join('models', 'xx_ent_wiki_sm', 'xx_ent_wiki_sm', 'xx_ent_wiki_sm-3.8.0')
model_path = os.path.join(base_path, model_relative_path)

# Load the model
nlp = spacy.load(model_path)

def is_valid_pdf(pdf_path: str) -> bool:
    """
    Validates if the PDF is structurally correct and can be opened.
    """
    if not os.path.isfile(pdf_path):
        print(f"File does not exist: {pdf_path}")
        return False

    try:
        # Check if it starts with %PDF-
        with open(pdf_path, 'rb') as file:
            header = file.read(4)
            if header != b'%PDF':
                print(f"File is not a valid PDF: {pdf_path}")
                return False

            # Attempt to open and read the PDF
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)

            # Check if there are any pages
            if num_pages == 0:
                print(f"PDF has no pages: {pdf_path}")
                return False
            
            # Check if any page has extractable text
            text_found = False
            for page in pdf_reader.pages:
                if page.extract_text():
                    text_found = True
                    break

            if not text_found:
                print(f"No extractable text in PDF: {pdf_path}")
                return False
            
            print(f"PDF is valid: {pdf_path} (pages: {num_pages})")
            return True
            
    except Exception as e:
        print(f"Invalid PDF {pdf_path}: {e}")
        return False
def validate_extracted_text(text: str, threshold: float = 0.6) -> bool:
    """
    Validates the extracted text from a PDF based on the word count threshold.
    """
    try:
        doc = nlp(text)
        word_count = sum(1 for token in doc if token.is_alpha)
        total_count = len(doc)
        print(f"word count: {word_count}, total count: {total_count}, rate: {word_count / total_count}")
        if total_count == 0:
            return False
        return (word_count / total_count) >= threshold
    except Exception as e:
        print(f"Error in text validation: {e}")
        return False

def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extracts text from a PDF file, first using PyMuPDF and then OCR as a fallback.

    Args:
        pdf_path (str): The path to the PDF file.

    Returns:
        str: The extracted text from the PDF file.
    """
    text = ""

    # Intentar extraer texto usando PyMuPDF
    try:
        with fitz.open(pdf_path) as pdf:
            for page_num in range(len(pdf)):
                page = pdf.load_page(page_num)
                text += page.get_text()
    except Exception as e:
        print(f"Error extracting text from {pdf_path} using PyMuPDF: {e}")
        text = ""

    # Si se extrajo texto y pasa la validación, devolverlo
    if text and validate_extracted_text(text):
        return text
    else:
        print("No valid text found or extraction failed, proceeding to OCR.")

    # Intentar extraer texto usando OCR
    try:
        images = convert_from_path(pdf_path)
        extracted_text = ""
        for image in images:
            extracted_text += pytesseract.image_to_string(image, lang='spa')
        return extracted_text
    except Exception as e:
        print(f"Error during OCR process: {e}")
        return "OCR extraction failed."




def prompt_generation(hds: str) -> tuple:
    """
    Generates the system and user prompts for the OpenAI API based on the HDS text."""
    system_prompt = """
    Eres un asistente que extrae información de Hojas de Datos de Seguridad (HDS) para estudios de riesgo por sustancias químicas. 
    La información extraida debe estar siempre ** EN ESPAÑOL ** y seguir el formato especificado.
    Extrae la información relevante de la HDS proporcionada y devuélvela siguiendo el formato especificado, para el caso de las propiedades numericas,
    no pongas 0 si no ecnuentras la propiedad, el valor default es none."""
    user_prompt = f"La HDS es la siguiente:\n\n{hds}"
    return system_prompt, user_prompt

def get_completion_openai(system_prompt: str, user_prompt: str,
                          client ,model:str ="gpt-4o-2024-08-06",
                          max_tokens:int=4000):
    """
    Gets the completion from the OpenAI API using the system and user prompts.
    
    Args:
        system_prompt (str): The system prompt to provide context.
        user_prompt (str): The user prompt to generate the completion.
        client (openai.OpenAI): The OpenAI API client.
        model (str): The model to use for the completion.
        max_tokens (int): The maximum number of tokens to generate.

    Returns:
        HDSData: The parsed HDS data object from the completion.

    """
    
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
    # print(hds_data.model_dump())
    return hds_data


def flatten_hds_data(hds_dict: dict, retc_data: dict, gei_data: dict) -> list:
    """
    Flattens the HDS data into a list of dictionaries for Excel export.

    Args:
        hds_dict (dict): The HDS data dictionary.
        retc_data (dict): The RETC data dictionary.
        gei_data (dict): The GEI data dictionary.

    Returns:
        list: A list of dictionaries with flattened data for each substance.
    """
    sustancias_flat = []

    # Información común para todas las filas de la sustancia
    common_info = {
        'Archivo': hds_dict.get('Archivo'),
        'Nombre de la Sustancia Química': hds_dict.get('nombre_sustancia_quimica'),
        'Idioma de la HDS': hds_dict.get('idioma'),
        'Estado Físico': hds_dict.get('estado_fisico'),
        'Sujeta a RETC': hds_dict.get('sujeta_retc'),
        'Sujeta a GEI': hds_dict.get('sujeta_gei'),
        'pH de la sustancia': hds_dict.get('ph'),
        'Palabra de Advertencia': hds_dict.get('palabra_advertencia'),
        'Indicaciones de toxicología': hds_dict.get('indicaciones_toxicologia'),
        'Pictogramas': ', '.join(hds_dict.get('pictogramas', [])),
        'Olor': hds_dict.get('olor'),
        'Color': hds_dict.get('color'),
        'Propiedades Explosivas': hds_dict.get('propiedades_explosivas'),
        'Propiedades Comburentes': hds_dict.get('propiedades_comburentes'),
        'Tamaño de Partícula': hds_dict.get('tamano_particula'),
    }

    # Desanidar propiedades numéricas
    propiedades = [
        'temperatura_ebullicion',
        'punto_congelacion',
        'densidad',
        'punto_inflamacion',
        'solubilidad_agua',
        'velocidad_evaporacion',
        'presion_vapor'
    ]
    for prop in propiedades:
        propiedad = hds_dict.get(prop)
        if propiedad:
            common_info[f'{prop} (Valor)'] = propiedad.get('valor')
            common_info[f'{prop} (Unidades)'] = propiedad.get('unidades')
        else:
            common_info[f'{prop} (Valor)'] = None
            common_info[f'{prop} (Unidades)'] = None

    # Límites de inflamabilidad
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

            # Verificar si el número CAS está en la tabla RETC
            cas_number = componente.get('numero_cas')
            nombre_componente = componente.get('nombre')

            # Buscar en RETC por número CAS
            retc_entry = retc_data.get(cas_number)

            # Si no se encuentra por CAS, intentar buscar por nombre usando coincidencia difusa
            if not retc_entry and nombre_componente:
                # Crear una lista de nombres comunes de RETC
                retc_nombres = [entry['componente_retc'] for entry in retc_data.values()]
                # Realizar coincidencia difusa
                match = process.extractOne(nombre_componente, 
                                           retc_nombres, 
                                           scorer=fuzz.token_sort_ratio)
                if match and match[1] >= 80:  # Umbral de similitud del 80%
                    # Obtener la entrada correspondiente
                    for entry in retc_data.values():
                        if entry['componente_retc'] == match[0]:
                            retc_entry = entry
                            break

            # Asignar datos de RETC si se encontró una coincidencia
            if retc_entry:
                row['Componente RETC'] = retc_entry.get('componente_retc')
                row['MPU'] = retc_entry.get('mpu')
                row['Emision Transferencia'] = retc_entry.get('emision_transferencia')
            else:
                row['Componente RETC'] = None
                row['MPU'] = None
                row['Emision Transferencia'] = None

            # Hacer lo mismo para GEI
            # Buscar en GEI por número CAS
            gei_entry = gei_data.get(cas_number)

            # Si no se encuentra por CAS, intentar buscar por nombre usando coincidencia difusa
            if not gei_entry and nombre_componente:
                gei_nombres = [entry['nombre_comun'] for entry in gei_data.values()]
                match = process.extractOne(nombre_componente,
                                            gei_nombres,
                                            scorer=fuzz.token_sort_ratio)
                if match and match[1] >= 90:
                    print(f"GEI match between {nombre_componente} and {match[0]}: {match[1]}")
                    for entry in gei_data.values():
                        if entry['nombre_comun'] == match[0]:
                            gei_entry = entry
                            break

            if gei_entry:
                row['Componente GEI'] = gei_entry.get('nombre_comun')
                row['Potencial de Calentamiento Global'] = gei_entry.get('PCG')
            else:
                row['Componente GEI'] = None
                row['Potencial de Calentamiento Global'] = None

            sustancias_flat.append(row)
    else:
        # Si no hay componentes, agregar la sustancia sin detalles de componentes
        row = common_info.copy()
        # Agregar campos vacíos
        row['Nombre del Componente'] = None
        row['Número CAS del Componente'] = None
        row['Porcentaje del Componente'] = None
        row['Componente RETC'] = None
        row['MPU'] = None
        row['Emision Transferencia'] = None
        row['Componente GEI'] = None
        row['Potencial de Calentamiento Global'] = None
        sustancias_flat.append(row)

    # Manejar Indicaciones de peligro H y Consejos de Prudencia P
    indicaciones_h = hds_dict.get('identificaciones_peligro_h', [])
    consejos_p = hds_dict.get('consejos_prudencia_p', [])

    # Concatenar códigos y descripciones
    common_info['Identificaciones de peligro H'] = '; '.join(
        [f"{i['codigo']}: {i['descripcion']}" for i in indicaciones_h])
    common_info['Consejos de Prudencia P'] = '; '.join(
        [f"{i['codigo']}: {i['descripcion']}" for i in consejos_p])

    # Actualizar las filas existentes con las indicaciones
    for row in sustancias_flat:
        row.update({
            'Identificaciones de peligro H': common_info['Identificaciones de peligro H'],
            'Consejos de Prudencia P': common_info['Consejos de Prudencia P']
        })

    return sustancias_flat



def export_to_excel(data:list, output_file:str ="output.xlsx")->None:
    """
    Exporta los datos de las sustancias a un archivo Excel con celdas combinadas y alineadas.

    Args:
        data (list): La lista de datos de las sustancias.
        output_file (str): La ruta del archivo de salida Excel.


    """

    # Crear DataFrame de pandas
    df = pd.DataFrame(data)

    # Define the desired order of columns
    column_order = [
        'Archivo',                         # 1. File name
        'Nombre de la Sustancia Química',   # 2. Substance name
        'Idioma de la HDS',                 # 3. Language
        'Nombre del Componente',            # 4. Component name
        'Sujeta a GEI',                     # 5. GEI component
        'Sujeta a RETC',                    # 6. RETC component
        'Número CAS del Componente',        # 7. Component CAS number
        'Porcentaje del Componente',        # 8. Component percentage
        'Componente GEI',                   # 9. GEI component
        'Potencial de Calentamiento Global',# 10. Global Warming Potential
        'Componente RETC',                  # 11. RETC component
        'MPU',                              # 12. MPU
        'Emision Transferencia',            # 13. Emission transfer
        'Palabra de Advertencia',           # 14. Warning word
        'Indicaciones de toxicología',      # 15. Toxicology indications
        'Pictogramas',                      # 16. Pictograms
        'Estado Físico',                    # 17. Physical state
        'Identificaciones de peligro H',    # 17. Hazard identifications H
        'Consejos de Prudencia P',          # 18. Precautionary statements P
        'Olor',                             # 19. Odor
        'Color',                            # 20. Color
        'pH de la sustancia',               # 21. pH
        'temperatura_ebullicion (Valor)',   # 22. Boiling temperature (Value)
        'temperatura_ebullicion (Unidades)',# 23. Boiling temperature (Units)
        'punto_congelacion (Valor)',        # 24. Freezing point (Value)
        'punto_congelacion (Unidades)',     # 25. Freezing point (Units)
        'densidad (Valor)',                 # 26. Density (Value)
        'densidad (Unidades)',              # 27. Density (Units)
        'punto_inflamacion (Valor)',        # 28. Flash point (Value)
        'punto_inflamacion (Unidades)',     # 29. Flash point (Units)
        'velocidad_evaporacion (Valor)',    # 30. Evaporation rate (Value)
        'velocidad_evaporacion (Unidades)', # 31. Evaporation rate (Units)
        'presion_vapor (Valor)',            # 32. Vapor pressure (Value)
        'presion_vapor (Unidades)',         # 33. Vapor pressure (Units)
        'solubilidad_agua (Valor)',         # 34. Water solubility (Value)
        'solubilidad_agua (Unidades)',      # 35. Water solubility (Units)
        'Propiedades Explosivas',           # 36. Explosive properties
        'Propiedades Comburentes',          # 37. Oxidizing properties
        'Tamaño de Partícula',              # 38. Particle size
        'Límite inferior de inflamabilidad',# 39. Lower flammability limit
        'Límite superior de inflamabilidad'# 40. Upper flammability limit        
    ]

    # Reorder the DataFrame columns
    df = df[column_order]

    # Save the DataFrame to a temporary Excel file
    temp_file = "temp_output.xlsx"
    df.to_excel(temp_file, index=False)

    # Load the Excel file with openpyxl for customization
    wb = load_workbook(temp_file)
    ws = wb.active

    # Merge cells for specified columns only if the substance is the same
    merge_columns = [
        'Archivo',
        'Nombre de la Sustancia Química',
        'Idioma de la HDS',
        'Sujeta a GEI',
        'Sujeta a RETC',
        'Palabra de Advertencia',
        'Indicaciones de toxicología',
        'Pictogramas',
        'Estado Físico',
        'Identificaciones de peligro H',
        'Consejos de Prudencia P',
        'Olor',
        'Color',
        'pH de la sustancia',
        'temperatura_ebullicion (Valor)',
        'temperatura_ebullicion (Unidades)',
        'punto_congelacion (Valor)',
        'punto_congelacion (Unidades)',
        'densidad (Valor)',
        'densidad (Unidades)',
        'punto_inflamacion (Valor)',
        'punto_inflamacion (Unidades)',
        'velocidad_evaporacion (Valor)',
        'velocidad_evaporacion (Unidades)',
        'presion_vapor (Valor)',
        'presion_vapor (Unidades)',
        'solubilidad_agua (Valor)',
        'solubilidad_agua (Unidades)',
        'Propiedades Explosivas',
        'Propiedades Comburentes',
        'Tamaño de Partícula',
        'Límite inferior de inflamabilidad',
        'Límite superior de inflamabilidad',
    ]

    # Initialize dictionaries to track current values and start rows for merging
    current_values = {col: None for col in merge_columns}
    start_rows = {col: 2 for col in merge_columns}  # Assume starting at row 2 (after header)
    current_sustancia = None


    for row in range(2, ws.max_row + 1):
        sustancia_value = ws[f'B{row}'].value  # Column B for Substance Name
        if sustancia_value != current_sustancia:
            # If the substance changes, merge previous cells
            for col_name in merge_columns:
                if current_values[col_name] is not None:
                    ws.merge_cells(start_row=start_rows[col_name], start_column=df.columns.get_loc(col_name) + 1, 
                                   end_row=row - 1, end_column=df.columns.get_loc(col_name) + 1)
                current_values[col_name] = ws[f'{get_column_letter(df.columns.get_loc(col_name) + 1)}{row}'].value
                start_rows[col_name] = row
            current_sustancia = sustancia_value
            start_row_sustancia = row
        else:
            # If the substance does not change, check columns for merging
            for col_name in merge_columns:
                col_letter = get_column_letter(df.columns.get_loc(col_name) + 1)  # Get column letter
                cell_value = ws[f'{col_letter}{row}'].value

                if cell_value == current_values[col_name]:
                    ws[f'{col_letter}{row}'].value = None  # Clear duplicate cells
                else:
                    # If a different value is detected, merge previous cells
                    if current_values[col_name] is not None:
                        ws.merge_cells(start_row=start_rows[col_name], start_column=df.columns.get_loc(col_name) + 1, 
                                       end_row=row - 1, end_column=df.columns.get_loc(col_name) + 1)
                    current_values[col_name] = cell_value
                    start_rows[col_name] = row

    # Merge cells for the last group of the same substance
    for col_name in merge_columns:
        if current_values[col_name] is not None:
            ws.merge_cells(start_row=start_rows[col_name], start_column=df.columns.get_loc(col_name) + 1, 
                           end_row=ws.max_row, end_column=df.columns.get_loc(col_name) + 1)

    # Set the column widths as expected
    fixed_column_widths = {
        'A': 10,   # Archivo
        'B': 20,   # Nombre de la Sustancia Química
        'C': 5,    # Idioma de la HDS
        'D': 15,   # Nombre del Componente
        'E': 12,   # Número CAS del Componente
        'F': 10,   # Porcentaje del Componente
        'G': 15,   # Componente GEI
        'H': 8,   # Potencial de Calentamiento Global
        'I': 12,   # Componente RETC
        'J': 8,   # MPU
        'K': 8,   # Emisión Transferencia
        'L': 12,   # Palabra de Advertencia
        'M': 25,   # Indicaciones de toxicología
        'N': 20,   # Pictogramas
        'O': 15,   # Estado Físico
        'P': 30,   # Identificaciones de peligro H
        'Q': 30,   # Consejos de Prudencia P
        'R': 15,   # Olor
        'S': 15,   # Color
        'T': 8,    # pH de la sustancia
        'U': 12,   # temperatura_ebullicion (Valor)
        'V': 12,   # temperatura_ebullicion (Unidades)
        'W': 12,   # punto_congelacion (Valor)
        'X': 12,   # punto_congelacion (Unidades)
        'Y': 12,   # densidad (Valor)
        'Z': 12,   # densidad (Unidades)
        'AA': 12,   # punto_inflamacion (Valor)
        'AB': 12,  # punto_inflamacion (Unidades)
        'AC': 12,  # velocidad_evaporacion (Valor)
        'AD': 12,  # velocidad_evaporacion (Unidades)
        'AE': 12,  # presion_vapor (Valor)
        'AF': 12,  # presion_vapor (Unidades)
        'AG': 12,  # solubilidad_agua (Valor)
        'AH': 12,  # solubilidad_agua (Unidades)
        'AI': 20,  # Propiedades Explosivas
        'AJ': 20,  # Propiedades Comburentes
        'AK': 15,  # Tamaño de Partícula
        'AL': 12,  # Límite inferior de inflamabilidad
        'AM': 12   # Límite superior de inflamabilidad
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
    ##Eliminar el archivo temporal
    os.remove(temp_file)


def process_hds_folder(folder_path:str, project_name:str, client, excel_output:str, json_output:str, progress_callback=None):
    """
    Procesa una carpeta de archivos PDF de Hojas de Datos de Seguridad (HDS) y genera un archivo Excel con los datos extraídos.

    Args:
        folder_path (str): La ruta de la carpeta que contiene los archivos PDF.
        project_name (str): El nombre del proyecto o carpeta.
        client (openai.OpenAI): El cliente de OpenAI API.
        excel_output (str): La ruta del archivo Excel de salida.
        json_output (str): La ruta del archivo JSON de salida.
        progress_callback (function): Una función de callback para informar del progreso.

    Returns:
        tuple: Una tupla con una lista de errores y una lista de todas las sustancias procesadas.
    """
    # Cargar la tabla RETC
    # Determine base path
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    

    retc_file_path = os.path.join(base_path,'data_sets', 'retc_table.json')
    print("RETC FILE PATH: ",retc_file_path)

    gei_file_path = os.path.join(base_path,'data_sets', 'gei_table.json')
    print("GEI FILE PATH: ",gei_file_path)

    with open(retc_file_path, 'r', encoding='utf-8') as f:
        retc_list = json.load(f)
        # Convertir la lista de dicts a un dict con el número CAS como clave para acceso rápido
        retc_data = {entry['numero_cas']: entry for entry in retc_list}
    with open(gei_file_path, 'r', encoding='utf-8') as f:
        gei_list = json.load(f)
        # Crear un diccionario con número CAS como clave
        gei_data_cas = {entry['CAS']: entry for entry in gei_list}
        # Crear un diccionario con nombre común como clave
        gei_data_nombre = {entry['nombre_comun']: entry for entry in gei_list}
        # Combinar ambos
        gei_data = {**gei_data_cas, **gei_data_nombre}
    all_sustancias = []
    errores = []
    sustancias_data = []  # Lista para almacenar los datos procesados para Excel
    progress_percent = 0

    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    total_files = len(pdf_files)
    # progress_callback(progress_percent, pdf_files[0] if pdf_files else "")

    for index, filename in enumerate(pdf_files):
        
        
        file_path = os.path.join(folder_path, filename)
         # Calcular el progreso y llamar al callback
        if progress_callback:
            progress_callback(progress_percent, filename)

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
            sustancias_data.extend(flatten_hds_data(hds_dict, retc_data,gei_data))

            print(f"Procesado correctamente: {filename}")
            progress_percent = (index + 1) * 100 // total_files
            all_sustancias.append(hds_dict)
        except Exception as e:
            print(f"Error al procesar {filename}: {e}")
            errores.append(filename)


    # Guardar el JSON final con todas las sustancias
    with open(json_output, 'w', encoding='utf-8') as output_file:
        json.dump(all_sustancias, output_file, indent=2, ensure_ascii=False)

    # Generar el Excel formateado con celdas combinadas y alineadas
    export_to_excel(sustancias_data, excel_output)

    print(f"Proceso completado. Los archivos {json_output} y {excel_output} han sido generados.")
    return errores, all_sustancias

if __name__ == "__main__":
    load_dotenv()
    openai.api_key = os.getenv("OPENAI_API_KEY")
    example_path="ejemplo/"
    client = openai.OpenAI()
    process_hds_folder(example_path, "Prueba GEI y RETC",client,'ejemplo/prueba2.xlsx','ejemplo/prueba2.json')
