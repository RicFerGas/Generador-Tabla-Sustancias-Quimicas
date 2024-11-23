# generate_llm_answer.py
from typing import Any
import openai
from pydantic import BaseModel
from schemas import HDSData

def get_completion_openai(system_prompt: str,
                          user_prompt: str,
                          client:Any,
                          response_format:BaseModel,
                          model: str = "gpt-4o-2024-08-06") -> BaseModel:
    """
    Gets the completion from the OpenAI API using the system and user prompts.

    Args:
        system_prompt (str): The system prompt.
        user_prompt (str): The user prompt.
        client (openai.OpenAI): The OpenAI client.
        model (str): The model to use for the completion.
        max_tokens (int): The maximum number of tokens to generate.

    """
    completion = client.beta.chat.completions.parse(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        response_format=response_format,
        max_tokens=4000,
        temperature=0.2,
        top_p=0.9,
    )
    response_object = completion.choices[0].message.parsed
    return response_object
def extract_info_from_hds_txt(hds_text: str,client:Any) -> HDSData:
    """
    Extracts the information from the HDSData object.
    """
    system_prompt="""
    Eres un asistente experto en seguridad industrial que extrae información de Hojas de Datos de Seguridad (HDS) para estudios de riesgo por sustancias químicas. 
    La información extraida debe estar siempre ** EN ESPAÑOL ** y seguir el formato especificado.
    Extrae la información relevante de la HDS proporcionada y devuélvela siguiendo el formato especificado, para el caso de las propiedades numericas,
    no pongas 0 si no ecnuentras la propiedad, el valor default es none."""

    user_prompt = f"""La HDS es la siguiente:\n\n{hds_text}\n\n
    Extrae la información relevante de la HDS proporcionada y devuélvela 
    siguiendo el formato especificado."""
    extracted_data_obj = get_completion_openai(system_prompt, user_prompt, client, HDSData)
    return extracted_data_obj

# Example usage
if __name__ == "__main__":
    import os
    from dotenv import load_dotenv
    from preprocess import DocumentPreprocessor
    load_dotenv()
    openai.api_key = os.getenv("OPENAI_API_KEY")
    CLIENT = openai.OpenAI()
    PATH_JSON = "ejemplo/mini_test/64-Ultra Fine Dry Mold Shield 300.json"
    # HDS_DATA = "Ejemplo de HDS como texto"
    HDS_PATH="ejemplo/mini_test/64-Ultra Fine Dry Mold Shield 300.pdf"
    processor=DocumentPreprocessor()
    HDS_DATA = processor.extract_text(HDS_PATH)
    extracted_data = extract_info_from_hds_txt(HDS_DATA,CLIENT)
    if PATH_JSON:
        with open(PATH_JSON, "w",encoding="utf-8") as f:
            f.write(extracted_data.model_dump_json(indent=2))
    print(extracted_data.model_dump())
    