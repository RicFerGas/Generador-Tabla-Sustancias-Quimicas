import os
import openai

def get_openai_client(api_key: str=None):
    """Dependency to get the OpenAI client."""
    openai_api_key = api_key if api_key else os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        raise ValueError("OpenAI API key not found in environment variables")
    try:
        client = openai.OpenAI(api_key=openai_api_key)
    except Exception as e:
        raise ValueError(f"Error creating OpenAI client: {str(e)}")
    return client