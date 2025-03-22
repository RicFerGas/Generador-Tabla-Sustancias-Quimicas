#api/routes/extract.py
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from api.services.extract_info import extract_info_from_hds_txt
from schemas import HDSData
from api.dependancies import get_openai_client
router = APIRouter(prefix="/extract", tags=["extract"])

class TextExtractionRequest(BaseModel):
    text: str
    api_key: str
    

@router.post("/", response_model=HDSData)
async def extract_hds_info(
    request: TextExtractionRequest
):
    """
    Extracts structured data from HDS document text
    """
    try:
        client = get_openai_client(api_key=request.api_key)
        extracted_data = extract_info_from_hds_txt(request.text, client)
        return extracted_data
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error extracting data: {str(e)}")
