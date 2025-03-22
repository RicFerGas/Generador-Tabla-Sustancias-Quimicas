from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from api.services.preprocess import DocumentPreprocessor
import os

router = APIRouter(prefix="/preprocess", tags=["preprocess"])

class FilePathRequest(BaseModel):
    file_path: str

class TextExtractionResponse(BaseModel):
    filename: str
    text: str

@router.post("/", response_model=TextExtractionResponse)
async def preprocess_hds(request: FilePathRequest):
    """
    Extracts text from an HDS document file (PDF, DOCX, etc.) using a local file path
    """
    try:
        # Validate that the file exists
        if not os.path.exists(request.file_path):
            raise HTTPException(status_code=404, detail=f"File not found: {request.file_path}")
        
        # Get just the filename from the path
        filename = os.path.basename(request.file_path)
        
        # Process the file
        doc_preprocessor = DocumentPreprocessor()
        text = doc_preprocessor.extract_text(request.file_path)
        
        return TextExtractionResponse(
            filename=filename,
            text=text
        )
    except Exception as e:
        if isinstance(e, HTTPException):
            raise e
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")