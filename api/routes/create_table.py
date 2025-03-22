#api/routes/create_table.py
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
import os
from api.services.excel_postprocess import GeneradorTablaSustQ
from schemas import HDSData

router = APIRouter(prefix="/postprocess", tags=["postprocess"])

class ExcelGenerationRequest(BaseModel):
    hds_data_list: List[HDSData]
    filename: str = "hds_data.xlsx"
    

@router.post("/excel")
async def generate_excel_file(request: ExcelGenerationRequest):
    """
    Generates an Excel file from a list of HDS data
    """
    try:
        # Ensure we have an absolute path
        output_directory = "output"
        os.makedirs(output_directory, exist_ok=True)
        
        # Full path to the file
        file_path = os.path.join(output_directory, request.filename)
        
        # Generate the Excel file
        generator = GeneradorTablaSustQ()
        flatten_data = generator.flatten_hds_data(request.hds_data_list)
        generator.export_to_excel_with_template(flatten_data, file_path)
        
        # Return the file as a response
        return FileResponse(
            path=file_path,  # The path parameter is required
            filename=request.filename,  # This will be the download name
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating Excel: {str(e)}")