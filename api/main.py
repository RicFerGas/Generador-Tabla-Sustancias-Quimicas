#main.py
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import os
from dotenv import load_dotenv

from api.routes import preprocess, extract, create_table

# Load environment variables
load_dotenv()

app = FastAPI(
    title="HDS Processing API",
    description="API for processing Safety Data Sheets (HDS)",
    version="1.0.0"
)

# Configure CORS for frontend access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Adjust this in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(preprocess.router)
app.include_router(extract.router)
app.include_router(create_table.router)

@app.get("/")
async def root():
    return {"message": "Welcome to HDS Processing API"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api.main:app", host="0.0.0.0", port=8000, reload=True)