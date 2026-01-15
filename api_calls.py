# -*- coding: utf-8 -*-
"""
Created on Sat Jan 10 16:44:18 2026

@author: Krishna
"""
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
from pydantic import BaseModel
import tempfile
import os
from typing import Optional
from parser import parse_html, json_to_excel

app = FastAPI()



class TableRequest(BaseModel):
    html: str
    table_id: str
    hyperlink: Optional[str] = None
    alternate_colors: Optional[bool] = False



@app.get("/convert-table-to-excel")
async def convert_table_to_excel(request: TableRequest):
    """
    Endpoint to convert HTML table to Excel file
    
    Request body:
    {
        "html": "<table>...</table>",
        "table_id": "table_13"
        hyperlink: Optional[str] = None - use to insert a link into the first cell
        alternate_colors: Optional[bool] = False - set to True if a white and grey patterning is wanted
    }
    """
    try:
        # Parse the HTML table
        json_data = parse_html(request.html)
        
        # Create temporary file
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        temp_file.close()
        
        hyper_link = request.hyperlink  # None if not provided
        should_alternate = request.alternate_colors  # False if not provided
        json_to_excel(json_data, temp_file.name, hyperlink_url=hyper_link, alternating_colors=should_alternate)
        
        return FileResponse(
            temp_file.name,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename=f"{request.table_id}.xlsx",
            background=BackgroundTask(os.unlink, temp_file.name)
            )
        
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/")
async def root():
    return {"message": "Table to Excel API is running"}


@app.get("/health")
async def health_check():
    return {"status": "healthy"}
