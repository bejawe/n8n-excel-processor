import io
import json
from fastapi import FastAPI, Request, Header, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy

app = FastAPI()

# --- Shared Cell Copying Logic (no changes here) ---
def copy_cell_with_style(source_cell, target_cell):
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        # ... (rest of the function is the same)
        target_cell.alignment = copy(source_cell.alignment)

# --- The Main API Endpoint (CHANGED) ---
# It now gets the file from the request body and JSON from a header
@app.post("/process-excel-panel/")
async def process_panel(request: Request, x_panel_data: str = Header(...)):
    try:
        # --- 1. Load the data ---
        panel_data = json.loads(x_panel_data) # Load JSON from header
        contents = await request.body() # Load file from raw request body
        
        workbook_stream = io.BytesIO(contents)
        wb = openpyxl.load_workbook(workbook_stream, keep_vba=True)
        
        # --- The rest of your excel logic is the same ---
        sheet_name = panel_data.get("projectName", "")
ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        # ... (all your logic for finding rows, copying cells, and writing data)

        # --- 5. Save and return the modified file ---
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        
        return StreamingResponse(output_stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
