import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy

# Initialize the FastAPI app
app = FastAPI()

# --- Shared Cell Copying Logic ---
def copy_cell_with_style(source_cell, target_cell):
    """Copies value and all style attributes from source_cell to target_cell."""
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# --- The Main API Endpoint ---
# It now accepts a file AND a string of panel data
@app.post("/process-excel-panel/")
async def process_panel(panel_data_json: str = Form(...), file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xlsm')):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- 1. Load the data ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        workbook_stream = io.BytesIO(contents)
        wb = openpyxl.load_workbook(workbook_stream, keep_vba=True)
        ws = wb[panel_data.get("projectName", "")] # Get sheet by project name

        if not ws:
            raise HTTPException(status_code=404, detail=f"Sheet '{panel_data.get('projectName')}' not found in Excel file.")

        # --- 2. Define Template and Find Next Row ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        TEMPLATE_START_COL = 1
        TEMPLATE_END_COL = 12
        start_row = ws.max_row + 2

        # --- 3. Copy the template block ---
        for row_idx in range(TEMPLATE_START_ROW, TEMPLATE_END_ROW + 1):
            for col_idx in range(TEMPLATE_START_COL, TEMPLATE_END_COL + 1):
                source_cell = ws.cell(row=row_idx, column=col_idx)
                target_cell = ws.cell(row=start_row + (row_idx - TEMPLATE_START_ROW), column=col_idx)
                copy_cell_with_style(source_cell, target_cell)
        
        # --- 4. Write the NEW data into the copied block ---
        # This logic is adapted from your n8n 'Code1' node
        recommendations = panel_data.get("recommendations", [])
        main_breaker_rec = next((r for r in recommendations if r['breakerSpec'].startswith('MCCB')), None)
        branch_breaker_recs = [r for r in recommendations if not r['breakerSpec'].startswith('MCCB')]

        # Offsets are based on a template starting at row 7
        ws.cell(row=start_row, column=4).value = panel_data.get("panelName") # D column (4)
        ws.cell(row=start_row + 11, column=1).value = panel_data.get("sourceImageUrl") # A column (1), offset 11 (row 18)
        ws.cell(row=start_row + 6, column=7).value = panel_data.get("mountingType", "SURFACE") # G column (7)
        ws.cell(row=start_row + 7, column=7).value = panel_data.get("ipDegree") # G column (7)
        ws.cell(row=start_row + 4, column=4).value = panel_data.get("shortCircuitCurrentRating", "10 kA") # D column (4)
        
        if main_breaker_rec:
            ws.cell(row=start_row + 11, column=2).value = main_breaker_rec.get("breakerSpec", "") # B column (2)
            ws.cell(row=start_row + 11, column=9).value = main_breaker_rec.get("matchedPart", {}).get("Reference number", "no recommendation") # I column (9)

        if branch_breaker_recs:
            for i, rec in enumerate(branch_breaker_recs):
                row = start_row + 13 + i # Starts at offset 13 (row 20)
                ws.cell(row=row, column=2).value = rec.get("breakerSpec", "") # B column
                ws.cell(row=row, column=6).value = rec.get("quantity", 0) # F column
                ws.cell(row=row, column=9).value = rec.get("matchedPart", {}).get("Reference number", "no recommendation") # I column

        # --- 5. Save and return the modified file ---
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        
        return StreamingResponse(output_stream, media_type="application/vnd.ms-excel.sheet.macroenabled.12")

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
