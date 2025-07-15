import io
from fastapi import FastAPI, UploadFile, File, HTTPException
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
@app.post("/process-excel-panel/")
async def process_panel(file: UploadFile = File(...)):
    """
    Receives an Excel file, appends a new panel block based on a template,
    and returns the modified file.
    """
    # Ensure the uploaded file is an Excel file
    if not file.filename.endswith(('.xlsx', '.xlsm')):
        raise HTTPException(status_code=400, detail="Invalid file format. Please upload an .xlsx or .xlsm file.")

    try:
        # Read the uploaded file directly into memory
        contents = await file.read()
        workbook_stream = io.BytesIO(contents)

        # Load the workbook with openpyxl, preserving macros
        wb = openpyxl.load_workbook(workbook_stream, keep_vba=True)
        
        # --- This assumes your target sheet is the first one ---
        # For a specific sheet: ws = wb["YourSheetName"]
        ws = wb.active

        # --- Define the template area to be copied ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30 # 24 rows total
        TEMPLATE_START_COL = 1 # Column A
        TEMPLATE_END_COL = 12 # Column L
        
        # Find the next empty row to paste the new panel
        next_row = ws.max_row + 2 # Add a little space

        # --- Copy the template block cell by cell ---
        for row_offset in range(TEMPLATE_START_ROW, TEMPLATE_END_ROW + 1):
            for col_offset in range(TEMPLATE_START_COL, TEMPLATE_END_COL + 1):
                source_cell = ws.cell(row=row_offset, column=col_offset)
                target_cell = ws.cell(row=next_row + (row_offset - TEMPLATE_START_ROW), column=col_offset)
                copy_cell_with_style(source_cell, target_cell)

        # (Optional) Here you would add logic to write the new panel data 
        # into the newly created block at 'next_row'. For now, it just copies.

        # Save the modified workbook to an in-memory stream
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0) # Rewind the stream to the beginning

        return StreamingResponse(
            output_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=modified_{file.filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred: {str(e)}")