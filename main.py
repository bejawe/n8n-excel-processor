import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy

app = FastAPI()

# ---------- helper ----------
def copy_cell_with_style(src, dst):
    """Copies value and all style attributes from source_cell to target_cell."""
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

# ---------- endpoint ----------
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    """
    Receives an Excel file and JSON data, modifies the Excel file by appending
    a new panel block based on a template, writes the panel data into the
    new block, and returns the modified file.
    """
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- 1. Load JSON & Excel ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)

        # --- 2. Choose sheet ---
        # Tries to find a sheet with the given projectName, otherwise uses the active one.
        sheet_name = panel_data.get("projectName", "")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        print(f"DEBUG: sheet chosen = {ws.title}")

        # --- 3. Define template bounds & find next available row ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        next_row = ws.max_row + 2
        print(f"DEBUG: next_row for new panel = {next_row}")

        # --- 4. Copy the template block ---
        for r_offset in range(TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1):
            for c in range(1, 13): # Assuming columns A to L
                src = ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                dst = ws.cell(row=next_row + r_offset, column=c)
                copy_cell_with_style(src, dst)

        # --- 5. Write panel data into the new block ---
        row = next_row  # The starting row of our new block

        # Panel metadata
        ws.cell(row=row, column=4).value = panel_data.get("panelName")
        ws.cell(row=row + 11, column=1).value = panel_data.get("sourceImageUrl")
        ws.cell(row=row + 6, column=7).value = panel_data.get("mountingType", "SURFACE")
        ws.cell(row=row + 7, column=7).value = panel_data.get("ipDegree")

        recommendations = panel_data.get("recommendations", [])

        # Main breaker (MCCB)
        main_rec = next((r for r in recommendations if "MCCB" in r.get("breakerSpec", "")), None)
        if main_rec:
            ws.cell(row=row + 11, column=2).value = main_rec.get("breakerSpec")
            ws.cell(row=row + 11, column=9).value = main_rec.get("matchedPart", {}).get("Reference number", "")

        # Branch breakers (non-MCCB)
        branch_recs = [r for r in recommendations if "MCCB" not in r.get("breakerSpec", "")]
        for i, rec in enumerate(branch_recs):
            current_row = row + 13 + i  # Start writing branch breakers at an offset
            ws.cell(row=current_row, column=2).value = rec.get("breakerSpec")
            ws.cell(row=current_row, column=6).value = rec.get("quantity")
            ws.cell(row=current_row, column=9).value = rec.get("matchedPart", {}).get("Reference number", "")
        
        print("DEBUG: Finished writing panel data.")

        # --- 6. Save and return the modified file ---
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=modified_{file.filename}"}
        )

    except Exception as e:
        # If any error occurs, return a 500 status with the error details
        print(f"ERROR: {str(e)}")
        raise HTTPException(status_code=500, detail=f"An error occurred during Excel processing: {str(e)}")
