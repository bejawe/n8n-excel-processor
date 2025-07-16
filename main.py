import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy

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

# --- Main endpoint (UploadFile + Form) ---
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    if not file.filename.endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # 1. load JSON
        panel_data = json.loads(panel_data_json)

        # 2. load Excel
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)

        # 3. pick sheet
        sheet_name = panel_data.get("projectName", "")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # ----------------------------------------------------------
        # 4. YOUR ORIGINAL EXCEL LOGIC GOES HERE
        # ----------------------------------------------------------
        # Example template offsets
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        next_row = ws.max_row + 2

        # copy template block
        for src_row in range(TEMPLATE_START_ROW, TEMPLATE_END_ROW + 1):
            for col in range(1, 13):
                src_cell = ws.cell(row=src_row, column=col)
                dst_cell = ws.cell(
                    row=next_row + (src_row - TEMPLATE_START_ROW),
                    column=col
                )
                copy_cell_with_style(src_cell, dst_cell)

        # example: write panel data
        ws.cell(row=next_row, column=4).value = panel_data.get("panelName")
        # … add the rest of your writing logic here …

        # 5. save & return
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)

        return StreamingResponse(
            output_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=modified_{file.filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel processing error: {str(e)}")
