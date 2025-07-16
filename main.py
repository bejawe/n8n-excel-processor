import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy

app = FastAPI()

# (Helper function 'copy_cell_with_style' remains the same)
def copy_cell_with_style(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    try:
        # --- 1. Load JSON & Excel ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)

        # --- 2. Choose sheet ---
        sheet_name = panel_data.get("projectName", "Default")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        print(f"DEBUG: Sheet chosen -> {ws.title}")

        # --- 3. Define template & find next row ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        next_row = ws.max_row + 2
        print(f"DEBUG: Next row for panel -> {next_row}")

        # --- 4. Copy the template block ---
        for r in range(TEMPLATE_START_ROW, TEMPLATE_END_ROW + 1):
            for c in range(1, 13):
                src = ws.cell(row=r, column=c)
                # Calculate destination row correctly
                dst = ws.cell(row=next_row + (r - TEMPLATE_START_ROW), column=c)
                copy_cell_with_style(src, dst)
        print("DEBUG: Template block copied.")

        # --- 5. Write panel data into the new block ---
        row = next_row # Use 'next_row' as the base for the new block
        
        panelName = panel_data.get("panelName")
        ws.cell(row=row, column=4).value = panelName
        print(f"DEBUG: Wrote panelName -> {panelName}")

        sourceImageUrl = panel_data.get("sourceImageUrl")
        ws.cell(row=row + 11, column=1).value = sourceImageUrl
        print(f"DEBUG: Wrote sourceImageUrl -> {sourceImageUrl}")
        
        # (The rest of the write logic for breakers, etc. would go here)
        # For now, we are just testing if the basic writing works.

        # --- 6. Save and return the modified file ---
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        print("DEBUG: Saved workbook to memory. Returning file.")
        return StreamingResponse(
            out,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=modified_{file.filename}"}
        )

    except Exception as e:
        print(f"ERROR: An exception occurred: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
