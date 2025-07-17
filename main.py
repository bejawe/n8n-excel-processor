import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy
from openpyxl.styles import Font

app = FastAPI()

# --- Helper Functions (No Changes Here) ---
def copy_cell_with_style(src, dst):
    # ... (same as before)
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

def find_last_schedule_row(worksheet, start_row, column_to_check):
    # ... (same as before)
    for row_num in range(worksheet.max_row, start_row - 1, -1):
        cell_value = worksheet.cell(row=row_num, column=column_to_check).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num
    # If no schedule is found, return 0 to indicate the template is empty
    return 0

# ---------- Main API Endpoint ----------
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    # ... (start of function is the same)
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- 1. Load & 2. Choose Sheet (Same as before) ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)
        sheet_name = panel_data.get("projectName", "Default")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # *** THIS IS THE FINAL, CORRECTED LOGIC ***
        
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        TEMPLATE_HEIGHT = TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1
        
        last_schedule_row = find_last_schedule_row(ws, start_row=TEMPLATE_START_ROW, column_to_check=3)
        
        is_first_panel = (last_schedule_row == 0)
        
        if is_first_panel:
            # This is the first panel. Write directly into the template.
            write_row = TEMPLATE_START_ROW
            print("INFO: First panel. Writing directly into template.")
        else:
            # This is a subsequent panel. Insert, copy, and then write.
            insertion_row = last_schedule_row + 1
            source_start_row = last_schedule_row - TEMPLATE_HEIGHT + 1
            
            # Insert blank rows
            ws.insert_rows(insertion_row, amount=TEMPLATE_HEIGHT)
            
            # Copy previous schedule into the new space
            for r_offset in range(TEMPLATE_HEIGHT):
                for c in range(1, 13):
                    src_cell = ws.cell(row=source_start_row + r_offset, column=c)
                    dst_cell = ws.cell(row=insertion_row + r_offset, column=c)
                    copy_cell_with_style(src_cell, dst_cell)
            
            write_row = insertion_row
            print(f"INFO: Subsequent panel. Inserting new block at row {write_row}.")

        # --- Write Panel Data (This logic is now universal) ---
        row = write_row
        
        ws.cell(row=row, column=4).value = panel_data.get("panelName")
        
        link_cell = ws.cell(row=row + 11, column=1)
        source_image_url = panel_data.get("sourceImageUrl")
        if source_image_url:
            link_cell.value = "panel image"
            link_cell.hyperlink = source_image_url
            link_cell.font = Font(color="0000FF", underline="single")
        
        # (The rest of your data-writing logic remains the same)

        # --- Save and Return (Same as before) ---
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        
        original_filename = file.filename
        if original_filename.lower().endswith('.xlsm'):
            media_type = 'application/vnd.ms-excel.sheet.macroenabled.12'
        else:
            media_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            
        return StreamingResponse(
            out,
            media_type=media_type,
            headers={"Content-Disposition": f"attachment; filename={original_filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred during Excel processing: {str(e)}")
