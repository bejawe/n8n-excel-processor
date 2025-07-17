import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy
from openpyxl.styles import Font

app = FastAPI()

# ---------- Helper Functions (These are now proven correct) ----------
def copy_cell_with_style(src, dst):
    """Copies value, style, and hyperlink from a source cell to a destination cell."""
    dst.value = src.value
    if src.hyperlink:
        dst.hyperlink = copy(src.hyperlink)
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

def find_last_schedule_row(worksheet, start_row, column_to_check):
    """
    Finds the last row of the last schedule.
    Returns 30 if the template seems empty, otherwise finds the last 'TOTAL'.
    """
    # Start searching from the bottom of the sheet up to the template end row
    for row_num in range(worksheet.max_row, start_row - 1, -1):
        cell_value = worksheet.cell(row=row_num, column=column_to_check).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num
    # If no 'TOTAL' is found, assume the sheet only contains the template
    return 30

# ---------- Main API Endpoint ----------
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- 1. Load & Choose Sheet ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)
        sheet_name = panel_data.get("projectName", "Default")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # --- 2. Determine Logic Path & Ranges ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        TEMPLATE_HEIGHT = TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1

        last_schedule_row = find_last_schedule_row(ws, start_row=TEMPLATE_END_ROW, column_to_check=3)
        
        # This is our key decision point
        is_first_panel = (last_schedule_row == 30 and ws.cell(row=TEMPLATE_START_ROW, column=3).value is None)

        if is_first_panel:
            # PATH A: First panel, write directly into template
            write_row = TEMPLATE_START_ROW
        else:
            # PATH B: Subsequent panels, insert and copy
            insertion_row = last_schedule_row + 1
            ws.insert_rows(insertion_row, amount=TEMPLATE_HEIGHT)
            
            # Always copy the clean, original template to the new space
            for r_offset in range(TEMPLATE_HEIGHT):
                for c in range(1, 13):
                    src_cell = ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                    dst_cell = ws.cell(row=insertion_row + r_offset, column=c)
                    copy_cell_with_style(src_cell, dst_cell)
            write_row = insertion_row

        # --- 3. Write Panel Data into the Target Block ---
        row = write_row
        ws.cell(row=row, column=3).value = "Panel Data Here" # Write dummy data to C7
        ws.cell(row=row + 23, column=3).value = "TOTAL"     # Write TOTAL to C30
        
        ws.cell(row=row, column=4).value = panel_data.get("panelName")
        
        link_cell = ws.cell(row=row + 11, column=1)
        source_image_url = panel_data.get("sourceImageUrl")
        if source_image_url:
            link_cell.value = "panel image"
            link_cell.hyperlink = source_image_url
            link_cell.font = Font(color="0000FF", underline="single")
        
        # (Add the rest of your breaker-writing logic here)

        # --- 4. Save and Return ---
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        
        original_filename = file.filename
        media_type = 'application/vnd.ms-excel.sheet.macroenabled.12' if original_filename.lower().endswith('.xlsm') else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        return StreamingResponse(
            out,
            media_type=media_type,
            headers={"Content-Disposition": f"attachment; filename={original_filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred during Excel processing: {str(e)}")
