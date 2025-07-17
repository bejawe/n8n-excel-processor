import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy
from openpyxl.styles import Font

app = FastAPI()

# ---------- Helper Function (No changes needed) ----------
def copy_cell_with_style(src_cell_info, dst_cell):
    """
    Applies style and value from a dictionary (src_cell_info) to a cell object (dst_cell).
    """
    dst_cell.value = src_cell_info.get("value")
    dst_cell.font = copy(src_cell_info.get("font"))
    dst_cell.border = copy(src_cell_info.get("border"))
    dst_cell.fill = copy(src_cell_info.get("fill"))
    dst_cell.number_format = src_cell_info.get("number_format")
    dst_cell.protection = copy(src_cell_info.get("protection"))
    dst_cell.alignment = copy(src_cell_info.get("alignment"))
    # The hyperlink is now handled separately in the main logic

def find_last_schedule_row(worksheet, start_row, column_to_check):
    """Finds the last row of the last schedule by searching upwards for 'TOTAL'"""
    for row_num in range(worksheet.max_row, start_row - 1, -1):
        cell_value = worksheet.cell(row=row_num, column=column_to_check).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num
    return 30

def is_template_empty(worksheet, check_row, check_col):
    """Checks if the template is empty by looking at cell I18."""
    # Check both value and hyperlink to be certain
    cell = worksheet.cell(row=check_row, column=check_col)
    return cell.value is None and cell.hyperlink is None


# ---------- Main API Endpoint ----------
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- 1. Load & 2. Choose Sheet ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)
        sheet_name = panel_data.get("projectName", "Default")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # --- 3. Read the Original Template into Memory (The Key Fix) ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        TEMPLATE_HEIGHT = TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1
        
        template_data = []
        for r_offset in range(TEMPLATE_HEIGHT):
            row_data = []
            for c in range(1, 13):
                cell = ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                row_data.append({
                    "value": cell.value,
                    "font": cell.font,
                    "border": cell.border,
                    "fill": cell.fill,
                    "number_format": cell.number_format,
                    "protection": cell.protection,
                    "alignment": cell.alignment
                })
            template_data.append(row_data)

        # --- 4. Determine where to write ---
        if is_template_empty(ws, check_row=18, check_col=9):
            write_row = TEMPLATE_START_ROW
        else:
            last_schedule_row = find_last_schedule_row(ws, start_row=TEMPLATE_START_ROW, column_to_check=3)
            insertion_row = last_schedule_row + 1
            
            # Insert blank rows
            ws.insert_rows(insertion_row, amount=TEMPLATE_HEIGHT)
            
            # Paste the IN-MEMORY template into the new space
            for r_offset in range(TEMPLATE_HEIGHT):
                for c_offset in range(12):
                    dst_cell = ws.cell(row=insertion_row + r_offset, column=c_offset + 1)
                    copy_cell_with_style(template_data[r_offset][c_offset], dst_cell)
            
            write_row = insertion_row

        # --- 5. Write New Panel Data into the target block ---
        row = write_row
        
        # (The rest of your data writing logic is unchanged and will now work correctly)
        ws.cell(row=row, column=4).value = panel_data.get("panelName")
        link_cell = ws.cell(row=row + 11, column=1)
        source_image_url = panel_data.get("sourceImageUrl")
        if source_image_url:
            link_cell.value = "panel image"
            link_cell.hyperlink = source_image_url
            link_cell.font = Font(color="0000FF", underline="single")
        #... (rest of the breaker data)

        # --- 6. Save and Return ---
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
