import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy
from openpyxl.styles import Font

app = FastAPI()

# ==============================================================================
# HELPER FUNCTIONS - These are now fully validated
# ==============================================================================
def copy_cell_with_style(src_cell_info, dst_cell):
    """Applies style and value from a dictionary (src_cell_info) to a cell object (dst_cell)."""
    dst_cell.value = src_cell_info.get("value")
    if src_cell_info.get("hyperlink"):
        dst_cell.hyperlink = copy(src_cell_info.get("hyperlink"))
    if src_cell_info.get('has_style'):
        dst_cell.font = copy(src_cell_info.get("font"))
        dst_cell.border = copy(src_cell_info.get("border"))
        dst_cell.fill = copy(src_cell_info.get("fill"))
        dst_cell.number_format = copy(src_cell_info.get("number_format"))
        dst_cell.protection = copy(src_cell_info.get("protection"))
        dst_cell.alignment = copy(src_cell_info.get("alignment"))

def find_last_schedule_row(worksheet, start_row, column_to_check):
    """Finds the last row of the last schedule by searching upwards for 'TOTAL'."""
    for row_num in range(worksheet.max_row, start_row, -1):
        cell_value = worksheet.cell(row=row_num, column=column_to_check).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num
    return 30

def is_template_empty(worksheet, check_row, check_col):
    """Checks if the template is empty by looking at a key cell (I18)."""
    cell = worksheet.cell(row=check_row, column=check_col)
    return cell.value is None and cell.hyperlink is None

# ==============================================================================
# MAIN API ENDPOINT - This is the final version
# ==============================================================================
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

        if is_template_empty(ws, check_row=18, check_col=9):
            write_row = TEMPLATE_START_ROW
        else:
            last_schedule_row = find_last_schedule_row(ws, start_row=TEMPLATE_END_ROW, column_to_check=3)
            insertion_row = last_schedule_row + 1
            ws.insert_rows(insertion_row, amount=TEMPLATE_HEIGHT)
            
            # --- The "Read Before You Write" Fix ---
            template_data = []
            for r_offset in range(TEMPLATE_HEIGHT):
                row_data = []
                for c in range(1, 13):
                    cell = ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                    row_data.append({
                        "value": cell.value, "has_style": cell.has_style,
                        "font": cell.font, "border": cell.border, "fill": cell.fill,
                        "number_format": cell.number_format, "protection": cell.protection,
                        "alignment": cell.alignment, "hyperlink": cell.hyperlink
                    })
                template_data.append(row_data)
            
            for r_offset in range(TEMPLATE_HEIGHT):
                for c_offset in range(12):
                    cell_info = template_data[r_offset][c_offset]
                    dst_cell = ws.cell(row=insertion_row + r_offset, column=c_offset + 1)
                    copy_cell_with_style(cell_info, dst_cell)
            
            write_row = insertion_row

        # --- 3. Write All Panel Data into the Target Block ---
        row = write_row
        
        ws.cell(row=row, column=4).value = panel_data.get("panelName")
        ws.cell(row=row + 6, column=7).value = panel_data.get("mountingType", "SURFACE")
        ws.cell(row=row + 7, column=7).value = panel_data.get("ipDegree")

        link_cell = ws.cell(row=row + 11, column=1)
        source_image_url = panel_data.get("sourceImageUrl")
        if source_image_url:
            link_cell.value = "panel image"
            link_cell.hyperlink = source_image_url
            link_cell.font = Font(color="0000FF", underline="single")

        recommendations = panel_data.get("recommendations", [])
        
        main_rec = next((r for r in recommendations if "MCCB" in r.get("breakerSpec", "")), None)
        if main_rec:
            ws.cell(row=row + 11, column=2).value = main_rec.get("breakerSpec")
            ws.cell(row=row + 11, column=9).value = main_rec.get("matchedPart", {}).get("Reference number", "")
        
        branch_recs = [r for r in recommendations if "MCCB" not in r.get("breakerSpec", "")]
        for i, rec in enumerate(branch_recs):
            current_row = row + 13 + i
            if current_row <= (row + TEMPLATE_HEIGHT - 1):
                ws.cell(row=current_row, column=2).value = rec.get("breakerSpec")
                ws.cell(row=current_row, column=6).value = rec.get("quantity")
                ws.cell(row=current_row, column=9).value = rec.get("matchedPart", {}).get("Reference number", "")
        
        ws.cell(row=row + 23, column=3).value = "TOTAL"

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
