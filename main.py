import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy
from openpyxl.styles import Font # <-- ADD THIS IMPORT FOR STYLING

app = FastAPI()

# ---------- Helper Functions ----------

def copy_cell_with_style(src, dst):
    # ... (this function remains the same)
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

def find_last_schedule_row(worksheet, start_row, column_to_check):
    # ... (this function remains the same)
    for row_num in range(worksheet.max_row, start_row - 1, -1):
        cell_value = worksheet.cell(row=row_num, column=column_to_check).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num
    return 30


# ---------- Main API Endpoint ----------

@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    # ... (the start of the function remains the same)
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # ... (steps 1-5 remain the same: Load, Choose Sheet, Find Row, Insert, Copy)
        
        # --- 1. Load JSON & Excel ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)

        # --- 2. Choose sheet ---
        sheet_name = panel_data.get("projectName", "Default")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # --- 3. Find Insertion Point & Define Template ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        TEMPLATE_HEIGHT = TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1
        
        last_schedule_row = find_last_schedule_row(ws, start_row=TEMPLATE_START_ROW, column_to_check=3)
        insertion_row = last_schedule_row + 1

        # --- 4. Insert Blank Rows ---
        ws.insert_rows(insertion_row, amount=TEMPLATE_HEIGHT)

        # --- 5. Copy Template to New Blank Space ---
        for r_offset in range(TEMPLATE_HEIGHT):
            for c in range(1, 13):
                src_cell = ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                dst_cell = ws.cell(row=insertion_row + r_offset, column=c)
                copy_cell_with_style(src_cell, dst_cell)
        
        # --- 6. Write Panel Data into the New Block ---
        row = insertion_row
        
        # Panel metadata
        ws.cell(row=row, column=4).value = panel_data.get("panelName")
        
        # *** THIS IS THE NEW HYPERLINK LOGIC ***
        # Get the URL from the panel data
        source_image_url = panel_data.get("sourceImageUrl")
        # Target the specific cell (A18 in the new block)
        link_cell = ws.cell(row=row + 11, column=1)
        
        # Check if a URL was actually provided
        if source_image_url:
            # Set the visible text of the cell
            link_cell.value = "panel image"
            # Set the hyperlink property of the cell
            link_cell.hyperlink = source_image_url
            # Add blue, underlined font to make it look like a link
            link_cell.font = Font(color="0000FF", underline="single")
        
        ws.cell(row=row + 6, column=7).value = panel_data.get("mountingType", "SURFACE")
        ws.cell(row=row + 7, column=7).value = panel_data.get("ipDegree")
        
        recommendations = panel_data.get("recommendations", [])
        
        # ... (the rest of the breaker logic remains the same)
        main_rec = next((r for r in recommendations if "MCCB" in r.get("breakerSpec", "")), None)
        if main_rec:
            ws.cell(row=row + 11, column=2).value = main_rec.get("breakerSpec")
            ws.cell(row=row + 11, column=9).value = main_rec.get("matchedPart", {}).get("Reference number", "")
            
        branch_recs = [r for r in recommendations if "MCCB" not in r.get("breakerSpec", "")]
        for i, rec in enumerate(branch_recs):
            current_row = row + 13 + i
            ws.cell(row=current_row, column=2).value = rec.get("breakerSpec")
            ws.cell(row=current_row, column=6).value = rec.get("quantity")
            ws.cell(row=current_row, column=9).value = rec.get("matchedPart", {}).get("Reference number", "")

        # --- 7. Save and return the modified file ---
        # ... (this part remains the same)
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
