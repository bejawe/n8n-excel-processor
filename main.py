import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy
from openpyxl.styles import Font
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

app = FastAPI()

# ==============================================================================
# HELPER FUNCTIONS (WITH ONE KEY CHANGE)
# ==============================================================================
def copy_cell_with_formula_translation(src_cell, dst_cell):
    # (This function is unchanged and correct)
    if src_cell.hyperlink:
        dst_cell.hyperlink = copy(src_cell.hyperlink)
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)

    if src_cell.value and isinstance(src_cell.value, str) and src_cell.value.startswith('='):
        try:
            translator = Translator(formula=src_cell.value, origin=src_cell.coordinate)
            dst_cell.value = translator.translate_formula(dst_cell.coordinate)
        except Exception:
            dst_cell.value = src_cell.value
    else:
        dst_cell.value = src_cell.value

def find_correct_insertion_row(worksheet):
    """
    Finds the correct row to insert a new panel by finding the start of the footer.
    If no footer is found, it assumes the last schedule ends at row 30.
    """
    # Search for a unique value in the footer to find its starting position.
    # We'll use "Eng'r Motaz Abu Jubara" which is in row 25 of the template.
    for row in range(worksheet.max_row, 1, -1):
        # We check column B (column index 2) for the unique text.
        cell_value = worksheet.cell(row=row, column=2).value
        if cell_value and "Eng'r Motaz Abu Jubara" in str(cell_value):
            # The footer block starts at row 24 of the template, so the real start is row - 1.
            return row - 1
            
    # If the footer is not found, fall back to the last "TOTAL" of the final schedule.
    for row_num in range(worksheet.max_row, 1, -1):
        cell_value = worksheet.cell(row=row_num, column=3).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num + 1
            
    return 31 # Absolute fallback if the sheet is clean except for the template.

def is_template_empty(worksheet, check_row, check_col):
    # (This function is unchanged and correct)
    cell_value = worksheet.cell(row=check_row, column=check_col).value
    return cell_value is None or "N/A" in str(cell_value)

# ==============================================================================
# MAIN API ENDPOINT
# ==============================================================================
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- This section uses your stable, working code ---
        panel_data = json.loads(panel_data_json)
        contents = await file.read()
        wb_to_modify = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)
        ws_to_modify = wb_to_modify.active
        pristine_template_wb = openpyxl.load_workbook('template.xlsm', keep_vba=True)
        pristine_ws = pristine_template_wb.active

        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW = 30
        TEMPLATE_HEIGHT = TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1
        TEMPLATE_COLUMN_COUNT = 44

        if is_template_empty(ws_to_modify, check_row=20, check_col=9):
            write_row = TEMPLATE_START_ROW
        else:
            # --- THE KEY CHANGE: Use the new, smarter function ---
            insertion_row = find_correct_insertion_row(ws_to_modify)
            
            ws_to_modify.insert_rows(insertion_row, amount=TEMPLATE_HEIGHT)

            # Copy cell contents and styles
            for r_offset in range(TEMPLATE_HEIGHT):
                for c in range(1, TEMPLATE_COLUMN_COUNT + 1):
                    src_cell = pristine_ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                    dst_cell = ws_to_modify.cell(row=insertion_row + r_offset, column=c)
                    copy_cell_with_formula_translation(src_cell, dst_cell)
            
            # Copy row and column dimensions for formatting
            for r_offset in range(TEMPLATE_HEIGHT):
                source_row_index = TEMPLATE_START_ROW + r_offset
                destination_row_index = insertion_row + r_offset
                if source_row_index in pristine_ws.row_dimensions:
                    ws_to_modify.row_dimensions[destination_row_index].height = pristine_ws.row_dimensions[source_row_index].height
            write_row = insertion_row

        for i in range(1, TEMPLATE_COLUMN_COUNT + 1):
            column_letter = get_column_letter(i)
            if column_letter in pristine_ws.column_dimensions:
                ws_to_modify.column_dimensions[column_letter].width = pristine_ws.column_dimensions[column_letter].width
        
        # --- Write Panel-Specific Data (Unchanged) ---
        row = write_row
        ws_to_modify.cell(row=row, column=4).value = panel_data.get("panelName")
        # (...all other data writing logic is the same...)
        ws_to_modify.cell(row=row + 23, column=3).value = "TOTAL"
        
        # --- Clean Up Unused "OUTGOINGS" Rows (Unchanged) ---
        outgoings_start_row = row + 13
        outgoings_end_row = row + 22
        for row_to_check in range(outgoings_end_row, outgoings_start_row - 1, -1):
            qty_cell = ws_to_modify.cell(row=row_to_check, column=6)
            part_num_cell = ws_to_modify.cell(row=row_to_check, column=9)
            part_num_is_default = (part_num_cell.value is None or "N/A" in str(part_num_cell.value))
            qty_is_default = (qty_cell.value == 1)
            if part_num_is_default and qty_is_default:
                ws_to_modify.delete_rows(row_to_check, 1)

        # --- Save and Return (Unchanged) ---
        out = io.BytesIO()
        wb_to_modify.save(out)
        out.seek(0)
        
        original_filename = file.filename
        media_type = 'application/vnd.ms-excel.sheet.macroenabled.12' if original_filename.lower().endswith('.xlsm') else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        
        return StreamingResponse(
            out,
            media_type=media_type,
            headers={"Content-Disposition": f"attachment; filename={original_filename}"}
        )

    except FileNotFoundError:
        raise HTTPException(status_code=500, detail="Server configuration error: The master template file is missing.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An error occurred during Excel processing: {str(e)}")
