import io
import json
from fastapi import FastAPI, (columns A-AR, which is 44 columns).
3.  **The "Insert and Copy" Fix UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy:** When a new panel needs to be added (i.e., the sheet is not empty), we will:
    *    import copy
from openpyxl.styles import Font
from openpyxl.formula.translate import Translator

app = FastAPI()

# ==============================================================================
# FINAL, VALIDATED HELPER FUNCTIONS
# =================Find the last row of the last *panel block*.
    *   Insert enough blank rows for a new panel *at that location*. This will automatically push the footer down.
    *   Copy the pristine *panel block* (rows=============================================================
def copy_cell_with_formula_translation(src_cell, dst_cell): 7-23, columns A-AR) from the clean `template.xlsm` into this new space.

    if src_cell.hyperlink:
        dst_cell.hyperlink = copy(src_cell.hyperlink)
    if src_cell.has_style:
        dst_cell.font = copy(srcYou are absolutely right. My sincerest apologies. You have correctly identified that the last code    *   Write the new panel's data.

This is the final piece of the puzzle. It respects your working I provided does not handle the "footer" block correctly. It copies the entire template (rows 7-30), which incorrectly duplicates the footer inside every new panel.

Let's go back to the logic that correctly handled the footer and combine it with the wider column copy.

### The Correct Logic (Combining Our Successes)

1.  **Define Two Blocks:** We will be precise. The "Panel" is rows 7-23. The "Footer" is rows 24-29.
2.  **Find Insertion Point:** We will find the last "TOTAL" row of the last panel. The new panel will be inserted right after this.
3.  **Insert Rows:** We_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy code and adds the one targeted change needed to handle the footer correctly.

---

### The Final `main.py`

 will insert enough blank rows for a *Panel only*. This will correctly push the existing footer down.
4.  **Copy Wide:** We will copy the *Panel Block only* (rows 7-23) from the pristine template, but we will copy all columns from A to AR.
5.  **Write Data:** We will write the new data into this(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protectionThis code integrates the precise footer-pushing logic into your stable, working code.

```python
import io
import json newly created, wide panel block.

This combines all the successful pieces of logic we have developed.

---

###)
        dst_cell.alignment = copy(src_cell.alignment)

    if src_cell.value and isinstance(src_cell.value, str) and src_cell.value.startswith('='):
        try Final Production Code (with Footer and Wide Column Fixes)

This is the definitive code. It correctly handles the separate
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import:
            translator = Translator(formula=src_cell.value, origin=src_cell.coordinate)
            dst openpyxl
from copy import copy
from openpyxl.styles import Font
from openpyxl.formula_cell.value = translator.translate_formula(dst_cell.coordinate)
        except Exception:
            dst panel/footer blocks and copies the entire width of the panel.

```python
import io
import json
from.translate import Translator

app = FastAPI()

# ==============================================================================
# HELPER FUNCTIONS (These_cell.value = src_cell.value
    else:
        dst_cell.value = src_cell.value fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from are validated and correct)
# ==============================================================================
def copy_cell_with_formula_translation

def find_last_schedule_row(worksheet, start_row, column_to_check):
    for copy import copy
from openpyxl.styles import Font
from openpyxl.formula.translate import Translator

app(src_cell, dst_cell):
    if src_cell.hyperlink:
        dst_cell.hyper row_num in range(worksheet.max_row, start_row, -1):
        cell_value = = FastAPI()

# ==============================================================================
# FINAL, VALIDATED HELPER FUNCTIONS
# ================= worksheet.cell(row=row_num, column=column_to_check).value
        if isinstance(cell_valuelink = copy(src_cell.hyperlink)
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border=============================================================
def copy_cell_with_formula_translation(src_cell, dst_cell):, str) and "TOTAL" in cell_value.upper():
            return row_num
    return )
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number23

def is_template_empty(worksheet, check_row, check_col):
    cell
    if src_cell.hyperlink:
        dst_cell.hyperlink = copy(src_cell._format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_value = worksheet.cell(row=check_row, column=check_col).value
    return cellhyperlink)
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)

    if src__cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_value is None or "N/A" in str(cell_value)

# ==============================================================================cell.value and isinstance(src_cell.value, str) and src_cell.value.startswith('='_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy):
        try:
            translator = Translator(formula=src_cell.value, origin=src_cell
# FINAL, VALIDATED MAIN API ENDPOINT
# ==============================================================================
@app.post("/(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection.coordinate)
            dst_cell.value = translator.translate_formula(dst_cell.coordinate)
        process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form)
        dst_cell.alignment = copy(src_cell.alignment)

    if src_cell.valueexcept Exception:
            dst_cell.value = src_cell.value
    else:
        dst_cell(...),
    file: UploadFile = File(...)
):
    if not file.filename.lower().endswith and isinstance(src_cell.value, str) and src_cell.value.startswith('='):
        try.value = src_cell.value

def find_last_schedule_row(worksheet, start_row:
            translator = Translator(formula=src_cell.value, origin=src_cell.coordinate)
            dst((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file_cell.value = translator.translate_formula(dst_cell.coordinate)
        except Exception:
            dst format.")

    try:
        panel_data = json.loads(panel_data_json)
        , column_to_check):
    for row_num in range(worksheet.max_row, start_row,_cell.value = src_cell.value
    else:
        dst_cell.value = src_ -1):
        cell_value = worksheet.cell(row=row_num, column=column_tocontents = await file.read()
        
        wb_to_modify = openpyxl.load_workbook(cell.value

def find_last_schedule_row(worksheet, start_row, column_to_check):
io.BytesIO(contents), keep_vba=True)
        ws_to_modify = wb_to_modify    for row_num in range(worksheet.max_row, start_row, -1):
        .active
        
        pristine_template_wb = openpyxl.load_workbook('template.xlcell_value = worksheet.cell(row=row_num, column=column_to_check).value
sm', keep_vba=True)
        pristine_ws = pristine_template_wb.active        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row

        # --- Redefined Constants for the PANEL block and its full width ---
        PANEL_TEMPLATE_START_num
    return 23 # End of the first panel template

def is_template_empty(worksheet,_ROW = 7
        PANEL_TEMPLATE_END_ROW = 23 # The Panel block ends at the check_row, check_col):
    cell_value = worksheet.cell(row=check_row, column=check TOTAL row
        PANEL_TEMPLATE_HEIGHT = PANEL_TEMPLATE_END_ROW - PANEL_TEMPLATE_START_col).value
    return cell_value is None or "N/A" in str(cell_value)_check).value
        if isinstance(cell_value, str) and "TOTAL" in cell_value.upper():
            return row_num
    return 23 # The end row of the first panel block

def is_template_empty(worksheet, check_row, check_col):
    cell_value = worksheet.cell(row=check_row, column=check_col).value
    return cell_value is None or "N/A" in str(cell_value)

# ==============================================================================
# MAIN API ENDPOINT
# ==============================================================================
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data_json: str = Form(...),
    file: UploadFile =_ROW + 1
        TEMPLATE_COLUMN_WIDTH = 44 # Column AR

        if is_template_empty(ws_to_modify, check_row=18, check_col=9):
            write

# ==============================================================================
# FINAL, VALIDATED MAIN API ENDPOINT
# ============================================================================== File(...)
):
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise_row = PANEL_TEMPLATE_START_ROW
        else:
            last_schedule_row = find_
@app.post("/process-excel-panel/")
async def process_panel(
    panel_data HTTPException(status_code=400, detail="Invalid file format.")

    try:
        panel_datalast_schedule_row(ws_to_modify, start_row=PANEL_TEMPLATE_END_ROW_json: str = Form(...),
    file: UploadFile = File(...)
):
    if not file = json.loads(panel_data_json)
        contents = await file.read()
        
        wb_, column_to_check=3)
            insertion_row = last_schedule_row + 1
            
to_modify = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)            ws_to_modify.insert_rows(insertion_row, amount=PANEL_TEMPLATE_HEIGHT).filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400,
        ws_to_modify = wb_to_modify.active
        
        pristine_template_wb

            # --- Copy the WIDER range of the pristine PANEL block ---
            for r_offset in range(PANEL_ detail="Invalid file format.")

    try:
        panel_data = json.loads(panel_data_ = openpyxl.load_workbook('template.xlsm', keep_vba=True)
        pristineTEMPLATE_HEIGHT):
                for c in range(1, TEMPLATE_COLUMN_WIDTH + 1): #json)
        contents = await file.read()
        
        wb_to_modify = openpyxl._ws = pristine_template_wb.active

        # --- PRECISE TEMPLATE DEFINITIONS ---
        PAN Loop from A to AR
                    src_cell = pristine_ws.cell(row=PANEL_TEMPLATE_load_workbook(io.BytesIO(contents), keep_vba=True)
        ws_to_modifySTART_ROW + r_offset, column=c)
                    dst_cell = ws_to_modify.cell(rowEL_START_ROW = 7
        PANEL_END_ROW = 23 # The panel ends = wb_to_modify.active
        
        pristine_template_wb = openpyxl.load_workbook at the 'TOTAL' row
        PANEL_HEIGHT = PANEL_END_ROW - PANEL_START_ROW + 1=insertion_row + r_offset, column=c)
                    copy_cell_with_formula_translation('template.xlsm', keep_vba=True)
        pristine_ws = pristine_template
        TEMPLATE_WIDTH_COLUMNS = 44 # Copy up to column AR

        # Check if the In(src_cell, dst_cell)
            
            write_row = insertion_row

        # --- Write_wb.active

        # --- Redefined Constants for PANEL block and full width ---
        PANEL_TEMPLATE All Panel Data into the Target Block ---
        row = write_row
        
        ws_to_modify.comer Part Number cell (I18) is empty
        if is_template_empty(ws_to_modify,_START_ROW = 7
        PANEL_TEMPLATE_END_ROW = 23
        PANEL_cell(row=row, column=4).value = panel_data.get("panelName")
        ws check_row=18, check_col=9):
            write_row = PANEL_START_ROW
        elseTEMPLATE_HEIGHT = PANEL_TEMPLATE_END_ROW - PANEL_TEMPLATE_START_ROW + 1
        :
            last_schedule_row = find_last_schedule_row(ws_to_modify, start_to_modify.cell(row=row + 6, column=7).value = panel_data.get_row=PANEL_END_ROW, column_to_check=3)
            insertion_row = lastTEMPLATE_COLUMN_WIDTH = 44 # Column AR

        if is_template_empty(ws_to_modify,("mountingType", "SURFACE")
        ws_to_modify.cell(row=row + 7, column check_row=18, check_col=9):
            write_row = PANEL_TEMPLATE_START_schedule_row + 1
            
            # --- THE FINAL FIX: Insert rows to push footer down ---
            =7).value = panel_data.get("ipDegree")

        link_cell = ws_to_modify_ROW
        else:
            last_schedule_row = find_last_schedule_row(ws_ws_to_modify.insert_rows(insertion_row, amount=PANEL_HEIGHT)

            # Copy.cell(row=row + 11, column=1)
        source_image_url = panelto_modify, start_row=PANEL_TEMPLATE_END_ROW, column_to_check=3 ONLY the pristine PANEL block into the new space
            for r_offset in range(PANEL_HEIGHT):
                _data.get("sourceImageUrl")
        if source_image_url:
            link_cell.value = ")
            insertion_row = last_schedule_row + 1
            
            ws_to_modifyfor c in range(1, TEMPLATE_WIDTH_COLUMNS + 1):
                    src_cell = pristinepanel image"
            link_cell.hyperlink = source_image_url
            link_cell.font.insert_rows(insertion_row, amount=PANEL_TEMPLATE_HEIGHT)

            # Copy ONLY the pristine_ws.cell(row=PANEL_START_ROW + r_offset, column=c)
                     = Font(color="0000FF", underline="single")

        recommendations = panel_data.get PANEL block, but for the FULL WIDTH
            for r_offset in range(PANEL_TEMPLATE_HEIGHT):dst_cell = ws_to_modify.cell(row=insertion_row + r_offset, column=
                for c in range(1, TEMPLATE_COLUMN_WIDTH + 1):
                    src_cell("recommendations", [])
        
        main_rec = next((r for r in recommendations if "MCCBc)
                    copy_cell_with_formula_translation(src_cell, dst_cell)
            
            write = pristine_ws.cell(row=PANEL_TEMPLATE_START_ROW + r_offset, column=c)
_row = insertion_row

        # --- Write All Panel Data (Your stable logic) ---
        row = write" in r.get("breakerSpec", "")), None)
        if main_rec:
            ws_to_modify                    dst_cell = ws_to_modify.cell(row=insertion_row + r_offset, column=c_row
        
        ws_to_modify.cell(row=row, column=4).value = panel.cell(row=row + 11, column=9).value = main_rec.get("matched)
                    copy_cell_with_formula_translation(src_cell, dst_cell)
            
            write_data.get("panelName")
        ws_to_modify.cell(row=row + 6, columnPart", {}).get("Reference number", "")
        
        branch_recs = [r for r in_row = insertion_row

        # --- Write All Panel Data into the Target Block ---
        row = write=7).value = panel_data.get("mountingType", "SURFACE")
        ws_to__row
        
        ws_to_modify.cell(row=row, column=4).value = panelmodify.cell(row=row + 7, column=7).value = panel_data.get("ip recommendations if "MCCB" not in r.get("breakerSpec", "")]
        for i, rec in enumerate(Degree")

        link_cell = ws_to_modify.cell(row=row + 11, columnbranch_recs):
            current_row = row + 13 + i
            if current_row <=_data.get("panelName")
        ws_to_modify.cell(row=row + 6, column=1)
        source_image_url = panel_data.get("sourceImageUrl")
        if source_image_ (row + PANEL_TEMPLATE_HEIGHT - 1):
                ws_to_modify.cell(row=current=7).value = panel_data.get("mountingType", "SURFACE")
        ws_to__row, column=6).value = rec.get("quantity")
                ws_to_modify.cell(rowurl:
            link_cell.value = "panel image"
            link_cell.hyperlink = sourcemodify.cell(row=row + 7, column=7).value = panel_data.get("ipDegree")

        link_cell = ws_to_modify.cell(row=row + 11, column_image_url
            link_cell.font = Font(color="0000FF", underline="=current_row, column=9).value = rec.get("matchedPart", {}).get("Reference number",=1)
        source_image_url = panel_data.get("sourceImageUrl")
        if source_image_single")

        recommendations = panel_data.get("recommendations", [])
        
        main_rec "")
        
        # --- Save and Return ---
        out = io.BytesIO()
        wb_url:
            link_cell.value = "panel image"
            link_cell.hyperlink = sourceto_modify.save(out)
        out.seek(0)
        
        original_filename = file.filename = next((r for r in recommendations if "MCCB" in r.get("breakerSpec", "")), None_image_url
            link_cell.font = Font(color="0000FF", underline="
        media_type = 'application/vnd.ms-excel.sheet.macroenabled.12' if)
        if main_rec:
            ws_to_modify.cell(row=row + 11, columnsingle")

        recommendations = panel_data.get("recommendations", [])
        
        main_rec original_filename.lower().endswith('.xlsm') else 'application/vnd.openxmlformats-officedocument.=9).value = main_rec.get("matchedPart", {}).get("Reference number", "")
        
        branch = next((r for r in recommendations if "MCCB" in r.get("breakerSpec", "")), Nonespreadsheetml.sheet'
        
        return StreamingResponse(
            out,
            media_type=media_type,_recs = [r for r in recommendations if "MCCB" not in r.get("breakerSpec",)
        if main_rec:
            ws_to_modify.cell(row=row + 11
            headers={"Content-Disposition": f"attachment; filename={original_filename}"}
        )

    except "")]
        for i, rec in enumerate(branch_recs):
            current_row = row + 13, column=9).value = main_rec.get("matchedPart", {}).get("Reference number", "")
 FileNotFoundError:
        raise HTTPException(status_code=500, detail="Server configuration error: The master template + i
            if current_row <= (row + PANEL_HEIGHT - 1):
                ws_to_        
        branch_recs = [r for r in recommendations if "MCCB" not in r.get file is missing.")
    except Exception as e:
        raise HTTPException(status_code=500, detailmodify.cell(row=current_row, column=6).value = rec.get("quantity")
                ws("breakerSpec", "")]
        for i, rec in enumerate(branch_recs):
            current_row = row=f"An error occurred during Excel processing: {str(e)}")
```_to_modify.cell(row=current_row, column=9).value = rec.get("matchedPart
