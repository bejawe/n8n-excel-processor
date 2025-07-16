import io
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl
from copy import copy

app = FastAPI()

# ---------- helper ----------
def copy_cell_with_style(src, dst):
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
    if not file.filename.lower().endswith((".xlsm", ".xlsx")):
        raise HTTPException(status_code=400, detail="Invalid file format.")

    try:
        # --- load JSON & Excel ---
        panel_data = json.loads(panel_data_json)
        contents  = await file.read()
        wb        = openpyxl.load_workbook(io.BytesIO(contents), keep_vba=True)

        # --- choose sheet ---
        sheet_name = panel_data.get("projectName", "")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        print(f"DEBUG: sheet = {ws.title}")

        # --- template bounds ---
        TEMPLATE_START_ROW = 7
        TEMPLATE_END_ROW   = 30
        next_row           = ws.max_row + 2
        print(f"DEBUG: next_row = {next_row}")

        # --- copy template block (24 rows × 12 cols) ---
        for r_offset in range(TEMPLATE_END_ROW - TEMPLATE_START_ROW + 1):
            for c in range(1, 13):
                src = ws.cell(row=TEMPLATE_START_ROW + r_offset, column=c)
                dst = ws.cell(row=next_row + r_offset, column=c)
                copy_cell_with_style(src, dst)

        # --- write panel data into the new block ---
        # === WRITE PANEL DATA ===
# Row offsets are relative to the *start* of the copied block (next_row)

row = next_row  # first row of the new block

# Panel name
ws.cell(row=row, column=4).value = panel_data.get("panelName")

# Source image URL
ws.cell(row=row + 11, column=1).value = panel_data.get("sourceImageUrl")

# Mounting type & IP code
ws.cell(row=row + 6, column=7).value = panel_data.get("mountingType", "SURFACE")
ws.cell(row=row + 7, column=7).value = panel_data.get("ipDegree")

# Main breaker
main_rec = next((r for r in panel_data.get("recommendations", []) if "MCCB" in r["breakerSpec"]), None)
if main_rec:
    ws.cell(row=row + 11, column=2).value = main_rec["breakerSpec"]
    ws.cell(row=row + 11, column=9).value = main_rec["matchedPart"].get("Reference number", "")

# Branch breakers
for i, rec in enumerate(
    [r for r in panel_data.get("recommendations", []) if "MCCB" not in r["breakerSpec"]]
):
    r = row + 13 + i
    ws.cell(row=r, column=2).value = rec["breakerSpec"]
    ws.cell(row=r, column=6).value = rec["quantity"]
    ws.cell(row=r, column=9).value = rec["matchedPart"].get("Reference number", "")

        # main breaker
        main_rec = next((r for r in panel_data.get("recommendations", []) if "MCCB" in r["breakerSpec"]), None)
        if main_rec:
            ws.cell(row=next_row + 11, column=2).value = main_rec["breakerSpec"]
            ws.cell(row=next_row + 11, column=9).value = main_rec["matchedPart"].get("Reference number", "")

        # branch breakers
        branch_recs = [r for r in panel_data.get("recommendations", []) if "MCCB" not in r["breakerSpec"]]
        for idx, rec in enumerate(branch_recs):
            row = next_row + 13 + idx
            ws.cell(row=row, column=2).value = rec["breakerSpec"]
            ws.cell(row=row, column=6).value = rec["quantity"]
            ws.cell(row=row, column=9).value = rec["matchedPart"].get("Reference number", "")

        # --- return new file ---
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return StreamingResponse(
            out,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=modified_{file.filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
