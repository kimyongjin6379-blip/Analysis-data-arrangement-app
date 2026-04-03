"""
펩리치 바이오펩톤 성분 분석 자동화 — FastAPI 백엔드
"""

from __future__ import annotations

import json
import os
import uuid
import tempfile
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from processor import classify_file, prescan_files, process_all

app = FastAPI(title="Peptone Analysis Data Arrangement Tool", version="1.0.0")

TEMP_DIR = Path(tempfile.gettempdir()) / "peptone_analysis_outputs"
TEMP_DIR.mkdir(exist_ok=True)


@app.post("/api/prescan")
async def prescan_upload(files: List[UploadFile] = File(...)):
    """업로드 파일에서 시료명 자동 추출."""
    file_bytes_list = []
    file_info = []
    for f in files:
        fbytes = await f.read()
        fname = f.filename or "unknown.xlsx"
        ftype = classify_file(fname)
        file_bytes_list.append((fname, fbytes))
        file_info.append({"filename": fname, "type": ftype, "size": len(fbytes)})

    result = prescan_files(file_bytes_list)
    return JSONResponse(content={
        "success": True,
        "files": file_info,
        "lab_samples": result["lab_samples"],
        "summary_samples": result["summary_samples"],
    })


@app.post("/api/process")
async def process_upload(
    files: List[UploadFile] = File(...),
    sample_config_json: str = Form("[]"),
    sensang_data_json: str = Form("{}"),
    batch_date: str = Form(""),
):
    """파일 + 설정 → 가공된 엑셀 생성."""
    if not files:
        raise HTTPException(status_code=400, detail="파일을 업로드해 주세요.")

    file_bytes_list = []
    for f in files:
        fbytes = await f.read()
        if not fbytes:
            continue
        fname = f.filename or "unknown.xlsx"
        ext = Path(fname).suffix.lower()
        if ext not in (".xlsx", ".xls"):
            raise HTTPException(status_code=400, detail=f"지원되지 않는 파일: {fname}")
        file_bytes_list.append((fname, fbytes))

    if not file_bytes_list:
        raise HTTPException(status_code=400, detail="유효한 엑셀 파일이 없습니다.")

    try:
        sample_config = json.loads(sample_config_json) if sample_config_json else []
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="시료 설정 JSON이 올바르지 않습니다.")

    try:
        sensang_data = json.loads(sensang_data_json) if sensang_data_json else {}
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="성상 데이터 JSON이 올바르지 않습니다.")

    if not sample_config:
        raise HTTPException(status_code=400, detail="시료 설��이 필요합니다.")

    try:
        excel_bytes, summary_info = process_all(file_bytes_list, sample_config, sensang_data, batch_date)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"데이터 처리 오류: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"서버 처리 오류: {str(e)}")

    file_id = str(uuid.uuid4())
    # 파일명: "샘플명 분석.xlsx" 형태
    sample_names = [sc["display_name"] for sc in sample_config]
    if len(sample_names) <= 3:
        name_part = ", ".join(sample_names)
    else:
        name_part = f"{sample_names[0]} 외 {len(sample_names)-1}종"
    output_filename = f"{name_part} 분석.xlsx"
    output_path = TEMP_DIR / f"{file_id}.xlsx"
    output_path.write_bytes(excel_bytes)

    return JSONResponse(content={
        "success": True,
        "file_id": file_id,
        "filename": output_filename,
        "summary": summary_info,
        "message": "데이터 가공이 완료되었습니다.",
    })


@app.get("/api/download/{file_id}")
async def download_file(file_id: str, filename: str = "output.xlsx"):
    """가공된 엑셀 파일 다운로드."""
    output_path = TEMP_DIR / f"{file_id}.xlsx"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다. 다시 가공해 주세요.")

    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )


static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
