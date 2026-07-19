import logging
import tempfile
from os import path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response
from sheet_unprotect import remove_sheet_protection
import uvicorn

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

app = FastAPI(title="Excel Unprotector API", version="0.1.0")

ALLOWED_EXTENSIONS = {".xlsx"}


def is_valid_excel_file(filename: str) -> bool:
    ext = path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS


@app.post("/unprotect")
async def unprotect_excel(file: UploadFile = File(...)):
    logger.info(f"Received file upload request: {file.filename}, size: {file.size} bytes")

    if not file.filename:
        logger.error("Uploaded file has no filename")
        raise HTTPException(status_code=400, detail="上传的文件没有文件名")

    if not is_valid_excel_file(file.filename):
        logger.error(f"Invalid file extension: {file.filename}")
        raise HTTPException(status_code=400, detail="只支持 .xlsx 格式的 Excel 文件")

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = path.join(temp_dir, file.filename)
            output_filename = f"unprotected_{file.filename}"
            output_path = path.join(temp_dir, output_filename)

            logger.info(f"Saving uploaded file to: {input_path}")
            with open(input_path, "wb") as f:
                content = await file.read()
                f.write(content)

            logger.info(f"Starting sheet protection removal for: {file.filename}")
            remove_sheet_protection(input_path, output_path)
            logger.info(f"Sheet protection removal completed: {output_filename}")

            with open(output_path, "rb") as f:
                output_content = f.read()

        return Response(
            content=output_content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={output_filename}"
            }
        )

    except Exception as e:
        logger.error(f"Error processing file {file.filename}: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"文件处理失败: {str(e)}")


@app.get("/health")
async def health_check():
    return {"status": "ok", "service": "Excel Unprotector API"}


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)