import io
import logging
import os
from os import path
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response
from sheet_unprotect import remove_sheet_protection_stream
import uvicorn

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Unprotector API",
    version="0.1.0",
    description="用于移除 Excel 文件工作表保护的 API 服务。支持上传受保护的 .xlsx 文件，并返回已解除保护的文件。",
    openapi_tags=[
        {
            "name": "Excel Protection",
            "description": "Excel 文件工作表保护相关操作"
        },
        {
            "name": "Health Check",
            "description": "服务健康检查"
        }
    ],
    contact={
        "name": "Excel Unprotector Support",
        "email": "support@example.com"
    }
)

ALLOWED_EXTENSIONS = {".xlsx"}


def is_valid_excel_file(filename: str) -> bool:
    ext = path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXTENSIONS


@app.post(
    "/unprotect",
    tags=["Excel Protection"],
    summary="解除 Excel 文件工作表保护",
    description="上传受保护的 .xlsx Excel 文件，服务端将移除所有工作表的保护设置，并返回已解除保护的文件。",
    responses={
        200: {
            "description": "成功解除保护的 Excel 文件",
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    "schema": {
                        "type": "string",
                        "format": "binary"
                    },
                    "example": "unprotected_example.xlsx"
                }
            }
        },
        400: {
            "description": "请求参数错误",
            "content": {
                "application/json": {
                    "schema": {
                        "type": "object",
                        "properties": {
                            "detail": {"type": "string"}
                        }
                    },
                    "examples": {
                        "no_filename": {
                            "value": {"detail": "上传的文件没有文件名"}
                        },
                        "invalid_extension": {
                            "value": {"detail": "只支持 .xlsx 格式的 Excel 文件"}
                        }
                    }
                }
            }
        },
        500: {
            "description": "服务器内部错误",
            "content": {
                "application/json": {
                    "schema": {
                        "type": "object",
                        "properties": {
                            "detail": {"type": "string"}
                        }
                    },
                    "example": {"detail": "文件处理失败: 内部错误"}
                }
            }
        }
    }
)
async def unprotect_excel(file: UploadFile = File(description="要上传的 .xlsx Excel 文件")):
    """
    解除 Excel 文件工作表保护

    - **file**: 上传的 .xlsx 格式 Excel 文件
    - **返回**: 已解除保护的 Excel 文件（.xlsx 格式）

    支持的文件格式：
    - .xlsx (Excel 2007及以上版本)

    注意事项：
    - 该接口仅移除工作表的保护设置，不修改文件中的其他内容
    - 上传文件大小不受限制，但建议单次上传不超过 50MB
    """
    logger.info(f"Received file upload request: {file.filename}, size: {file.size} bytes")

    if not file.filename:
        logger.error("Uploaded file has no filename")
        raise HTTPException(status_code=400, detail="上传的文件没有文件名")

    if not is_valid_excel_file(file.filename):
        logger.error(f"Invalid file extension: {file.filename}")
        raise HTTPException(status_code=400, detail="只支持 .xlsx 格式的 Excel 文件")

    try:
        output_filename = f"unprotected_{file.filename}"
        content = await file.read()

        logger.info(f"Starting sheet protection removal for: {file.filename}")
        input_stream = io.BytesIO(content)
        output_stream = remove_sheet_protection_stream(input_stream)
        output_content = output_stream.read()
        logger.info(f"Sheet protection removal completed: {output_filename}")

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


@app.get(
    "/health",
    tags=["Health Check"],
    summary="健康检查",
    description="检查 API 服务是否正常运行。",
    responses={
        200: {
            "description": "服务正常",
            "content": {
                "application/json": {
                    "schema": {
                        "type": "object",
                        "properties": {
                            "status": {"type": "string"},
                            "service": {"type": "string"}
                        }
                    },
                    "example": {"status": "ok", "service": "Excel Unprotector API"}
                }
            }
        }
    }
)
async def health_check():
    """
    健康检查

    返回服务状态信息，用于监控和负载均衡探测。
    """
    return {"status": "ok", "service": "Excel Unprotector API"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)