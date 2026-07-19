# Excel Unprotector Server

用于移除 Excel 文件工作表保护的 API 服务。

## 功能特性

- 上传受保护的 `.xlsx` Excel 文件
- 自动移除所有工作表的保护设置
- 返回已解除保护的 Excel 文件
- 提供健康检查接口

## 技术栈

- Python 3.13
- FastAPI
- openpyxl
- uvicorn
- **uv** (推荐的 Python 包管理器)

## 快速开始

### 安装 uv

确保已安装 uv（Python 包管理器）。若未安装，请参考以下方式：

**Windows (PowerShell):**
```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

**其他平台:**
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 安装依赖

推荐使用 uv 安装项目依赖：

```bash
uv sync
```

此命令会自动：
1. 创建虚拟环境（若不存在）
2. 安装所有依赖（基于 `pyproject.toml` 和 `uv.lock`）
3. 确保依赖版本一致

### 启动服务

```bash
uv run python main.py
```

或使用 uvicorn 直接启动：

```bash
uv run uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

服务启动后访问：
- API 文档：http://localhost:8000/docs
- 健康检查：http://localhost:8000/health

## API 接口

### POST /unprotect

**解除 Excel 文件工作表保护**

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| file | UploadFile | 是 | 要上传的 `.xlsx` Excel 文件 |

**示例请求:**
```bash
curl -X POST "http://localhost:8000/unprotect" -F "file=@protected_file.xlsx" -o unprotected_file.xlsx
```

**响应:**
- 200: 返回已解除保护的 Excel 文件（`.xlsx` 格式）
- 400: 请求参数错误（文件名缺失或格式不支持）
- 500: 服务器内部错误

### GET /health

**健康检查**

**示例请求:**
```bash
curl http://localhost:8000/health
```

**响应:**
```json
{"status": "ok", "service": "Excel Unprotector API"}
```

## 项目结构

```
excel_unprotector_server/
├── main.py          # FastAPI 应用主文件
├── sheet_unprotect.py  # 工作表保护移除逻辑
├── pyproject.toml   # 项目配置与依赖声明
├── uv.lock          # uv 依赖锁文件
├── .python-version  # Python 版本指定
├── data/            # 示例数据目录
│   └── protected_file.xlsx
└── test.http        # HTTP 测试文件
```

## 环境变量

| 变量名 | 默认值 | 描述 |
|--------|--------|------|
| PORT | 8000 | 服务监听端口 |

## 注意事项

- 仅支持 `.xlsx` 格式（Excel 2007 及以上版本）
- 该接口仅移除工作表的保护设置，不修改文件中的其他内容
- 建议单次上传文件不超过 50MB
