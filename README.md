# Excel Unprotector

一个用于移除 Excel 文件工作表保护的 Web 应用。支持上传受保护的 `.xlsx` 文件，自动解除所有工作表的保护设置，并返回已解除保护的文件。

## 功能特性

- **简单易用**: 支持拖拽上传和点击选择文件
- **免密解除**: 无需密码即可移除工作表保护
- **批量处理**: 自动处理文件中的所有工作表
- **实时反馈**: 上传进度和处理状态实时显示
- **安全可靠**: 文件在服务器端处理，本地不保存任何数据

## 技术架构

```
┌─────────────────────────────────────────────────────────────┐
│                    Excel Unprotector                        │
├───────────────────────┬─────────────────────────────────────┤
│   Frontend (Client)   │         Backend (Server)            │
├───────────────────────┼─────────────────────────────────────┤
│  React 19 + TypeScript│  Python 3.13 + FastAPI             │
│  Vite 8 + Axios       │  openpyxl + uvicorn                 │
│                       │                                     │
│  http://localhost:5173│  http://localhost:8000              │
└───────────────────────┴─────────────────────────────────────┘
```

## 项目结构

```
excel_unprotector/
├── excel-unprotector-client/   # 前端应用
│   ├── src/
│   │   ├── api/                # API 请求封装
│   │   ├── components/         # React 组件
│   │   ├── hooks/              # 自定义 Hooks
│   │   ├── types/              # TypeScript 类型定义
│   │   └── utils/              # 工具函数
│   ├── package.json
│   ├── vite.config.ts
│   └── README.md
├── excel_unprotector_server/   # 后端服务
│   ├── main.py                 # FastAPI 应用主文件
│   ├── sheet_unprotect.py      # 工作表保护移除逻辑
│   ├── pyproject.toml          # 项目配置与依赖声明
│   └── README.md
└── README.md                   # 项目根文档（本文件）
```

## 快速开始

### 1. 启动后端服务

进入后端目录并启动服务：

```bash
cd excel_unprotector_server

# 安装依赖（需要先安装 uv）
uv sync

# 启动服务
uv run python main.py
```

服务将在 `http://localhost:8000` 运行。

### 2. 启动前端应用

打开新终端，进入前端目录并启动：

```bash
cd excel-unprotector-client

# 安装依赖
pnpm install

# 配置环境变量
cp .env.example .env

# 启动开发服务器
pnpm dev
```

应用将在 `http://localhost:5173` 运行，并自动代理 API 请求到后端服务。

### 3. 使用应用

1. 打开浏览器访问 `http://localhost:5173`
2. 拖拽或点击选择受保护的 `.xlsx` 文件
3. 等待上传和处理完成
4. 点击下载按钮获取已解除保护的文件

## API 接口

| 接口 | 方法 | 描述 |
|------|------|------|
| `/unprotect` | POST | 上传 Excel 文件并解除工作表保护 |
| `/health` | GET | 健康检查 |

### POST /unprotect

**请求参数:**

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| file | UploadFile | 是 | 要上传的 `.xlsx` Excel 文件 |

**示例请求:**

```bash
curl -X POST "http://localhost:8000/unprotect" -F "file=@protected_file.xlsx" -o unprotected_file.xlsx
```

**响应:**

- 200: 返回已解除保护的 Excel 文件
- 400: 请求参数错误（文件名缺失或格式不支持）
- 500: 服务器内部错误

### GET /health

**响应:**

```json
{"status": "ok", "service": "Excel Unprotector API"}
```

## 支持的文件格式

- `.xlsx` 格式（Excel 2007 及以上版本）
- 文件大小限制: 50MB

## 环境要求

### 后端

- Python >= 3.13
- uv（推荐的 Python 包管理器）

### 前端

- Node.js >= 18
- pnpm

## 开发

### 后端开发

```bash
cd excel_unprotector_server
uv run uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

访问 `http://localhost:8000/docs` 查看 API 文档。

### 前端开发

```bash
cd excel-unprotector-client
pnpm dev
```

### 构建生产版本

```bash
# 构建前端
cd excel-unprotector-client
pnpm build
```

## 注意事项

- 该工具仅移除工作表的保护设置，不修改文件中的其他内容
- 建议单次上传文件不超过 50MB
- 文件在服务器端处理完成后立即返回，不会在服务器上保存