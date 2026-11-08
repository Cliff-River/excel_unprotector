# Excel Unprotector

一个用于移除 Excel 文件工作表保护的工具，提供两种形态：

- **Web 版**：上传受保护的 `.xlsx` 文件，由 FastAPI 服务端解除所有工作表的保护后返回。
- **桌面版**：基于 Tauri v2 的完全离线客户端，复用同一套 Next.js 页面，解保护由 Rust 在本地完成，结果通过原生“另存为”对话框保存。

## 功能特性

- **简单易用**：支持拖拽上传和点击选择文件
- **免密解除**：无需密码即可移除工作表保护
- **批量处理**：自动处理文件中的所有工作表
- **实时反馈**：Web 版显示上传进度与处理状态，桌面版显示处理中状态
- **双构建目标**：同一套前端源码可构建 Web（standalone）与桌面（静态导出）两种产物
- **完全离线**：桌面版不依赖后端服务与网络，文件不离开本机

## 技术架构

```
┌──────────────────────── Web 版 ────────────────────────┐
│ Next.js 16 (App Router, 客户端渲染)                     │
│   axios → /api/unprotect → FastAPI + openpyxl          │
└────────────────────────────────────────────────────────┘

┌─────────────────────── 桌面版 ─────────────────────────┐
│ Tauri v2 WebView ← Next.js 静态导出物（out/ → dist/）   │
│   invoke("unprotect_excel") → Rust：zip + quick-xml    │
│   invoke("save_file")       → 原生另存为 + 写盘         │
└────────────────────────────────────────────────────────┘
```

## 项目结构

```
excel_unprotector/
├── excel-unprotector-nextjs/          # Next.js 前端（Web + 桌面共用源码）
│   ├── app/                           # App Router 页面
│   ├── components/                    # UI 组件
│   ├── lib/
│   │   ├── api/                       # 平台适配层（web / tauri 双实现）
│   │   ├── hooks/                     # useUpload 等
│   │   ├── types/
│   │   └── utils/
│   ├── scripts/export-desktop.mjs     # 桌面静态导出 + 同步到 tauri/dist
│   ├── i18n/、messages/、tests/、public/
│   ├── Dockerfile
│   └── next.config.ts                 # 按 BUILD_TARGET 切换构建目标
├── excel-unprotector-server/          # FastAPI 后端（Web 版使用，桌面版不依赖）
│   ├── src/excel_unprotector_server/
│   │   ├── main.py                    # API 入口
│   │   └── sheet_unprotect.py         # openpyxl 解保护逻辑
│   ├── data/protected_file.xlsx       # 示例文件 / Rust 单测夹具
│   └── Dockerfile
├── excel-unprotector-tauri/           # Tauri v2 桌面壳
│   ├── dist/                          # 构建生成：Next.js 静态导出物
│   ├── package.json                   # 仅 @tauri-apps/cli
│   ├── src-tauri/
│   │   ├── src/lib.rs                 # IPC 命令注册
│   │   ├── src/excel.rs               # Rust 解保护算法与单测
│   │   └── tauri.conf.json
│   └── README.md                      # 桌面端开发与构建说明
├── requirements.md                    # 桌面端开发需求
├── docker-compose.yml
├── DEPLOYMENT.md                      # Web 版部署文档
├── LICENSE
└── README.md
```

## 快速开始

### 方式一：Docker Compose（Web 版）

```bash
docker compose up -d --build
docker compose down
```

- 前端页面：http://localhost:3001
- 后端 API：http://localhost:8001
- API 文档：http://localhost:8001/docs

### 方式二：本地开发（Web 版）

#### 1. 启动后端服务

```bash
cd excel-unprotector-server
uv sync
uv run python -m excel_unprotector_server.main
```

服务运行在 `http://localhost:8000`。

#### 2. 启动前端

```bash
cd excel-unprotector-nextjs
pnpm install
pnpm dev
```

访问 `http://localhost:3000`，`/api` 请求会代理到本地后端。

### 方式三：桌面版

```bash
cd excel-unprotector-tauri
pnpm install
pnpm tauri dev
```

`tauri dev` 会自动在 `1420` 端口拉起 Next.js 开发服务器并打开桌面窗口。生产包构建：

```bash
pnpm tauri build
```

产物位于 `src-tauri/target/release/bundle/`（deb / AppImage）。环境依赖、架构与测试说明见 [excel-unprotector-tauri/README.md](excel-unprotector-tauri/README.md)。

## API 接口

| 接口 | 方法 | 描述 |
|------|------|------|
| `/unprotect` | POST | 上传 Excel 文件并解除工作表保护 |
| `/health` | GET | 健康检查 |

### POST /unprotect

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| file | UploadFile | 是 | 要上传的 `.xlsx` Excel 文件 |

```bash
curl -X POST "http://localhost:8000/unprotect" -F "file=@protected_file.xlsx" -o unprotected_file.xlsx
```

响应：

- 200：返回已解除保护的 Excel 文件
- 400：请求参数错误（文件名缺失或格式不支持）
- 500：服务器内部错误

### GET /health

```json
{"status": "ok", "service": "Excel Unprotector API"}
```

## 支持的文件格式

- `.xlsx`（Excel 2007 及以上版本）
- 文件大小限制：50MB

## 环境要求

### 后端

- Python >= 3.13
- uv

### 前端 / 桌面

- Node.js >= 22.0.0
- pnpm >= 9.0.0
- 桌面版额外需要 Rust stable 与 Tauri v2 系统依赖（Linux 见 `excel-unprotector-tauri/README.md`）

## 开发与构建

| 场景 | 命令（目录） |
|------|--------------|
| 后端开发 | `uv run uvicorn excel_unprotector_server.main:app --reload`（`excel-unprotector-server/`） |
| Web 开发 | `pnpm dev`（`excel-unprotector-nextjs/`） |
| Web 生产构建 | `pnpm build`（`excel-unprotector-nextjs/`） |
| 桌面开发 | `pnpm tauri dev`（`excel-unprotector-tauri/`） |
| 桌面生产构建 | `pnpm tauri build`（`excel-unprotector-tauri/`） |
| 仅导出桌面静态资源 | `pnpm build:desktop`（`excel-unprotector-nextjs/`） |
| Rust 单测 | `cargo test`（`excel-unprotector-tauri/src-tauri/`） |

## Docker 部署

```bash
# 构建镜像
docker build -t excel-unprotector-nextjs ./excel-unprotector-nextjs
docker build -t excel-unprotector-server ./excel-unprotector-server

# 运行容器
docker run -d -p 8001:8000 --name excel-unprotector-backend excel-unprotector-server
docker run -d -p 3001:3000 --name excel-unprotector-frontend excel-unprotector-nextjs
```

更完整的部署说明见 [DEPLOYMENT.md](DEPLOYMENT.md)。

## 注意事项

- Web 版在服务器端处理文件，处理完成后立即返回，不在服务器保存；桌面版全程在本机处理，不发起网络请求。
- 仅移除**工作表保护**，不修改单元格内容、公式或样式；不处理工作簿结构保护与 chartsheet。
- 仅支持 `.xlsx`，不支持 `.xlsm` / `.xls`。
- 建议单次处理文件不超过 50MB。

## 开源协议

本项目采用 [MIT License](LICENSE) 开源协议。
