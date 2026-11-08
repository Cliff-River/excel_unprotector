# Excel Unprotector Next.js

基于 Next.js 16 构建的 Excel 工作表保护解除工具前端。同一套源码支持两种构建目标：

- **Web 构建**（默认）：`output: "standalone"` + `/api` 代理，Docker 部署，由 FastAPI 服务端处理文件。
- **桌面构建**（`BUILD_TARGET=desktop`）：静态导出 `out/`，供 Tauri 桌面壳加载，通过 IPC 调用 Rust 在本地处理。

## 功能特性

- **文件上传** - 支持拖拽和点击上传 Excel 文件
- **实时进度** - Web 版显示上传与处理进度，桌面版显示处理中状态
- **文件下载** - Web 版浏览器下载，桌面版弹出系统原生“另存为”对话框
- **错误处理** - 完善的文件格式验证和错误提示
- **响应式设计** - 适配桌面端和移动端
- **双构建目标** - 一份源码同时产出 Web 与桌面产物，平台差异收敛在 `lib/api`

## 技术栈

- **框架**: Next.js 16 (App Router，页面为客户端渲染)
- **语言**: TypeScript
- **构建工具**: Turbopack
- **样式**: CSS Variables + Tailwind CSS + HeroUI
- **HTTP 客户端**: Axios（Web）/ Tauri IPC（桌面）

## 项目结构

```
excel-unprotector-nextjs/
├── app/                      # Next.js App Router 页面
│   ├── layout.tsx            # 根布局组件
│   ├── page.tsx              # 主页面（"use client"）
│   └── globals.css           # 全局样式
├── components/               # UI 组件
│   ├── NavigationBar/        # 导航栏组件
│   ├── Header.tsx            # 页头组件（含 GitHub 外链）
│   ├── UploadZone.tsx        # 文件上传区域
│   ├── ProgressContainer.tsx # 进度显示组件
│   ├── ResultContainer.tsx   # 结果展示组件
│   ├── ErrorContainer.tsx    # 错误提示组件
│   ├── Features.tsx          # 功能特性展示
│   ├── FeatureCard.tsx       # 功能卡片组件
│   └── Footer.tsx            # 页脚组件
├── i18n/                     # 国际化配置
├── lib/
│   ├── api/
│   │   ├── api.ts            # 平台分发入口（isTauri）
│   │   ├── unprotect.web.ts  # Web 实现：axios + 锚点下载
│   │   ├── unprotect.tauri.ts# 桌面实现：invoke + 原生保存
│   │   └── types.ts          # UploadOptions
│   ├── hooks/
│   │   └── useUpload.ts      # 上传/处理/下载状态管理
│   ├── types/
│   │   └── index.ts          # 通用类型
│   └── utils/
│       └── fileUtils.ts      # 文件验证工具
├── messages/                 # 国际化语言包
├── public/                   # 静态资源
├── scripts/
│   └── export-desktop.mjs    # 桌面静态导出 + 同步到 tauri/dist
├── tests/                    # 测试文件
├── next.config.ts            # Next.js 配置（按 BUILD_TARGET 分支）
├── tsconfig.json             # TypeScript 配置
├── package.json              # 项目依赖
└── Dockerfile                # Docker 构建配置（仅 Web 目标）
```

## 快速开始

### 环境要求

- Node.js >= 22.0.0
- pnpm >= 9.0.0

### 安装依赖

```bash
pnpm install
```

### 开发模式

```bash
pnpm dev
```

启动后访问 http://localhost:3000。

### 构建目标

```bash
pnpm build            # Web：standalone + /api rewrites，用于 Docker
pnpm build:desktop    # 桌面：静态导出 out/ 并同步到 ../excel-unprotector-tauri/dist/
```

`build:desktop` 会以 `BUILD_TARGET=desktop` 执行 `next build`，产出 `out/` 后清空并全量复制到 Tauri 项目的 `dist/`。桌面构建不注册 `rewrites`，不依赖后端。

### Docker 部署（Web）

```bash
# 构建镜像
docker build -t excel-unprotector-nextjs .

# 运行容器
docker run -p 3000:3000 excel-unprotector-nextjs
```

使用根目录 `docker-compose.yml` 时，前端映射到 http://localhost:3001。

## API 配置（Web 构建）

前端应用在 Web 构建中将 `/api/*` 请求代理到后端服务。

### 开发环境

在 `next.config.ts` 的 `rewrites()` 中配置代理目标：

```typescript
async rewrites() {
    return [
        {
            source: "/api/:path*",
            destination: "http://localhost:8000/:path*",
        },
    ];
},
```

### 环境变量

| 变量名 | 默认值 | 描述 |
|--------|--------|------|
| NEXT_PUBLIC_API_BASE_URL | /api/ | API 基础路径（仅 Web 构建） |
| PORT | 3000 | 服务端口 |

### Docker 环境

在 Docker 环境中，API 请求会自动代理到 `http://backend:8000`。

## 平台适配层

`lib/api/api.ts` 在调用时通过 `isTauri()` 选择实现，组件不感知平台：

| 能力 | Web 实现 | 桌面实现 |
|------|----------|----------|
| 处理文件 | axios POST `/api/unprotect` | `invoke("unprotect_excel")` |
| 保存结果 | `<a download>` 锚点下载 | `invoke("save_file")` 原生另存为 |
| 进度 | XHR 上传进度 | 直接进入 processing 状态 |

桌面端入口参数与错误文案详见根目录 `requirements.md`。

## 使用说明

1. **选择文件**: 点击上传区域或拖拽 `.xlsx` 文件到上传区域
2. **文件验证**: 系统会自动验证文件格式和大小（最大 50MB）
3. **上传处理**: Web 版由服务端处理；桌面版由 Rust 在本机处理
4. **下载结果**: 处理完成后，点击"下载文件"按钮保存结果

## 文件格式支持

- **支持**: `.xlsx`（Excel 2007 及以上版本）
- **大小限制**: 50MB

## 常见问题

### Q: 上传文件后显示错误？

A: 请检查：
- 文件格式是否为 `.xlsx`
- 文件大小是否超过 50MB
- Web 模式下后端服务是否正常运行

### Q: Docker 部署后无法连接后端？

A: 确保：
- 前端和后端在同一 Docker 网络中
- 后端服务名称为 `backend`
- 后端端口为 `8000`

### Q: 开发模式下 API 请求失败？

A: 请先启动后端服务：

```bash
cd ../excel-unprotector-server
uv run python -m excel_unprotector_server.main
```

### Q: 桌面构建后页面空白或资源 404？

A: 确认 `excel-unprotector-tauri/dist/` 中已有 `index.html` 与 `_next/` 目录；重新执行 `pnpm build:desktop`，并检查 Tauri 的 `frontendDist` 指向 `../dist`。

### Q: 样式显示异常？

A: 检查是否正确配置了 Tailwind CSS，确保 `globals.css` 文件中包含 `@import "tailwindcss";`

## 构建配置

### Next.js 配置要点

- **App Router**: 使用根目录 `app/` 结构
- **双构建目标**: `BUILD_TARGET=desktop` → `output: "export"`；默认 → `output: "standalone"` + rewrites
- **路径别名**: `@/*` 指向项目根目录
- **API 代理**: 仅 Web 构建通过 `rewrites` 转发请求

## License

MIT
