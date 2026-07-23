# Excel Unprotector Next.js

基于 Next.js 16 构建的 Excel 文件工作表保护解除工具前端应用。

## 功能特性

- **文件上传** - 支持拖拽和点击上传 Excel 文件
- **实时进度** - 显示文件上传和处理进度
- **文件下载** - 处理完成后一键下载解除保护的文件
- **错误处理** - 完善的文件格式验证和错误提示
- **响应式设计** - 适配桌面端和移动端
- **服务端渲染** - 使用 Next.js App Router 实现 SSR

## 技术栈

- **框架**: Next.js 16 (App Router)
- **语言**: TypeScript
- **构建工具**: Turbopack
- **样式**: CSS Variables + Tailwind CSS
- **HTTP 客户端**: Axios

## 项目结构

```
excel-unprotector-nextjs/
├── src/
│   ├── app/                  # Next.js App Router 页面
│   │   ├── layout.tsx        # 根布局组件
│   │   ├── page.tsx          # 主页面
│   │   └── globals.css       # 全局样式
│   ├── components/           # UI 组件
│   │   ├── Header.tsx        # 页头组件
│   │   ├── UploadZone.tsx    # 文件上传区域
│   │   ├── ProgressContainer.tsx  # 进度显示组件
│   │   ├── ResultContainer.tsx    # 结果展示组件
│   │   ├── ErrorContainer.tsx     # 错误提示组件
│   │   ├── Features.tsx      # 功能特性展示
│   │   ├── FeatureCard.tsx   # 功能卡片组件
│   │   └── Footer.tsx        # 页脚组件
│   ├── hooks/                # 自定义 Hooks
│   │   └── useUpload.ts      # 文件上传状态管理
│   ├── api/                  # API 请求封装
│   │   └── api.ts            # 文件上传 API
│   ├── types/                # TypeScript 类型定义
│   │   └── index.ts          # 通用类型
│   └── utils/                # 工具函数
│       └── fileUtils.ts      # 文件验证工具
├── public/                   # 静态资源
├── next.config.ts            # Next.js 配置
├── tsconfig.json             # TypeScript 配置
├── package.json              # 项目依赖
└── Dockerfile                # Docker 构建配置
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
pnpm run dev
```

启动后访问: http://localhost:3000

### 生产构建

```bash
pnpm run build
pnpm run start
```

### Docker 部署

```bash
# 构建镜像
docker build -t excel-unprotector-nextjs .

# 运行容器
docker run -p 3000:3000 excel-unprotector-nextjs
```

## API 配置

前端应用默认将 `/api/*` 请求代理到后端服务。

### 开发环境

在 `next.config.ts` 中配置代理目标：

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
| NEXT_PUBLIC_API_BASE_URL | /api/ | API 基础路径 |
| PORT | 3000 | 服务端口 |

### Docker 环境

在 Docker 环境中，API 请求会自动代理到 `http://backend:8000`。

## 使用说明

1. **选择文件**: 点击上传区域或拖拽 `.xlsx` 文件到上传区域
2. **文件验证**: 系统会自动验证文件格式和大小（最大 50MB）
3. **上传处理**: 文件上传后，服务端会自动移除工作表保护
4. **下载结果**: 处理完成后，点击"下载文件"按钮获取解除保护的文件

## 文件格式支持

- **支持**: `.xlsx` (Excel 2007 及以上版本)
- **大小限制**: 50MB

## 常见问题

### Q: 上传文件后显示错误？

A: 请检查：
- 文件格式是否为 `.xlsx`
- 文件大小是否超过 50MB
- 后端服务是否正常运行

### Q: Docker 部署后无法连接后端？

A: 确保：
- 前端和后端在同一 Docker 网络中
- 后端服务名称为 `backend`
- 后端端口为 `8000`

### Q: 开发模式下 API 请求失败？

A: 请先启动后端服务：

```bash
cd ../excel_unprotector_server
uv run uvicorn main:app --host 0.0.0.0 --port 8000
```

### Q: 样式显示异常？

A: 检查是否正确配置了 Tailwind CSS，确保 `globals.css` 文件中包含 `@import "tailwindcss";`

## 构建配置

### Next.js 配置要点

- **App Router**: 使用 `src/app/` 目录结构
- **Standalone Output**: 启用 `output: "standalone"` 优化 Docker 镜像大小
- **路径别名**: `@/*` 指向 `src/` 目录
- **API 代理**: 通过 `rewrites` 配置实现请求转发

## License

MIT