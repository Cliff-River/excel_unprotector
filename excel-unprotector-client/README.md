# Excel Unprotector Client

一个基于 React 的前端应用，用于上传 Excel 文件并解除工作表保护。

## 技术栈

- React 19
- Vite 8
- TypeScript
- Axios
- React Compiler (babel-plugin-react-compiler)
- pnpm

## 功能特性

- **简单上传**: 支持拖拽上传和点击选择文件
- **安全处理**: 文件在服务器端处理，本地不保存任何数据
- **快速高效**: 免密解除，批量移除工作表保护
- **实时反馈**: 上传进度和处理状态实时显示

## 支持的文件格式

- `.xlsx` 格式的 Excel 文件
- 文件大小限制: 50MB

## 环境要求

- Node.js >= 18
- pnpm

## 安装与运行

### 1. 安装依赖

```bash
pnpm install
```

### 2. 配置环境变量

复制 `.env.example` 文件并修改配置：

```bash
cp .env.example .env
```

编辑 `.env` 文件：

```env
BACKEND_HOST=localhost
BACKEND_PORT=8000
```

### 3. 启动开发服务器

```bash
pnpm dev
```

应用将在 `http://localhost:5173` 运行，并自动代理 API 请求到配置的后端服务。

### 4. 构建生产版本

```bash
pnpm build
```

构建产物将输出到 `dist` 目录。

### 5. 预览生产版本

```bash
pnpm preview
```

## 可用脚本

| 脚本 | 描述 |
|------|------|
| `pnpm dev` | 启动开发服务器 |
| `pnpm build` | 构建生产版本 |
| `pnpm lint` | 使用 oxlint 进行代码检查 |
| `pnpm preview` | 预览生产构建 |
| `pnpm format` | 使用 Prettier 格式化代码 |

## 项目结构

```
src/
├── api/           # API 请求封装
│   └── api.ts
├── components/    # React 组件
│   ├── ErrorContainer.tsx   # 错误状态展示
│   ├── FeatureCard.tsx      # 功能卡片
│   ├── Features.tsx         # 功能介绍区域
│   ├── Footer.tsx           # 页脚
│   ├── Header.tsx           # 页眉
│   ├── ProgressContainer.tsx # 上传/处理进度展示
│   ├── ResultContainer.tsx  # 处理完成结果展示
│   └── UploadZone.tsx       # 文件上传区域
├── hooks/         # 自定义 Hooks
│   └── useUpload.ts         # 上传状态管理
├── types/         # TypeScript 类型定义
│   └── index.ts
├── utils/         # 工具函数
│   └── fileUtils.ts         # 文件验证和大小格式化
├── App.css        # 应用样式
├── App.tsx        # 主应用组件
├── index.css      # 全局样式
└── main.tsx       # 入口文件
```

## 状态管理

应用使用状态机管理上传流程：

- `idle`: 初始状态，等待选择文件
- `uploading`: 文件上传中
- `processing`: 文件上传完成，服务器处理中
- `completed`: 处理完成，可下载结果
- `error`: 发生错误

## 后端集成

前端通过 Vite 代理将以下请求转发到后端：

- `/unprotect`: 上传 Excel 文件并解除保护
- `/health`: 健康检查

后端服务默认运行在 `http://localhost:8000`。

## 代码检查

```bash
pnpm lint
```

## 代码格式化

```bash
pnpm format
```
