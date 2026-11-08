---
description: Excel Unprotector 前端（Next.js）项目规则：双构建目标、平台适配层与编码约定
globs:
alwaysApply: true
---

# Excel Unprotector 前端项目规则

## 项目定位

Next.js 16 + React 19 前端，一套源码同时产出两种形态：

- **Web 构建**（`pnpm build`）：`output: "standalone"` + `/api` rewrites，由 FastAPI 后端处理文件，Docker 部署。
- **桌面构建**（`pnpm build:desktop`）：`BUILD_TARGET=desktop` → `output: "export"`，静态产物同步到 `../excel-unprotector-tauri/dist/`，由 Tauri WebView 加载并通过 IPC 调用 Rust。

任何改动都必须同时兼容这两种构建，不允许只保证其中一种能跑。

## 常用命令

```bash
pnpm dev              # 开发服务（桌面 dev 时由 Tauri 以 --port 1420 启动）
pnpm lint             # ESLint
pnpm build            # Web：standalone + rewrites
pnpm build:desktop    # 桌面：静态导出 out/ 并同步到 tauri/dist
```

提交前至少执行 `pnpm lint` 与 `pnpm build`；改动涉及桌面链路时再执行 `pnpm build:desktop`。

## 架构约定

- 页面与组件全部是客户端组件（交互密集），需要 hook/事件的组件必须加 `"use client"`。
- 平台差异只允许收敛在 `lib/api/`：
  - `api.ts` 为统一入口，通过 `isTauri()` 分发；
  - `unprotect.web.ts` 用 axios + 锚点下载；
  - `unprotect.tauri.ts` 用 `invoke` + 原生保存对话框。
- **判断是否运行在桌面端，统一使用 `lib/api/api.ts` 中的 `isDesktopRuntime()`**（即 `isTauri()` 的封装），禁止在组件、hook 或其他模块中直接调用 `isTauri()`，也禁止用 `navigator.userAgent`、环境变量等方式自行判断。
- 组件（`components/`）只能调用 `lib/api/api.ts` 暴露的 `uploadFile` / `saveResult`，**禁止**直接 `import` axios 或 `@tauri-apps/api`。
- 业务状态集中在 `lib/hooks/useUpload.ts`，组件不自行拼装上传/保存流程。
- 文件校验统一走 `lib/utils/fileUtils.ts`：仅 `.xlsx`，上限 50MB。新增校验规则需两端行为一致。

## 硬性限制

1. 桌面构建不能注册 `rewrites`，也不能引入 Route Handler / 服务端数据获取 / `next/image` 优化等静态导出不支持的能力。
2. 不要修改 `next.config.ts` 中按 `BUILD_TARGET` 分支的结构；新增配置时同步考虑两种目标。
3. `scripts/export-desktop.mjs` 会清空并全量覆盖 `../excel-unprotector-tauri/dist/`，不要在该目录内手写文件。
4. 保持与后端的契约：`POST /api/unprotect`（Web）与 Tauri 命令 `unprotect_excel`、`save_file`；改动命令名/参数必须同步修改 Rust 端。
5. 桌面端“取消保存”是正常流程（返回 `cancelled`），不要当成错误提示。
6. 错误文案直接面向用户、使用中文，且与 Rust / FastAPI 返回的文案保持一致。

## 代码风格

- TypeScript + App Router，路径别名 `@/*` 指向项目根目录。
- 样式使用 Tailwind CSS 4 + Hero UI 3 + CSS Variables；不引入新的全局样式方案，不新增主题/暗色模式。
- 保持现有组件写法与缩进（4 空格），不进行与任务无关的重构、重命名或格式化。
- 注释只写“为什么”，不复述代码；中文注释与现有风格一致。

## 验证清单

- `pnpm lint` 通过。
- `pnpm build` 通过（Web 目标）。
- 涉及桌面时 `pnpm build:desktop` 成功，`out/index.html` 与 `out/_next/` 存在。
- 手动确认上传、进度/处理中、下载（Web）/ 原生另存为（桌面）、错误提示四条路径未被破坏。
