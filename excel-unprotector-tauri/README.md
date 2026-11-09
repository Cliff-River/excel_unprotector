# Excel Unprotector 桌面端（Tauri v2）

基于 Tauri v2 的**完全离线**桌面客户端。UI 复用 `excel-unprotector-nextjs` 的静态导出物，工作表解保护由 Rust 原生完成（`zip` + `quick-xml`），处理结果通过系统原生“另存为”对话框保存。无需后端服务、无需网络。

## 功能

- 选择或拖拽 `.xlsx` 文件，本地移除所有工作表的保护
- 结果显示为 `unprotected_<原文件名>`，中文文件名正常
- 点击“下载文件”弹出原生保存对话框，取消后停留在结果页
- 与 Web 版一致的错误提示（非 `.xlsx`、超过 50MB、文件损坏）
- 离线可用；不含任何远程请求

## 架构

```
┌──────────────────────── Tauri WebView ────────────────────────┐
│  excel-unprotector-tauri/dist/（Next.js 静态导出）            │
│    UI ── invoke("unprotect_excel") ──┐                        │
│    UI ── invoke("save_file") ──┐     │                        │
└────────────────────────────────┼─────┼────────────────────────┘
                                 │     │ IPC 原始字节
                                 ▼     ▼
                    Rust 命令：解保护（zip 重写 + 剔除 <sheetProtection>）
                               保存（原生另存为 + std::fs::write）
```

平台分流在前端 `lib/api/api.ts` 中通过 `isTauri()` 完成：Web 构建仍走 axios + `/api` 代理，桌面构建走 IPC，页面组件与交互不变。

## 目录结构

```
excel-unprotector-tauri/
├── dist/                     # 构建生成：Next.js 静态导出物（已 gitignore）
├── package.json              # 仅保留 @tauri-apps/cli
└── src-tauri/
    ├── src/
    │   ├── lib.rs            # 命令注册：unprotect_excel / save_file
    │   └── excel.rs          # 解保护算法 + 单元测试
    ├── capabilities/default.json
    ├── tauri.conf.json       # 窗口、CSP、构建编排
    ├── Cargo.toml
    └── icons/
```

## 环境要求

- Node.js >= 22、pnpm >= 9
- Rust stable 工具链
- Linux 系统依赖（Tauri v2）：

```bash
sudo apt install libwebkit2gtk-4.1-dev libgtk-3-dev librsvg2-dev \
  libayatana-appindicator3-dev build-essential curl wget file libssl-dev
```

## 开发

```bash
cd excel-unprotector-tauri
pnpm install
pnpm tauri dev
```

`beforeDevCommand` 会自动在 `1420` 端口启动 Next.js 开发服务器（HMR 生效），Tauri 窗口加载 `http://localhost:1420`。此模式下处理与保存已走真实 IPC，可直接联调。

## 构建

```bash
cd excel-unprotector-tauri
pnpm tauri build
```

构建前会自动执行 `excel-unprotector-nextjs/scripts/export-desktop.mjs`：以 `BUILD_TARGET=desktop` 执行 `next build`，将 `out/` 全量同步到本项目的 `dist/`，再编译 Rust 并打包。

产物：

- `src-tauri/target/release/bundle/deb/Excel Unprotector_0.1.0_amd64.deb`
- `src-tauri/target/release/bundle/appimage/Excel Unprotector_0.1.0_amd64.AppImage`

仅验证前端静态导出：

```bash
cd ../excel-unprotector-nextjs
pnpm build:desktop   # 生成 out/ 并同步到 excel-unprotector-tauri/dist/
```

## 测试

```bash
cd src-tauri
cargo test
```

覆盖：

- 夹具处理后条目名称/数量不变、工作表不再含 `sheetProtection`
- 非工作表条目字节与压缩方式保持一致（raw copy）
- 工作表 XML 除 `sheetProtection` 外事件序列完全一致
- 幂等（对已解保护文件再次处理不损坏）
- 非 zip 字节、无工作表 zip、非法工作表 XML 均返回错误而非 panic

夹具位于 `../excel-unprotector-server/data/protected_file.xlsx`（只读）。

## 实现要点

- 仅匹配 `xl/worksheets/sheet<N>.xml`；其余条目（`[Content_Types].xml`、rels、sharedStrings、styles 等）原样 raw copy。
- 按 XML 本地名匹配 `sheetProtection`，兼容自闭合与成对标签、命名空间前缀。
- 保存对话框由 Rust 侧调用，不经过 fs 插件 scope，capabilities 仅保留 `core:default` 与 `opener:default`。
- Header 的 GitHub 外链在桌面端由 `@tauri-apps/plugin-opener` 在系统浏览器打开，避免 WebView 内跳转。

## 边界

- 仅支持 `.xlsx`；不支持 `.xlsm` / `.xls`
- 仅移除**工作表保护**，不处理工作簿结构保护（`<workbookProtection>`）与 chartsheet 保护
- 本期仅构建 Linux 产物；Windows（NSIS）与 macOS（仅 `.app`，未验证）需在对应平台构建
- 图标仍为脚手架默认值，正式发布前需替换
