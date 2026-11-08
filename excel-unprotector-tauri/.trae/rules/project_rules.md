---
description: Excel Unprotector 桌面端（Tauri v2）项目规则：Rust 解保护核心、IPC 契约与打包约定
globs:
alwaysApply: true
---

# Excel Unprotector 桌面端项目规则

## 项目定位

Tauri v2 桌面壳，**不包含前端源码**。窗口加载的前端是 `excel-unprotector-nextjs` 以 `BUILD_TARGET=desktop` 导出的静态产物，构建时同步到本项目的 `dist/`。

- `dist/` 是生成产物，禁止手工编辑、提交修改或作为源文件使用。
- 文件解保护在本机 Rust 侧完成，不依赖 `excel-unprotector-server`，断网可用。

## 目录结构

```
excel-unprotector-tauri/
├── src-tauri/
│   ├── src/
│   │   ├── excel.rs   # 解保护核心逻辑 + cargo 单元测试
│   │   ├── lib.rs     # Tauri 命令与插件注册
│   │   └── main.rs    # 入口
│   ├── capabilities/  # 权限声明
│   ├── icons/         # 应用图标
│   ├── Cargo.toml
│   └── tauri.conf.json
├── dist/              # 前端静态产物（由 nextjs 构建生成）
└── package.json       # 仅含 @tauri-apps/cli
```

## 常用命令

```bash
pnpm tauri dev        # 启动桌面窗口（自动执行 beforeDevCommand 拉起 Next dev :1420）
pnpm tauri build      # 打包 Linux deb + AppImage
cargo test --manifest-path src-tauri/Cargo.toml   # 解保护核心测试
```

改动 Rust 逻辑必须通过 `cargo test`；改动打包配置再执行 `pnpm tauri build` 验证。

## 构建链路（不可随意改动）

- `beforeDevCommand`: `pnpm --dir ../excel-unprotector-nextjs dev --port 1420`
- `devUrl`: `http://localhost:1420`
- `beforeBuildCommand`: `node ../excel-unprotector-nextjs/scripts/export-desktop.mjs`
- `frontendDist`: `../dist`

`tauri.conf.json` 中这四项与前端脚本强耦合，修改任一项都需要同步校验前端 `package.json` 与 `scripts/export-desktop.mjs`。

## Rust 解保护约束

1. 仅处理 `xl/worksheets/sheetN.xml` 中的 `<sheetProtection>`：
   - 兼容自闭合与成对标签、带命名空间（按 `local_name` 比较）；
   - 不得处理 `xl/workbook.xml` 的 `<workbookProtection>`，也不处理 chartsheet —— 与 openpyxl 行为对齐。
2. 非工作表条目必须用 `by_index_raw` + `raw_copy_file` 原样复制，保留原始压缩方式；只有工作表重新 Deflated 写入。
3. 输出必须是合法 zip：`writer.finish()` 后 `rewind()` 再 `into_inner()`。
4. 错误统一为 `ExcelError`（`InvalidZip` / `InvalidXml` / `NoWorksheet` / `Io`），`Display` 输出面向用户的中文文案，前端直接展示。
5. 测试夹具固定为 `../../excel-unprotector-server/data/protected_file.xlsx`，**只读**，不得修改或删除。
6. 只支持 `.xlsx`；命令入口需校验文件名后缀，大小上限由前端（50MB）保证。

## IPC 契约

与 `excel-unprotector-nextjs/lib/api/unprotect.tauri.ts` 严格对应，改动需两端同时调整：

| 命令 | 入参 | 返回 |
|------|------|------|
| `unprotect_excel` | `fileName: String`, `data: Vec<u8>` | `tauri::ipc::Response`（原始字节） |
| `save_file` | `defaultFileName: String`, `data: Vec<u8>` | `Option<String>`，`None` 表示用户取消 |

- 新增命令必须用 `tauri::generate_handler![]` 注册。
- `save_file` 使用 `tauri-plugin-dialog` 的 `blocking_save_file`，取消返回 `Ok(None)`，写盘失败返回中文错误。
- 新增插件时同步更新 `capabilities/default.json`；CSP 保持最小，确需放开协议/域名时显式修改并说明原因。

## 打包与图标

- 当前仅配置 Linux `deb` + `appimage`；新增平台需同时补充图标与 `bundle.icon` 列表。
- 替换图标时保持 `icons/` 下现有文件名与尺寸，或同步更新 `tauri.conf.json` 的 `icon` 数组。
- `identifier` 为 `app.cliff.excel-unprotector`，非必要不改（影响安装包升级与数据目录）。

## 代码风格

- Rust edition 2024；错误处理用 `Result` + 明确错误类型，禁止 `unwrap()` / `expect()` 出现在非测试代码。
- 不为不可能发生的场景增加兜底分支；保持实现简单、可测试。
- 不进行与任务无关的重构、依赖升级或格式化。
