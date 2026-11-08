# Excel Unprotector 桌面端开发需求

> 文档类型：**可执行开发需求**（技术选型与架构决策已冻结，不需要再做方案对比）
> 日期：2026-11-08
> 交付目标：以 Tauri v2 打包一个**完全离线**的 Linux 桌面应用，复用 `excel-unprotector-nextjs` 的页面与交互；同时保证现有 Web/Docker 部署行为零回退。

---

## 0. 执行约定（AI 实施前必读）

1. 本文所有决策均为最终决策，**不得重新选型**、不得引入未列出的技术方案或依赖。
2. 严格按 §4 的任务顺序实施（T1 → T5），每个任务完成后必须执行该任务列出的验证命令；验证失败必须先修复，不得带着失败进入下一阶段。
3. **禁止修改**：`excel-unprotector-server/**`、根目录 `docker-compose.yml`、`excel-unprotector-nextjs/Dockerfile`、`excel-unprotector-nextjs/app/**`。
4. **禁止修改 UI 结构**：`components/**` 除 `Header.tsx` 外一律不动（`Header.tsx` 仅允许增加桌面端外链处理，见 T4）。
5. 允许新增的依赖仅限：
   - Next.js 侧：`@tauri-apps/api@^2`、`@tauri-apps/plugin-opener@^2`
   - Rust 侧：`zip`、`quick-xml`、`tauri-plugin-dialog`
6. 禁止引入 Python sidecar、Node sidecar、本地 HTTP 服务、远程 API 调用。
7. 不创建 README/说明文档；不执行 `git commit`、`git push`。
8. 代码风格与现有文件保持一致；只写必要的注释，不重构无关代码。
9. 交付时必须输出一份实施报告（写在最终回复中，不落盘），包含：实际执行的验证命令与结果、未验证项（Windows/macOS）、遗留风险。

---

## 1. 目标行为

桌面应用启动后必须满足：

1. 页面样式、HeroUI 组件、字体与 Web 版一致，无资源 404、无白屏。
2. 点击选择或拖拽 `.xlsx` 文件后，由 Rust 在本地完成解保护，**全程不发起任何网络请求**。
3. 结果页显示文件名 `unprotected_<原文件名>`，中文文件名保持正常。
4. 点击“下载文件”弹出系统原生“另存为”对话框；确认后文件落盘且可被 Excel/LibreOffice 正常打开，工作表保护已解除；取消对话框后停留在结果页，不报错。
5. 非 `.xlsx`、超过 50MB、文件损坏三类错误的提示文案与 Web 版一致。
6. 断开网络后上述功能全部可用。
7. 同一套 Next.js 源码仍可通过现有命令构建出 Web 版（standalone + `/api` 代理），Docker 部署链路不受影响。

**非目标**：工作簿结构保护（`<workbookProtection>`）与 chartsheet 保护、`.xlsm`/`.xls`、批量处理、自动更新、代码签名、macOS 打包、i18n、任何服务端改造。

---

## 2. 现状基线（实施前请自行核对，以下为事实）

> 下表及全文路径均相对仓库根目录（`excel_unprotector/`）。

| 位置 | 现状 |
| --- | --- |
| `excel-unprotector-nextjs/next.config.ts` | `output: "standalone"`；`rewrites()` 将 `/api/:path*` 代理到 `backend:8000`（开发态 `localhost:8000`） |
| `excel-unprotector-nextjs/lib/api/api.ts` | 仅有 `uploadFile(file, options)`：axios POST `${API_BASE_URL}unprotect`，`responseType: "blob"`，解析 `Content-Disposition`（优先 `filename*`），网络错误归一化 |
| `excel-unprotector-nextjs/lib/hooks/useUpload.ts` | 状态机 `idle → uploading → processing → completed/error`；`handleDownload` 用 `<a download>` 锚点；负责 object URL 创建与回收 |
| `excel-unprotector-nextjs/lib/utils/fileUtils.ts` | `validateFile`：仅 `.xlsx`，上限 50MB |
| `excel-unprotector-nextjs/app/page.tsx` | `"use client"`，按状态渲染 `UploadZone/ProgressContainer/ResultContainer/ErrorContainer` |
| `excel-unprotector-nextjs/components/Header.tsx` | 含 GitHub 外链（`target="_blank"`），桌面端必须拦截改为系统浏览器打开 |
| `excel-unprotector-tauri/src-tauri/tauri.conf.json` | 脚手架默认值：`beforeDevCommand: "pnpm dev"`、`devUrl: 1420`、`frontendDist: "../dist"`、`csp: null`、窗口 800×600 |
| `excel-unprotector-tauri/src-tauri/src/lib.rs` | 仅 `greet` 命令 + opener 插件 |
| `excel-unprotector-tauri/src-tauri/Cargo.toml` | `tauri 2`、`tauri-plugin-opener 2`、serde；edition 2024 |
| `excel-unprotector-tauri/src/`、`index.html`、`vite.config.ts`、`public/` | Vite + React 模板，待清理 |
| `excel-unprotector-server/src/excel_unprotector_server/sheet_unprotect.py` | 唯一业务逻辑：遍历 `wb.worksheets` 执行 `sheet.protection.disable()` 后保存 |
| `excel-unprotector-server/src/excel_unprotector_server/main.py` | 输出文件名规则 `unprotected_{filename}`，MIME `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` |
| `excel-unprotector-server/data/protected_file.xlsx` | Rust 单测夹具（只读，不得改动） |

---

## 3. 已冻结的技术决策

| # | 决策 | 结果 |
| --- | --- | --- |
| D1 | 解保护实现方式 | Rust 原生：`zip` 重写 + `quick-xml` 剥离 `<sheetProtection>`；不携带 Python |
| D2 | Next.js 接入方式 | 同一源码、双构建目标：`BUILD_TARGET=desktop` → `output: "export"`；未设置 → 现有 standalone + rewrites |
| D3 | 静态产物落点 | `next build` 的 `out/` 由 Node 脚本同步到 `excel-unprotector-tauri/dist/`，Tauri `frontendDist` 保持 `../dist` |
| D4 | 结果保存方式 | Rust 命令经 `tauri-plugin-dialog` 弹原生另存为，`std::fs::write` 落盘；不使用 `<a download>` |
| D5 | 平台差异位置 | 全部收敛在 `lib/api/` 适配层 + `useUpload` 两处小改；页面与 UI 组件不感知平台 |
| D6 | IPC 通道 | Tauri v2 原始字节：JS 传 `Uint8Array`，Rust 收 `Vec<u8>`；返回 `tauri::ipc::Response` → JS `ArrayBuffer` |

---

## 4. 实施任务

### T1 Next.js 双构建目标与桌面导出脚本

**修改** `excel-unprotector-nextjs/next.config.ts`

- 读取 `const isDesktop = process.env.BUILD_TARGET === "desktop"`。
- desktop 分支：`output: "export"`、`images: { unoptimized: true }`、**不注册** `rewrites`。
- web 分支：`output: "standalone"`，`rewrites()` 原样保留（含 production/development 两套 destination）。
- 不得预设 `trailingSlash`、`assetPrefix`（仅在 §10 R1 实测失败后按预案启用）。
- 参考结构：

```ts
const isDesktop = process.env.BUILD_TARGET === "desktop";

const nextConfig: NextConfig = {
    output: isDesktop ? "export" : "standalone",
    ...(isDesktop ? { images: { unoptimized: true } } : {}),
    sassOptions: { includePaths: ["./"] },
    ...(isDesktop
        ? {}
        : {
              async rewrites() {
                  /* 现有实现原样保留 */
              },
          }),
};
```

**新增** `excel-unprotector-nextjs/scripts/export-desktop.mjs`

必须满足：

1. 以 `import.meta.url` 解析 `nextjsRoot` 与 `tauriDist = <repo>/excel-unprotector-tauri/dist`，**不依赖 process.cwd()**。
2. 以 `BUILD_TARGET=desktop` 环境变量执行 `pnpm exec next build`（通过 `spawnSync`，Windows 使用 `pnpm.cmd`；`stdio: "inherit"`），失败时以非 0 退出。
3. 构建后断言 `out/index.html` 存在，否则报错退出。
4. 先 `rmSync(tauriDist, { recursive: true, force: true })`，再 `mkdirSync` + `cpSync(out, tauriDist, { recursive: true })`，保证无旧产物残留。
5. 输出同步条目数与目标路径；全部使用 `node:fs`/`node:path`，不得使用 bash 专有命令。

**修改** `excel-unprotector-nextjs/package.json`

- `scripts` 新增：`"build:desktop": "node scripts/export-desktop.mjs"`。
- `dependencies` 新增：`"@tauri-apps/api": "^2"`、`"@tauri-apps/plugin-opener": "^2"`。
- 执行 `pnpm install` 更新 lockfile。

**验证**

```bash
cd excel-unprotector-nextjs
pnpm lint
pnpm build            # Web：standalone 构建成功
pnpm build:desktop    # 生成 out/ 并同步到 ../excel-unprotector-tauri/dist/
ls out/index.html ../excel-unprotector-tauri/dist/index.html
```

确认 desktop 构建未启用 `rewrites`（构建不报错即证明），且运行时不会发起 `/api/unprotect` 请求（在 dev 窗口网络面板验证）。Web 实现被一起打包不影响判定，运行时由 `isTauri()` 分流。

---

### T2 Tauri 壳清理与配置

**删除**（Vite 模板残留）：

```
excel-unprotector-tauri/src/
excel-unprotector-tauri/index.html
excel-unprotector-tauri/vite.config.ts
excel-unprotector-tauri/tsconfig.json
excel-unprotector-tauri/tsconfig.node.json
excel-unprotector-tauri/public/
```

**改写** `excel-unprotector-tauri/package.json` 为仅保留 Tauri CLI：

```json
{
  "name": "excel-unprotector-tauri",
  "private": true,
  "version": "0.1.0",
  "type": "module",
  "scripts": { "tauri": "tauri" },
  "devDependencies": { "@tauri-apps/cli": "^2" }
}
```

随后在 `excel-unprotector-tauri/` 执行 `pnpm install`（清理 React/Vite 依赖与 lockfile）。

**改写** `excel-unprotector-tauri/src-tauri/tauri.conf.json` 为目标值：

```json
{
  "$schema": "https://schema.tauri.app/config/2",
  "productName": "Excel Unprotector",
  "version": "0.1.0",
  "identifier": "app.cliff.excel-unprotector",
  "build": {
    "beforeDevCommand": "pnpm --dir ../excel-unprotector-nextjs dev --port 1420",
    "devUrl": "http://localhost:1420",
    "beforeBuildCommand": "node ../excel-unprotector-nextjs/scripts/export-desktop.mjs",
    "frontendDist": "../dist"
  },
  "app": {
    "windows": [
      {
        "label": "main",
        "title": "Excel Unprotector",
        "width": 960,
        "height": 720,
        "minWidth": 800,
        "minHeight": 600,
        "resizable": true,
        "dragDropEnabled": true
      }
    ],
    "security": {
      "csp": "default-src 'self'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; font-src 'self' data:; connect-src 'self' ipc: http://ipc.localhost"
    }
  },
  "bundle": {
    "active": true,
    "targets": ["deb", "appimage"],
    "icon": [
      "icons/32x32.png",
      "icons/128x128.png",
      "icons/128x128@2x.png",
      "icons/icon.icns",
      "icons/icon.ico"
    ]
  }
}
```

- 生产 `csp` 不得为 `null`。若开发态 HMR 因 CSP 异常，可单独设置 `app.security.devCsp` 放宽，但不得改动生产 `csp`。
- `bundle.targets` 固定为 `["deb", "appimage"]`；Windows 构建时用 `--bundles nsis` 覆盖，本期不在 Linux 上产出。
- `capabilities/default.json` 保持 `windows: ["main"]`，权限仅 `"core:default"`、`"opener:default"`；**不开放** fs 任意路径权限（保存由 Rust 侧完成，不经过 ACL）。

**验证**

```bash
cd excel-unprotector-tauri
pnpm tauri dev    # 窗口显示 Next.js 页面（此时上传功能尚未改造，仅验证加载/样式/拖拽）
```

---

### T3 Rust 解保护模块与命令

**依赖**（在 `src-tauri/` 执行 `cargo add`，以解析到的兼容版本为准并提交 `Cargo.lock`）：

| crate | 用途 |
| --- | --- |
| `zip` | 读取与重写 xlsx 归档，raw copy 非工作表条目 |
| `quick-xml` | 流式读写工作表 XML，剥离 `sheetProtection` |
| `tauri-plugin-dialog` | Rust 侧原生“另存为”对话框 |

保留 `tauri-plugin-opener`；保留 `serde`、`serde_json`。

**前置验证（在写业务代码前完成）**：临时增加一个 `echo_bytes(data: Vec<u8>) -> tauri::ipc::Response` 命令，用 dev 窗口以 1MB `Uint8Array` 调用并校验往返字节一致，通过后删除该命令。若当前 Tauri 版本不支持该通道，立即按 §10 R2 改用 base64 方案并在报告中记录。

**新增** `src-tauri/src/excel.rs`

```rust
pub enum ExcelError {
    InvalidZip,
    InvalidXml,
    NoWorksheet,
    Io(String),
}

impl std::fmt::Display for ExcelError { /* 见 §5.3 文案 */ }

/// 输入 xlsx 字节，输出解除工作表保护后的 xlsx 字节
pub fn unprotect(input: &[u8]) -> Result<Vec<u8>, ExcelError>;
```

算法必须严格按 §6 执行；不得对单元格、公式、样式、共享字符串做任何重算或改写；不得处理 chartsheet / workbookProtection。

**修改** `src-tauri/src/lib.rs`

- 删除 `greet`。
- 注册 `.plugin(tauri_plugin_opener::init())`、`.plugin(tauri_plugin_dialog::init())`。
- 注册命令 `unprotect_excel`、`save_file`（签名见 §5.1、§5.2）。

**单元测试**（`excel.rs` 内 `#[cfg(test)]`，夹具用 `std::fs::read` 读取 `CARGO_MANIFEST_DIR/../../excel-unprotector-server/data/protected_file.xlsx`，输出只写系统临时目录）：

1. 正常处理：输出可被 `ZipArchive` 打开；条目数量与名称集合与输入一致；所有 `xl/worksheets/sheet<N>.xml` 中不存在 `sheetProtection`。
2. 非工作表条目：处理后字节与输入完全一致，且压缩方式一致（证明 raw copy）。
3. 内容保真：对每个工作表，按顺序提取输入/输出的 XML 事件本地名（Start/Empty/End/Text），过滤掉 `sheetProtection` 后序列必须完全一致。
4. 幂等：对第 1 步输出再次调用 `unprotect`，仍成功且无 `sheetProtection`。
5. 异常输入：随机非 zip 字节 → `InvalidZip`；无 `xl/worksheets/sheet*.xml` 的合法 zip → `NoWorksheet`；两者均不得 panic。
6. 非法 XML 的工作表 → `InvalidXml`。

**验证**

```bash
cd excel-unprotector-tauri/src-tauri
cargo test
cargo build
```

---

### T4 前端平台适配层

**新增** `excel-unprotector-nextjs/lib/api/types.ts`

```ts
export interface UploadOptions {
    onProgress?: (progress: number) => void;
    onStatusChange?: (status: "uploading" | "processing") => void;
}
```

**新增** `lib/api/unprotect.web.ts`

- 将 `excel-unprotector-nextjs/lib/api/api.ts` 中现有 `uploadFile` 实现**逐行迁入**（axios、超时、进度、`Content-Disposition` 解析、错误归一化全部保持不变）。
- 新增 `saveResult(blob: Blob, fileName: string): Promise<"saved" | "cancelled">`：用临时 object URL + `<a download>` 触发下载，`finally` 中 `URL.revokeObjectURL`，返回 `"saved"`。

**新增** `lib/api/unprotect.tauri.ts`

- `uploadFile(file, options?)`：
  1. 不发起任何 HTTP 请求；
  2. `const data = new Uint8Array(await file.arrayBuffer())`；
  3. `const buf = await invoke<ArrayBuffer>("unprotect_excel", { fileName: file.name, data })`；
  4. 若返回值为 `Uint8Array` 则取其 `buffer`，统一转 `ArrayBuffer`；
  5. 返回 `{ blob: new Blob([buf], { type: XLSX_MIME }), fileName: "unprotected_" + file.name }`；
  6. 捕获字符串错误原样抛出；非字符串统一抛 `new Error("文件处理失败")`。
- `saveResult(blob, fileName)`：`invoke<string | null>("save_file", { defaultFileName: fileName, data: new Uint8Array(await blob.arrayBuffer()) })`；返回非空字符串 → `"saved"`，`null` → `"cancelled"`；异常的 message 固定为 `"保存文件失败，请重试"`。
- 常量 `XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"`。

**改写** `lib/api/api.ts` 为分发入口（不新增 `processFile`，继续沿用 `uploadFile` 名称以减小改动面）：

```ts
import { isTauri } from "@tauri-apps/api/core";
import * as web from "./unprotect.web";
import * as tauri from "./unprotect.tauri";

export function isDesktopRuntime(): boolean {
    return isTauri();
}

export const uploadFile: typeof web.uploadFile = (file, options) =>
    (isTauri() ? tauri.uploadFile : web.uploadFile)(file, options);

export const saveResult: typeof web.saveResult = (blob, fileName) =>
    (isTauri() ? tauri.saveResult : web.saveResult)(blob, fileName);
```

要求：判定必须发生在**函数调用时**，不得在模块顶层缓存结果（静态导出期间会在 Node 环境执行一次）。

**修改** `lib/hooks/useUpload.ts`（仅以下三处）

1. 新增 `const blobRef = useRef<Blob | null>(null)`；成功时保存 `result.blob`；`handleReset` 与卸载时置 `null`。
2. 初始状态：`const desktop = isDesktopRuntime()`，桌面态初始 `status: "processing"`（跳过 uploading），Web 态保持 `"uploading"`；桌面态忽略 `onProgress` 回调。
3. `handleDownload` 改为 async：先判空 `blobRef.current` 并直接返回；再 `await saveResult(blobRef.current, uploadState.fileName)`；返回 `"cancelled"` 时不做任何状态变更；抛错时切到 `status: "error"` 并展示错误 message（复用 `ErrorContainer`，不新增组件）。

其余状态机、object URL 生命周期、`downloadUrl` 字段保持原样。

**修改** `components/Header.tsx`（唯一允许改动的组件）

- 该文件存在 GitHub 外链（`target="_blank"`），桌面 WebView 内直接跳转会离开应用，必须拦截。
- 增加 `"use client"`；`onClick` 中若 `isDesktopRuntime()` 为真则 `e.preventDefault()` 并调用 `openUrl(GITHUB_URL)`（`@tauri-apps/plugin-opener`）；Web 态不拦截，行为不变。

**验证**

```bash
cd excel-unprotector-nextjs
pnpm lint
pnpm build
pnpm build:desktop
cd ../excel-unprotector-tauri
pnpm tauri dev
```

在 dev 窗口完成：选择文件 → 结果页 → 下载 → 取消/保存；再验证 Web 态 `pnpm dev` 下原上传下载流程不变。

---

### T5 打包与验收

```bash
cd excel-unprotector-tauri
pnpm tauri build          # 先执行 export-desktop.mjs，再编译并产出 deb/AppImage
```

产物路径：`src-tauri/target/release/bundle/deb/`、`.../appimage/`。安装/运行后按 §8 逐条执行验收。

---

## 5. 接口契约

### 5.1 `unprotect_excel`

```rust
#[tauri::command]
fn unprotect_excel(file_name: String, data: Vec<u8>) -> Result<tauri::ipc::Response, String>;
```

- `file_name` 为空或不以 `.xlsx`（大小写不敏感）结尾 → `Err("只支持 .xlsx 格式的 Excel 文件")`。
- 其余错误按 §5.3 返回。
- 成功时返回 `tauri::ipc::Response::new(bytes)`，JS 侧得到 `ArrayBuffer`。

### 5.2 `save_file`

```rust
#[tauri::command]
fn save_file(
    app: tauri::AppHandle,
    default_file_name: String,
    data: Vec<u8>,
) -> Result<Option<String>, String>;
```

- 通过 `tauri_plugin_dialog` 的 Rust 侧 blocking save 对话框，`set_file_name(default_file_name)`，添加 Excel 过滤器（`xlsx`）。
- 用户取消 → `Ok(None)`；成功 → `std::fs::write` 后返回 `Ok(Some(绝对路径))`。
- 写盘失败 → `Err("保存文件失败，请重试")`。
- 不得 panic；不经过 fs 插件 scope。

### 5.3 错误文案

| 情形 | 返回文案 |
| --- | --- |
| zip 打开失败 | `文件处理失败：文件格式无效或已损坏` |
| 无 `xl/worksheets/sheet*.xml` | `文件处理失败：文件格式无效或已损坏` |
| 工作表 XML 解析失败 | `文件处理失败：无法读取工作表内容` |
| 其他 IO 错误 | `文件处理失败` |
| 非 `.xlsx` 文件名 | `只支持 .xlsx 格式的 Excel 文件` |
| 保存对话框写盘失败 | `保存文件失败，请重试` |
| 用户取消保存 | 非错误，返回 `null` |

### 5.4 TypeScript 统一入口

```ts
uploadFile(file: File, options?: UploadOptions): Promise<UploadResult>
saveResult(blob: Blob, fileName: string): Promise<"saved" | "cancelled">
```

`UploadResult` 沿用 `{ blob: Blob; fileName: string }`（`lib/types` 不改）。

---

## 6. 解保护算法规格

1. 用 `Cursor<&[u8]>` + `ZipArchive` 打开输入；失败 → `InvalidZip`。
2. 创建 `ZipWriter<Cursor<Vec<u8>>>`，**按原顺序**遍历条目（`0..archive.len()`）。
3. 条目名判定为工作表：`starts_with("xl/worksheets/sheet")` 且 `ends_with(".xml")` 且中间部分全部为 ASCII 数字。`xl/worksheets/_rels/sheet1.xml.rels` 等不匹配，走 raw copy。
4. 工作表条目：
   - 读取该条目全部字节；
   - 用 `quick_xml::Reader` 逐事件读取，**显式关闭文本 trim**，不调用 `unescape`；
   - 用 `quick_xml::Writer` 写出；遇到本地名为 `sheetProtection` 的 `Start`/`Empty` 事件时丢弃；若为 `Start`，持续丢弃事件直到同名 `End`（支持嵌套深度计数，兼容 `<sheetProtection>...</sheetProtection>` 与自闭合两种形态）；
   - 其余事件（Decl、PI、Comment、Start/Empty/End、Text、CData）原样写出，保持属性与命名空间前缀不变；
   - 若 EOF 时仍处于跳过状态 → `InvalidXml`；
   - 用 Deflated 压缩写入同名条目（`zip` 2.x 使用 `SimpleFileOptions`，以实际解析版本的类型名为准）。
5. 其他所有条目（`[Content_Types].xml`、`_rels`、`sharedStrings`、`styles`、目录条目等）：`by_index_raw` + `raw_copy_file`，保留原压缩方式、条目名与字节内容，禁止反序列化重写。
6. 遍历结束后若未匹配到任何工作表条目 → `NoWorksheet`。
7. `finish()` 后取回 `Vec<u8>` 返回。

边界：只处理 worksheet；不移除 `workbookProtection`；不处理 chartsheet；不校验/重算公式；不修改任何单元格内容或样式。

---

## 7. 构建与运行命令

| 场景 | 命令（目录） |
| --- | --- |
| 桌面开发（HMR） | `pnpm tauri dev`（`excel-unprotector-tauri/`，自动拉起 next dev :1420） |
| 桌面生产包 | `pnpm tauri build`（`excel-unprotector-tauri/`） |
| 仅验证静态导出 | `pnpm build:desktop`（`excel-unprotector-nextjs/`） |
| Web 回归 | `pnpm lint && pnpm build`（`excel-unprotector-nextjs/`） |
| Docker 回归 | 根目录 `docker compose build`（配置不得改动） |
| Rust 单测 | `cargo test`（`excel-unprotector-tauri/src-tauri/`） |

版本号三处必须一致为 `0.1.0`：`excel-unprotector-nextjs/package.json`、`tauri.conf.json`、`Cargo.toml`。

---

## 8. 验收用例（必须逐条执行并记录结果）

### 8.1 Rust 层

- [ ] `cargo test` 全绿，且 §4 T3 的 6 项断言全部覆盖。
- [ ] 用 openpyxl 二次加载处理结果（有 Python 环境时执行；无则记录为未验证）：
  - 输出可被 `load_workbook` 打开；
  - 每个 `sheet.protection.sheet` 为 `False`；
  - 工作表数量、单元格值与原文件一致。

### 8.2 Web 回归（证明双目标未回退）

- [ ] `pnpm lint`、`pnpm build` 通过。
- [ ] 浏览器中上传 `.xlsx` → 进度 → 结果页 → 下载 → 重置全流程与改造前一致。
- [ ] 网络错误、500、超时提示文案与改造前一致。
- [ ] `docker compose build` 成功（不要求启动）。

### 8.3 桌面端（`tauri dev` 与 `tauri build` 产物各测一次）

- [ ] 首屏样式、HeroUI 组件、字体与 Web 版一致；控制台无资源 404、无 CSP 拦截报错。
- [ ] 点击选择 `.xlsx` 可用；拖拽 `.xlsx` 到窗口可用（Linux WebKitGTK 必测）。
- [ ] 处理后结果页文件名 `unprotected_<原文件名>`，中文文件名正常。
- [ ] 点击“下载文件”弹出原生保存框；保存后文件可被 Excel/LibreOffice 打开且保护已解除。
- [ ] 保存框点“取消”后停留在结果页，无错误提示。
- [ ] 非 `.xlsx`、超过 50MB、损坏文件三类提示与 Web 版一致。
- [ ] “处理其他文件”复位正常；连续处理 3 个不同文件无残留、无内存异常增长。
- [ ] 断开网络（或禁用网卡）后全流程仍可用。
- [ ] 点击 Header 的 GitHub 图标在系统浏览器打开，应用窗口不跳转、不白屏。
- [ ] AppImage 可执行；deb 可安装并启动。

---

## 9. 边界与禁止事项

1. 不得移除 `.xlsm`/`.xls` 拦截；前端 `validateFile` 保持唯一入口校验。
2. 不得把 `frontendDist` 指向 nextjs 项目内的 `out/`，必须复制到 `excel-unprotector-tauri/dist/`。
3. 不得为桌面端新增页面、弹窗组件或进度组件；进度复用 `ProgressContainer`。
4. 不得让 `lib/hooks` 或 `components` 直接依赖 `zip`/Rust 细节；差异只允许出现在 `lib/api/`。
5. 不得擅自扩大范围去“顺便兼容”`workbookProtection`、chartsheet 等本期明确排除的边界。
6. 不得在 Rust 侧使用 `unwrap()`/`expect()` 处理用户输入（测试代码除外）。
7. 不得在 Web 构建中引入 Node-only 模块或 Tauri 专属 API 的顶层调用。

---

## 10. 风险处置（先实测，失败才启用备选）

| # | 必须验证的点 | 失败时的唯一允许动作 |
| --- | --- | --- |
| R1 | 静态导出在 `tauri://localhost` 下 `/_next/...` 资源加载 | 仅 desktop 分支启用 `trailingSlash: true` 与相对 `assetPrefix`，并回归 Web 构建 |
| R2 | 当前 Tauri 小版本 `Uint8Array`/`Response` 原始字节通道 | 回退 base64（serde）传递，并在报告中记录；不得因此改变前端接口 |
| R3 | Linux WebKitGTK 拖拽是否被 Tauri 窗口拖放事件吞掉 | 仅在 `lib/api` 或 `useUpload` 内监听 `onDragDropEvent` 兜底，不改 `UploadZone` |
| R4 | Next 16 Turbopack 静态导出 + HeroUI 3 兼容性 | desktop 编排脚本可回退 `next build --webpack`，Web 构建不动 |
| R5 | 收紧 CSP 后内联脚本/样式被拦截 | 按脚本/样式分类实测调整 `csp`；生产环境不得留 `null` |
| R6 | 极端 OOXML 变体（异常命名、chartsheet） | 非工作表条目一律 raw copy，仅按精确路径处理，不做扩展 |
| R7 | 双构建目标漂移 | 交付前同时跑 `pnpm build` 与 `pnpm build:desktop`；Docker 配置锁死 |

---

## 11. 交付物与完成定义

**代码交付**

- `excel-unprotector-nextjs`：`next.config.ts`、`package.json`、`scripts/export-desktop.mjs`、`lib/api/{api.ts,types.ts,unprotect.web.ts,unprotect.tauri.ts}`、`lib/hooks/useUpload.ts`、`components/Header.tsx`。
- `excel-unprotector-tauri`：删除 Vite 模板；`package.json` 瘦身；`src-tauri/{Cargo.toml,Cargo.lock,tauri.conf.json,capabilities/default.json,src/lib.rs,src/excel.rs}`；`dist/` 为构建产物（已被 `.gitignore` 覆盖，不提交）。
- 不改动 `excel-unprotector-server/**`、Docker 相关文件、`app/**`、其余 `components/**`。

**完成定义**

- [ ] `cargo test` 全部通过。
- [ ] `pnpm build`（Web）与 `pnpm build:desktop`（desktop）均成功。
- [ ] `pnpm tauri dev` 页面正常，`pnpm tauri build` 产出 deb/AppImage。
- [ ] §8 验收用例逐条通过，并附执行记录。
- [ ] 仓库中无 `greet`、Vite 模板文件、模板示例资源残留。
- [ ] 最终回复给出：改动文件清单、验证命令与结果、未验证项（Windows/macOS、openpyxl 校验等）。
