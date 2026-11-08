# 需求方案：基于 Next.js 构建物的 Excel Unprotector Tauri 桌面端

> 状态：**方案待评审，暂不实施开发**
> 日期：2026-11-08
> 涉及项目：`excel-unprotector-nextjs/`（前端）、`excel-unprotector-tauri/`（Tauri 壳）、`excel-unprotector-server/`（现有 Python 后端，仅作行为基准）

---

## 1. 背景与目标

现有 Excel Unprotector 由三部分组成：

| 项目 | 技术栈 | 现状 |
| --- | --- | --- |
| `excel-unprotector-nextjs` | Next.js 16 (App Router) / React 19 / HeroUI 3 / Tailwind 4 / axios | 纯客户端页面，上传 `.xlsx` → `/api/unprotect` → 下载结果；`output: "standalone"`，Docker 部署 |
| `excel-unprotector-server` | FastAPI + openpyxl | 唯一业务逻辑：遍历工作表 `sheet.protection.disable()` 后重新保存 xlsx |
| `excel-unprotector-tauri` | Tauri v2 + Vite + React 模板 | 刚初始化，仅有 greet 示例，未做任何业务改造 |

**目标**：复用 nextjs 项目的页面与交互作为桌面端 UI，以 Tauri v2 打包为**完全离线、无需后端服务**的桌面应用（Linux / Windows），同时**不破坏现有 Web 版的 Docker 部署**。

**非目标**：

- 不重写 UI、不改变页面外观与交互流程；
- 不改造 Python 服务端，不新增服务器能力；
- 本期不含自动更新、代码签名、macOS 打包、国际化。

---

## 2. 现状关键约束（决定方案的事实）

1. **Tauri 只能加载静态资源**。生产态 WebView 通过自定义协议（`tauri://localhost` / `http://tauri.localhost`）加载 `frontendDist` 目录，不能承载 Next.js Node 服务器，因此 nextjs 必须产出**静态导出物**（`output: "export"` → `out/`）。
2. **当前 nextjs 配置与静态导出冲突**：`output: "standalone"` 与 `rewrites()`（服务端代理 `/api/*` → `backend:8000`）在 export 模式下均不可用；而 Web 版 Docker 仍依赖它们 → 必须做成**同一套代码、按构建目标切换**，不能直接改死。
3. **页面本身已是纯客户端渲染**（`page.tsx` 顶部 `"use client"`，无服务端数据获取），静态导出不存在数据预取障碍。
4. **后端逻辑极小且本质是 zip/XML 操作**：xlsx = zip，工作表保护对应 `xl/worksheets/sheetN.xml` 中的 `<sheetProtection>` 元素；openpyxl 的 `disable()` 保存时即省略该元素。移植到 Rust 成本低、可完全对齐。
5. **tauri 项目的 Vite 脚手架是多余的**：UI 全部来自 nextjs 构建物，Vite+React 模板文件（`src/`、`index.html`、`vite.config.ts` 等）应清理，tauri 项目退化为"Rust 壳 + 构建编排"。
6. 桌面端没有 HTTP 上传/下载链路：现有进度条（XHR upload progress）、Blob 锚点下载、`/api/unprotect` 代理在桌面态均需替换，但应封装在平台适配层内，UI 组件与 hooks 尽量不动。

---

## 3. 总体架构

```
┌─────────────────────────────── excel-unprotector-tauri（Rust 壳）──────────────────────────────┐
│                                                                                                │
│   Tauri WebView（WebView2 / WebKitGTK）                                                        │
│   ┌──────────────────────────────────────────────────────────────────────┐                     │
│   │  nextjs 静态导出物 out/（页面、HeroUI、Tailwind、React）              │                     │
│   │  ┌────────────────────────────────────────────────────────────────┐  │                     │
│   │  │ lib/api 平台适配层（isTauri 运行时判定）                        │  │                     │
│   │  │   Web 态  : axios → /api/unprotect（现有逻辑，原样保留）         │  │                     │
│   │  │   桌面态 : invoke("unprotect_excel") + invoke("save_file")      │  │                     │
│   │  └────────────────────────────────────────────────────────────────┘  │                     │
│   └──────────────────────────────────────────────────────────────────────┘                     │
│                    │ IPC（原始字节，见 §6）                        │ 原生保存对话框            │
│                    ▼                                                   ▼                        │
│   Rust 命令：unprotect_excel(Vec<u8>) → Vec<u8>          save_file()（tauri-plugin-dialog）     │
│              zip 重写 + quick-xml 剥离 <sheetProtection>          std::fs 写入用户选择路径     │
└────────────────────────────────────────────────────────────────────────────────────────────────┘
            ▲ 构建期：DESKTOP=1 next build → out/ → 同步到 tauri/dist
            │
┌───────────┴───────────────────┐
│ excel-unprotector-nextjs      │  同一套源码
│  BUILD_TARGET=web（默认）      │  standalone + rewrites → Docker（现有 Web 版，不受影响）
└───────────────────────────────┘
```

三条核心设计原则：

1. **一份前端源码，两个构建目标**：Web 版能力零回退，桌面版纯静态。
2. **业务逻辑原生内化**：用 Rust 复刻 openpyxl 行为，不携带 Python 运行时、不依赖网络。
3. **差异收敛在适配层**：页面组件、hooks、样式不感知平台。

---

## 4. 关键方案决策（含备选对比）

### 4.1 "解除保护"能力如何在桌面端实现

| 方案 | 描述 | 体积/资源 | 离线 | 维护成本 | 结论 |
| --- | --- | --- | --- | --- | --- |
| **A. Rust 原生命令（推荐）** | `zip` + `quick-xml` 重写 xlsx，剥离工作表保护元素 | 增量极小（数 MB），无额外运行时 | 完全离线 | 需移植约 20~50 行逻辑并做对齐测试 | **采用** |
| B. Python Sidecar | PyInstaller 打包 FastAPI+openpyxl，Tauri 拉起本地进程 | 安装包增加 60MB+，需管进程生命周期/端口/三平台二进制 | 离线 | 逻辑零移植，但分发与构建复杂度高 | 不采用 |
| C. 调用远程服务 | 桌面端直连线上 FastAPI | 最小 | 需联网，文件上传第三方，隐私风险 | 省 native 开发但与桌面端价值相悖 | 不采用 |
| D. Node Sidecar | 捆绑 Node + 用 JS zip 库处理 | 体积大，链路冗余 | 离线 | 两头不讨好 | 不采用 |

**选 A 的依据**：服务端真实逻辑仅 1 个函数、6 行有效代码，且无 Excel 重算（openpyxl 不执行公式），Rust 侧只做 zip 条目级 XML 过滤，行为可逐字节语义对齐，并用现有受保护样例文件做回归验证。

### 4.2 Next.js 构建物如何接入 Tauri

- **构建模式切换**：以环境变量 `BUILD_TARGET=desktop` 驱动 `next.config.ts` 分支：
  - desktop：`output: "export"`、`images.unoptimized = true`、**不注册** `rewrites`；
  - web（默认）：`output: "standalone"` + 现有 `rewrites`，Dockerfile 与 compose 完全不动。
- **产物落点**：静态导出产物为 `excel-unprotector-nextjs/out/`。Tauri 配置保持 `frontendDist: "../dist"`，由一个跨平台 Node 编排脚本（nextjs 项目内 `scripts/export-desktop.mjs`）完成 `next build` → 将 `out/` 全量同步到 `excel-unprotector-tauri/dist/`。
  - 不直接把 `frontendDist` 指到 `../../excel-unprotector-nextjs/out`：Tauri 对项目根目录之外的资源目录在部分打包目标上存在限制，复制进 tauri 项目最稳妥、可预期。
  - 同步采用"清空目标目录后递归复制"，避免旧产物残留。
- **开发态**：`beforeDevCommand` 直接启动 next dev 监听 1420 端口，`devUrl: http://localhost:1420`；桌面差异全部由前端运行时 `isTauri()` 判定，dev 态不需要静态导出，HMR 正常工作。

### 4.3 结果文件如何保存

- **不使用** WebView 内 `<a download>`（Linux WebKitGTK 下行为不稳定）。
- 由 Rust 命令 `save_file(default_name, bytes)` 通过 **tauri-plugin-dialog（Rust 侧 API）**弹原生"另存为"对话框，用户确认后 `std::fs::write` 落盘；取消保存返回 `null`。
- Rust 侧调用对话框不经过 WebView IPC 授权链，可避免 fs 插件 scope 配置问题；capabilities 保持最小集。
- 处理结果字节仍先回到前端形成 Blob（沿用现有结果页与"处理其他文件"流程），点击"下载文件"时才弹保存框，交互与 Web 版一致。

### 4.4 拖拽与文件选择

- Tauri v2 窗口的 `dragDropEnabled` 默认为 `true`，HTML5 拖放可直达 WebView，现有 `UploadZone` 拖拽逻辑预期可直接工作（列为 Linux/WebKitGTK 必测项，见 §10）。
- `<input type="file" accept=".xlsx">` 在 WebView 中自动调用系统文件选择器，无需 Tauri dialog 改造。

---

## 5. 详细改造清单

### 5.1 `excel-unprotector-nextjs/`

**修改**

| 文件 | 改动 |
| --- | --- |
| `next.config.ts` | 读取 `BUILD_TARGET`：desktop 分支启用 `output: "export"`、`images.unoptimized`、省略 `rewrites`；web 分支保持现状。备选开关 `trailingSlash`/`assetPrefix` 仅在自定义协议资源路径出问题时启用（见 §11 风险 R1） |
| `package.json` | 新增依赖 `@tauri-apps/api@^2`；新增脚本 `build:desktop`（调用编排脚本）；如页脚存在外链则增加 `@tauri-apps/plugin-opener@^2`（实施时确认 `Footer.tsx`） |
| `lib/api/api.ts` | 重构为平台分发：Web 实现（现有 axios 逻辑原样下沉）与 Tauri 实现（`invoke`）二选一；保持 `uploadFile(file, options)` 签名与错误文案不变 |
| `lib/hooks/useUpload.ts` | 仅两处小改：①桌面态跳过 uploading 进度，置为 `processing`；②`handleDownload` 改调适配层 `saveResult(blob, fileName)`（Web 内部仍是锚点下载）。状态机与对象 URL 生命周期管理不变 |

**新增**

| 文件 | 职责 |
| --- | --- |
| `scripts/export-desktop.mjs` | 跨平台编排：`BUILD_TARGET=desktop` 启动 `next build`；成功后清空并复制 `out/` → `../excel-unprotector-tauri/dist/`；失败非零退出 |
| `lib/api/unprotect.web.ts` | 现有 axios 上传、Content-Disposition 文件名解析、错误归一化逻辑原样迁入 |
| `lib/api/unprotect.tauri.ts` | `file.arrayBuffer()` → `invoke("unprotect_excel", { fileName, data })`（Uint8Array 原始字节）→ ArrayBuffer 转 Blob；`saveResult()` 调 `invoke("save_file", …)`；文件名生成 `unprotected_<原名>` 与服务端规则一致 |
| `lib/api/index.ts`（或在 api.ts 内） | `isTauri()` 判定并导出统一的 `processFile` / `saveResult` |

**不改**：`app/`、`components/`、`lib/types`、`lib/utils`、Dockerfile、`pnpm-lock` 之外的 Web 部署链路。

### 5.2 `excel-unprotector-tauri/`

**Rust 侧（`src-tauri/`）**

| 文件 | 改动 |
| --- | --- |
| `Cargo.toml` | 新增 `zip`（重写归档）、`quick-xml`（流式 XML 过滤）、`tauri-plugin-dialog`（保存对话框）；移除 `greet` 相关代码（依赖项保留 tauri/serde/serde_json） |
| `src/lib.rs` | 删除 `greet`；注册 dialog 插件与两个命令：`unprotect_excel`、`save_file`（详见 §6） |
| `src/excel.rs`（新增模块） | xlsx 解除保护纯函数实现 + `#[cfg(test)]` 单测 |
| `capabilities/default.json` | 维持 `core:default`；保留 `opener:default`（仅当确认存在外链需求）；不开放 fs 任意路径权限 |
| `tauri.conf.json` | 见下方配置项 |

`tauri.conf.json` 目标值（要点，非最终文本）：

- `productName`: `Excel Unprotector`；`identifier`: `app.cliff.excel-unprotector`；版本与 nextjs 对齐 `0.1.0`；
- 窗口：标题 `Excel Unprotector`，`960 × 720`，最小 `800 × 600`，可缩放；
- `build.devUrl`: `http://localhost:1420`；
- `build.beforeDevCommand`: `pnpm --dir ../excel-unprotector-nextjs dev --port 1420`；
- `build.beforeBuildCommand`: `node ../excel-unprotector-nextjs/scripts/export-desktop.mjs`；
- `build.frontendDist`: `../dist`（保持脚手架默认相对路径）；
- `app.security.csp`：首发用收紧白名单策略（仅 `self` + 必要 inline style；Next 静态导出的内联引导脚本需实测，若受阻则在发布前专项处理，不留 `null` 上线）；
- `bundle.targets`：Linux 现环境默认产物；Windows 安装包（NSIS，WebView2 引导下载）在原生 Windows / CI 上产出（见 §9）；
- 图标：暂用脚手架图标，正式发布前替换为产品图标（`tauri icon` 重新生成全套）。

**前端脚手架清理（删除）**

`src/`（App.tsx/main.tsx/App.css/assets/vite-env.d.ts）、`index.html`、`vite.config.ts`、`tsconfig.json`、`tsconfig.node.json`、`public/`（vite 示例图）。

`package.json` 瘦身为仅保留：`@tauri-apps/cli`（devDependency）与 `tauri` 脚本（`tauri dev` / `tauri build`）；移除 vite、react、模板 typescript 等依赖。React/HeroUI 依赖归属 nextjs 项目，不重复安装。

### 5.3 改造后目录结构（要点）

```
excel_unprotector/
├── excel-unprotector-nextjs/
│   ├── scripts/export-desktop.mjs      # 新增：构建+同步编排
│   ├── lib/api/
│   │   ├── api.ts                      # 平台分发
│   │   ├── unprotect.web.ts            # 新增（由现有逻辑迁入）
│   │   └── unprotect.tauri.ts          # 新增：IPC 适配
│   └── next.config.ts                  # 双目标分支
├── excel-unprotector-tauri/
│   ├── dist/                           # 新增（gitignore）：nextjs out/ 同步物
│   ├── package.json                    # 瘦身：仅保留 tauri CLI
│   └── src-tauri/
│       ├── src/lib.rs                  # 命令注册
│       ├── src/excel.rs                # 新增：解除保护逻辑+单测
│       ├── capabilities/default.json
│       └── tauri.conf.json
└── excel-unprotector-server/           # 不动；data/protected_file.xlsx 作为测试基准
```

---

## 6. Rust 核心逻辑规格（与 openpyxl 行为对齐）

### 6.1 命令接口

```rust
// 处理：输入原始字节，输出处理后的 xlsx 字节；错误映射为中文字符串
#[tauri::command]
fn unprotect_excel(file_name: String, data: Vec<u8>) -> Result<tauri::ipc::Response, String>;

// 保存：弹原生另存为对话框，落盘；取消返回 Ok(None)
#[tauri::command]
fn save_file(app: tauri::AppHandle, default_file_name: String, data: Vec<u8>)
    -> Result<Option<String>, String>;
```

- IPC 使用 Tauri v2 的**原始字节通道**：JS 侧传 `Uint8Array`，Rust 侧收 `Vec<u8>`；返回值用 `tauri::ipc::Response`，JS 侧收 `ArrayBuffer`，避免 JSON/base64 膨胀。实施时以当前 Tauri 小版本文档为准；若该版本通道不可用，回退 base64（serde）并在 §11 风险中记录。
- `save_file` 经 `tauri_plugin_dialog` 的 Rust 侧 blocking save 对话框取路径，`std::fs::write` 写入，返回实际保存路径。

### 6.2 解除保护算法

输入：xlsx 字节流。输出：新 xlsx 字节流。

1. 用 `zip::ZipArchive` 打开输入；无效 zip / 非 OOXML 包 → 返回错误（前端展示"文件处理失败"类文案）。
2. 创建 `ZipWriter` 逐条遍历：
   - 路径匹配 `xl/worksheets/sheet<N>.xml` 的条目：用 `quick-xml` 流式读写，**丢弃 `sheetProtection` 元素**（同时兼容自闭合 `<sheetProtection …/>` 与 `<sheetProtection>…</sheetProtection>` 两种形态；按本地名匹配，忽略命名空间前缀），其余 XML 事件原样写出；
   - 其他所有条目（含 `[Content_Types].xml`、rels、sharedStrings、styles、目录条目等）：**原始字节逐字复制**（zip crate 的 raw copy 能力，保留原压缩方式与条目名），不做任何重序列化。
3. 完成后 flush 返回。

**行为对齐边界（与 openpyxl 当前实现严格一致）**：

- 只处理工作表（worksheets）；**不**移除 `xl/workbook.xml` 的 `<workbookProtection>`（工作簿结构保护）；**不**处理 chartsheet；本期作为已知边界写入说明，不擅自扩大范围。
- 不校验/重算公式、不修改任何单元格内容。
- 仅接受 `.xlsx`；`.xlsm`/`.xls` 的拦截仍由前端 `validateFile` 完成（与 Web 版一致），Rust 对非 zip 输入返回明确错误即可。

### 6.3 错误文案映射

| Rust 侧情形 | 前端收到的文案（与现有体验对齐） |
| --- | --- |
| zip 打开失败 / 非 xlsx 包 | `文件处理失败：文件格式无效或已损坏` |
| XML 解析失败 | `文件处理失败：无法读取工作表内容` |
| IO/写盘失败 | `保存文件失败，请重试` |
| 用户取消保存对话框 | 非错误，停留在结果页 |

---

## 7. 前端适配层设计

- 统一接口（签名向现有 `uploadFile` 看齐）：
  - `processFile(file: File): Promise<UploadResult>`（返回 `{ blob, fileName }`）
  - `saveResult(blob: Blob, fileName: string): Promise<"saved" | "cancelled">`
- 运行时判定：`import { isTauri } from "@tauri-apps/api/core"`。该包在 Web 构建中被正常打包，`isTauri()` 返回 `false`，不产生任何 IPC 调用，Web 行为零变化。
- 桌面态 `processFile`：
  1. 先经现有 `validateFile`（hooks 层已做，不动）；
  2. `new Uint8Array(await file.arrayBuffer())` → `invoke("unprotect_excel", { fileName: file.name, data })`；
  3. 响应 `ArrayBuffer` → `new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })`；
  4. 输出文件名 `unprotected_${file.name}`（与 FastAPI 的 `Content-Disposition` 规则一致）。
- 桌面态进度：没有上传阶段，直接进入 `processing`（现有 UI 已有 Spinner 与"正在移除工作表保护，请稍候..."文案，无需新增 UI）。本地处理通常亚秒级；若大文件出现可感知耗时，后续可用 Tauri `Channel` 补进度，本期不做。
- 桌面态下载：`blob.arrayBuffer()` → `invoke("save_file", { defaultFileName, data })`；`cancelled` 时保持在结果页；`saved` 可轻量提示（复用现有结果页，不新增组件）。
- Web 态：现有 axios、进度、Content-Disposition 解析、锚点下载逻辑**逐行搬迁、不改行为**。

---

## 8. 开发与构建流程

| 场景 | 命令（在 `excel-unprotector-tauri/` 下，除注明外） |
| --- | --- |
| 桌面开发（热更新） | `pnpm tauri dev`（自动执行 `next dev --port 1420`） |
| 桌面生产包 | `pnpm tauri build`（先静态导出并同步 `dist/`，再编译 Rust + 打包） |
| 仅验证前端静态导出 | nextjs 目录：`pnpm build:desktop`，检查 `out/` 与 `../excel-unprotector-tauri/dist/` |
| Web 版回归 | nextjs 目录：`pnpm build` / `pnpm dev`；根目录 `docker compose build` 行为不变 |
| Rust 单测 | `src-tauri/` 下 `cargo test` |

环境前置（Linux）：Rust 工具链（已具备）、Node 22+/pnpm 9+（已具备）、Tauri v2 系统依赖（webkit2gtk 等，当前机器已完成首次编译，依赖齐备）。

---

## 9. 发布与打包规划

- **Linux（本期可交付）**：`tauri build` 产出 AppImage / deb（按 `bundle.targets` 收敛为需要的格式）。
- **Windows（本期准备配置，不在 Linux 上交叉编译）**：依据既往经验，Linux → Windows 交叉编译 Tauri 链路易碎且无法正常签名/验证。可选路径排序：
  1. GitHub Actions `windows-latest` runner 执行 `pnpm tauri build`（推荐，后续补 CI）；
  2. Windows 真机/虚拟机原生构建（MSVC + WebView2）。
  安装包采用 NSIS，WebView2 使用引导下载策略；代码签名证书后续再议。
- **macOS**：本期不构建、不验证（需真机处理签名/公证）。
- 版本号三处对齐：nextjs `package.json`、tauri `tauri.conf.json`、Cargo.toml。

---

## 10. 测试与验收标准

### 10.1 Rust 层（cargo test）

- 以 `excel-unprotector-server/data/protected_file.xlsx` 为夹具：处理后 zip 可正常打开；每个 `xl/worksheets/sheet*.xml` 中不再包含 `sheetProtection` 元素；其余条目数量与名称不变。
- 产出物用 openpyxl 二次加载（脚本比对）与现有 Python 服务产物做语义等价校验：两版文件均能打开、保护均解除、工作表数量/单元格内容一致。
- 异常输入：伪造 zip、非 zip 字节、不含 worksheets 目录的 zip，均返回预期错误而非 panic。
- 无保护元素的 xlsx 经过处理后内容等价（幂等不损坏）。

### 10.2 前端与集成

- Web 回归：`pnpm build`、`pnpm lint` 通过；浏览器中上传/进度/下载/错误/重置全流程与改造前一致（证明双目标切换无回退）。
- 桌面 `tauri dev` 与 `tauri build` 产物分别验证：
  1. 页面样式、HeroUI 组件、字体与 Web 版一致，无资源 404；
  2. 点击选择与**拖拽** `.xlsx` 均可用（含 Linux WebKitGTK）；
  3. 处理成功 → 结果页文件名 `unprotected_原文件名`（中文文件名正常）；
  4. "下载文件"弹原生保存框，选择路径后文件可被 Excel/LibreOffice 打开且工作表保护已解除；取消保存框停留在结果页；
  5. 非 xlsx、超 50MB、损坏文件三类错误提示与 Web 版一致；
  6. "处理其他文件"复位正常；连续处理多个文件无残留/无内存异常增长。
- 断网状态下上述全部功能可用（离线验收）。

### 10.3 完成定义

- 上述用例通过；`cargo test` 通过；两个前端构建目标均成功；无遗留模板示例代码（greet/Vite 模板）。

---

## 11. 风险与对策

| # | 风险 | 影响 | 对策 |
| --- | --- | --- | --- |
| R1 | Next 静态导出在 Tauri 自定义协议下绝对路径 `/_next/...` 资源加载异常（尤 `tauri://localhost`） | 桌面端白屏/样式丢失 | 先用默认绝对路径在三端实测；备选启用 `trailingSlash: true` + 相对 `assetPrefix`（仅 desktop 分支），以实测结果决定，不预先过度配置 |
| R2 | Tauri 小版本 IPC 原始字节通道与文档不符 | 50MB 文件传递受阻 | 实施前先写最小命令验证 Uint8Array/Response；回退 base64（接受临时内存放大，50MB 上限内可控） |
| R3 | Linux WebKitGTK 对 HTML5 拖放的接收与 Tauri 窗口拖放事件冲突 | 拖拽上传失效 | 列为最早验证项；必要时监听 Tauri `getCurrentWebview().onDragDropEvent` 兜底（仅桌面态适配层内处理，不污染 UploadZone） |
| R4 | Next 16 / Turbopack 静态导出与 HeroUI 3 的边缘兼容问题 | desktop 构建失败 | 编排脚本中 desktop 构建先于 Rust 编译，失败即终止；必要时 desktop 分支回退 webpack 构建 |
| R5 | CSP 收紧后 Next 内联引导脚本/HeroUI 内联样式被拦截 | 页面功能异常 | 按 script/style 分类实测收敛；发布前必须脱离 `csp: null` |
| R6 | 重写 zip 时极少数 OOXML 变体（异常命名、chartsheet）与 openpyxl 行为偏差 | 输出文件损坏或保护残留 | 非 worksheet 条目一律 raw 复制；仅按精确路径处理；以真实文件回归，不做"聪明"扩展 |
| R7 | 双构建目标长期漂移 | Web 版被误伤 | 适配层集中分支；CI/发布前同时跑两种构建；Docker 配置锁死不改 |
| R8 | Windows 产物无法在当前 Linux 环境验证 | 发布阻塞 | 本期 Linux 先行；Windows 走原生机/CI，发布计划单列，不承诺 Linux 交叉编译 |

---

## 12. 任务分解（实施阶段建议，本方案获批后再启动）

| 阶段 | 内容 | 检查点 |
| --- | --- | --- |
| P0 | 方案评审确认（本文档） | 决策签字：Rust 原生方案、双构建目标、Rust 侧保存框 |
| P1 | nextjs 双目标配置 + `export-desktop.mjs`；产物落 `tauri/dist` | `out/` 可由静态服务器独立运行；Web 构建回归通过 |
| P2 | 清理 Tauri Vite 脚手架；改 `tauri.conf.json` 与瘦身后 package.json | `tauri dev` 窗口中原样显示 nextjs 页面 |
| P3 | Rust `excel.rs` + `unprotect_excel` 命令 + cargo 测试（含与 Python 产物等价校验） | 单测全绿，CLI 级调用验证 |
| P4 | 前端适配层（双实现 + isTauri 分发）、`save_file` 命令与保存框联调 | §10.2 桌面端用例全部通过 |
| P5 | 窗口信息、CSP、外链 opener 核查、图标替换、`tauri build` 产物验收 | Linux 安装包可安装可运行，离线可用 |
| P6（可选） | GitHub Actions 多平台构建（Windows NSIS）；应用图标与版本规范 | Windows 产物在真机验证通过 |

---

## 13. 明确不在本期范围

- 工作簿结构保护（`<workbookProtection>`）、chartsheet 保护的移除；
- `.xlsm` / 旧版 `.xls` 支持、批量文件处理；
- 桌面端切换远程服务模式、自动更新、崩溃上报；
- i18n（仓库中 `i18n/`、`messages/` 目前为空占位）；
- macOS 打包与签名、Windows 代码签名；
- 对 Python 服务端的任何改动。
