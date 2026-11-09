---
name: "tauri-desktop-file-io"
description: "Diagnose and fix Tauri v2 desktop file input/output bugs: drag-drop files not reaching an HTML5 drop zone, and a save/download button that silently does nothing while working fine on Web. Invoke when Tauri desktop file drag or save fails."
---

# Tauri 桌面端文件输入输出排障

面向 **Tauri v2 + 前端（Next.js/React）双构建** 项目：同一套 UI 在 Web 正常，但在 Tauri 桌面端拖拽或保存失效时，按本技能定位。

两个高频问题的根因、修复与验证方式如下。

---

## 问题一：桌面端拖拽文件到上传区没有响应

### 现象

- Web 端拖拽正常，桌面端 `dragover` 可能有高亮，但 `drop` 后 `e.dataTransfer.files` 为空。
- 组件里的 `onDrop` 代码看起来完全正确，改组件没有效果。

### 根因

Tauri 的 WebView 有两条互斥的拖放通道：

| 配置 | 行为 |
| --- | --- |
| `dragDropEnabled: true`（默认） | Wry 注册原生拖放处理器，OS 拖放事件被原生层接管并返回 `true`，**HTML5 `drop` 事件不会触发** |
| `dragDropEnabled: false` | 原生处理器关闭，HTML5 事件可触发，但 **Linux/WebKitGTK 从文件管理器拖入时经常不投递 `DataTransfer.files`**，Windows/macOS 也不稳定 |

结论：**不要指望 HTML5 拖放接收外部文件**。正确做法是启用原生拖放，用 Tauri 的原生拖放事件拿文件路径。

> 判断依据来自依赖源码，不要只看文档或猜测：
> - `~/.cargo/registry/src/*/wry-*/src/lib.rs` → `with_drag_drop_handler` 注释明确说明返回 `true` 会阻止 OS 默认行为
> - `~/.cargo/registry/src/*/tauri-runtime-wry-*/src/lib.rs` → 启用处理器时最终返回 `true`

### 修复

**1. `src-tauri/tauri.conf.json` 打开原生拖放**

```json
"app": {
  "windows": [{ "label": "main", "dragDropEnabled": true }]
}
```

**2. Rust 侧提供按路径读取文件的命令**

用于把原生拖放给出的绝对路径转成字节，交给前端复用原有校验/处理流程。

```rust
#[tauri::command]
fn read_dropped_file(path: String) -> Result<tauri::ipc::Response, String> {
    if !path.to_lowercase().ends_with(".xlsx") {
        return Err("只支持 .xlsx 格式的 Excel 文件".to_string());
    }
    match std::fs::read(&path) {
        Ok(bytes) => Ok(tauri::ipc::Response::new(bytes)),
        Err(_) => Err("无法读取拖拽的文件，请重试".to_string()),
    }
}
```

必须注册：`tauri::generate_handler![..., read_dropped_file]`。
返回 `tauri::ipc::Response::new(bytes)` 可直接走二进制通道，避免 `Vec<u8>` 的 JSON 数组膨胀。

**3. 前端在桌面端监听原生拖放事件**

```tsx
useEffect(() => {
    if (!isDesktopRuntime()) return;

    let unlisten: (() => void) | undefined;
    let cancelled = false;

    const isInsideZone = (position: { x: number; y: number }) => {
        const rect = zoneRef.current?.getBoundingClientRect();
        if (!rect) return false;
        // 原生事件坐标是物理像素，需除以 devicePixelRatio 才能与 rect 比较
        const scale = window.devicePixelRatio || 1;
        const x = position.x / scale;
        const y = position.y / scale;
        return x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
    };

    getCurrentWebview()
        .onDragDropEvent(async (event) => {
            const payload = event.payload;
            if (payload.type === "enter" || payload.type === "over") {
                setIsDragging(isInsideZone(payload.position));
                return;
            }
            if (payload.type === "leave") {
                setIsDragging(false);
                return;
            }
            setIsDragging(false);
            if (!isInsideZone(payload.position)) return;

            const path =
                payload.paths.find((p) => p.toLowerCase().endsWith(".xlsx")) ??
                payload.paths[0];
            if (!path) return;

            try {
                const response = await invoke<ArrayBuffer | Uint8Array>(
                    "read_dropped_file",
                    { path },
                );
                const buffer =
                    response instanceof Uint8Array
                        ? (response.buffer as ArrayBuffer)
                        : response;
                const name = path.split(/[\\/]/).pop() ?? "file.xlsx";
                const file = new File([buffer], name, { type: XLSX_MIME });
                const dataTransfer = new DataTransfer();
                dataTransfer.items.add(file);
                onFileSelect(dataTransfer.files); // 复用点击选择的同一条链路
            } catch (error) {
                onDropError(typeof error === "string" ? error : "无法读取拖拽的文件，请重试");
            }
        })
        .then((fn) => {
            if (cancelled) fn();
            else unlisten = fn;
        });

    return () => {
        cancelled = true;
        unlisten?.();
    };
}, [onFileSelect, onDropError]);
```

### 关键细节（容易踩坑）

1. **坐标单位**：`position` 是物理像素，`getBoundingClientRect()` 是 CSS 像素，必须除以 `devicePixelRatio`，否则高 DPI 下判断区域会错位。
2. **构造 `File` 再塞回 `DataTransfer`**：让桌面端和 Web 端共用同一个 `onFileSelect(FileList)` 入口，校验、状态机、保存逻辑无需分叉。
3. **TS 类型**：`new File([bytes])` 会报 `Uint8Array<ArrayBufferLike>` 不是 `BlobPart`，用 `response.buffer as ArrayBuffer` 收窄。
4. **`isDesktopRuntime()` 统一判定**：不要用 `userAgent` 或环境变量自行判断平台；Web 构建里 `@tauri-apps/api` 的调用必须被条件隔离。
5. **必须完全重启**：改了 `tauri.conf.json` 或 Rust 代码，热更新不生效，要退出并重新 `pnpm tauri dev`。
6. 确认 `src-tauri/capabilities/default.json` 含 `core:default`（原生拖放事件与 IPC 已覆盖）。

---

## 问题二：桌面端「保存文件」按钮点击无反应

### 现象

- Web 端点击能下载，桌面端点击毫无反应，也没有报错、没有弹出保存对话框。
- 按钮 `disabled` 为 `false`，`pointer-events` 正常。

### 根因一（最常见）：effect cleanup 把结果 Blob 清空了

```ts
// 错误写法
useEffect(() => {
    return () => {
        if (uploadState.downloadUrl) URL.revokeObjectURL(uploadState.downloadUrl);
        blobRef.current = null; // ← 处理完成时 downloadUrl 变化，旧 effect 清理立刻执行，把刚设置的 blob 清掉
    };
}, [uploadState.downloadUrl]);

const handleDownload = async () => {
    if (!blobRef.current) return; // ← 直接静默返回，所以「没有任何反应」
    ...
};
```

设置新的 `downloadUrl` 时会触发上一次 effect 的 cleanup，此时 `blobRef.current` 刚被赋值就被置空，`handleDownload` 在守卫处提前返回。**Web 端同样受影响，只是常被其他路径掩盖。**

修复：cleanup 只回收 Object URL，不碰 Blob；Blob 只在显式重置时清空。

```ts
useEffect(() => {
    return () => {
        if (uploadState.downloadUrl) URL.revokeObjectURL(uploadState.downloadUrl);
    };
}, [uploadState.downloadUrl]);

const handleReset = useCallback(() => {
    if (uploadState.downloadUrl) URL.revokeObjectURL(uploadState.downloadUrl);
    blobRef.current = null;
    setUploadState(initialState);
}, [uploadState.downloadUrl]);
```

### 根因二：在主线程调用阻塞式原生保存框

`blocking_save_file()` 的官方文档明确禁止在主线程调用。同步 `#[tauri::command] fn` 可能跑在主线程，导致对话框不出现或界面卡死。

修复：命令改为 `async fn`。

```rust
#[tauri::command]
async fn save_file(
    app: tauri::AppHandle,
    default_file_name: String,
    data: Vec<u8>,
) -> Result<Option<String>, String> {
    use tauri_plugin_dialog::DialogExt;

    let path = app
        .dialog()
        .file()
        .set_file_name(&default_file_name)
        .add_filter("Excel 文件", &["xlsx"])
        .blocking_save_file();

    let Some(file_path) = path else { return Ok(None) }; // 用户取消

    let path_buf = file_path
        .into_path()
        .map_err(|_| "保存文件失败，请重试".to_string())?;
    std::fs::write(&path_buf, &data).map_err(|_| "保存文件失败，请重试".to_string())?;
    Ok(Some(path_buf.to_string_lossy().into_owned()))
}
```

> 参考：`~/.cargo/registry/src/*/tauri-plugin-dialog-*/src/lib.rs` 中 `blocking_save_file` 的注释与 `async fn` 示例。

### 前端调用契约

```ts
const path = await invoke<string | null>("save_file", {
    defaultFileName: fileName, // Rust 侧 default_file_name 自动映射为 camelCase
    data,
});
return path ? "saved" : "cancelled";
```

Rust 返回 `Ok(None)` 表示用户取消，**不要**把它当成错误抛给用户。

---

## 排障顺序（照此走，避免乱改）

1. **先在组件外面找原因**：`onDrop` / `onClick` 代码看起来对，就先怀疑宿主环境与状态时序。
2. **打印状态链路**：在 `handleDownload` 首行检查 `blobRef.current` 是否为 `null`、`uploadState` 是否有值。若为 `null`，直接查 effect cleanup。
3. **核对 IPC 契约**：命令名、参数 camelCase 映射、返回类型（`Response` / `Option<String>`）三处是否一致。
4. **核对插件注册与权限**：`Builder::default().plugin(...)`、`capabilities/default.json`。
5. **读依赖源码而不是猜**：`~/.cargo/registry/src/*/{wry,tauri-runtime-wry,tauri-plugin-dialog}-*/src/lib.rs`。
6. **改配置或 Rust 后必须重启 dev**。
7. **Web 端回归**：确认改动没有破坏浏览器路径（HTML5 拖放、Blob 下载）。

## 验证清单

- [ ] `pnpm lint` 通过
- [ ] `tsc --noEmit` 通过
- [ ] `cargo test` 全绿
- [ ] `pnpm build`（Web，`output: "standalone"`）通过
- [ ] `pnpm build:desktop`（`output: "export"` + 同步 Tauri `dist/`）通过
- [ ] 桌面端拖入 `.xlsx` → 进入处理中 → 处理完成
- [ ] 桌面端点击保存 → 出现原生保存框 → 文件成功写盘
- [ ] 桌面端点保存后取消 → 不报错
- [ ] Web 端拖拽与下载仍正常

## 反模式

- ❌ 反复修改 `UploadZone` 的 `onDrop` 实现 —— 问题不在组件。
- ❌ 只在 `dragDropEnabled` 的 `true/false` 之间来回切 —— 两条路都有平台缺陷，必须走原生事件。
- ❌ 用轮询、`setTimeout` 或重试点击来掩盖「无反应」—— 无反应通常是守卫提前 return。
- ❌ 在 cleanup 中清理业务数据（Blob、文件句柄），cleanup 只应释放本次 effect 自己创建的资源。
