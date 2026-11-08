"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { isDesktopRuntime } from "@/lib/api/api";

const XLSX_MIME =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

interface UploadZoneProps {
    onFileSelect: (files: FileList | null) => void;
    onDropError: (message: string) => void;
}

function UploadZone({ onFileSelect, onDropError }: UploadZoneProps) {
    const [isDragging, setIsDragging] = useState(false);
    const fileInputRef = useRef<HTMLInputElement>(null);
    const zoneRef = useRef<HTMLDivElement>(null);

    const handleDragOver = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDragging(true);
    }, []);

    const handleDragLeave = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDragging(false);
    }, []);

    const handleDrop = useCallback(
        (e: React.DragEvent) => {
            e.preventDefault();
            setIsDragging(false);
            onFileSelect(e.dataTransfer.files);
        },
        [onFileSelect],
    );

    const handleClick = useCallback(() => {
        fileInputRef.current?.click();
    }, []);

    // 桌面端使用 Tauri 原生拖放事件：HTML5 drop 在原生拖放启用时不会触发
    useEffect(() => {
        if (!isDesktopRuntime()) return;

        let unlisten: (() => void) | undefined;
        let cancelled = false;

        const isInsideZone = (position: { x: number; y: number }) => {
            const rect = zoneRef.current?.getBoundingClientRect();
            if (!rect) return false;
            const scale = window.devicePixelRatio || 1;
            const x = position.x / scale;
            const y = position.y / scale;
            return (
                x >= rect.left &&
                x <= rect.right &&
                y >= rect.top &&
                y <= rect.bottom
            );
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
                    payload.paths.find((p) =>
                        p.toLowerCase().endsWith(".xlsx"),
                    ) ?? payload.paths[0];
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
                    onFileSelect(dataTransfer.files);
                } catch (error) {
                    onDropError(
                        typeof error === "string"
                            ? error
                            : "无法读取拖拽的文件，请重试",
                    );
                }
            })
            .then((fn) => {
                if (cancelled) {
                    fn();
                } else {
                    unlisten = fn;
                }
            });

        return () => {
            cancelled = true;
            unlisten?.();
        };
    }, [onFileSelect, onDropError]);

    return (
        <div
            ref={zoneRef}
            className={`upload-zone ${isDragging ? "dragover" : ""}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={handleClick}
        >
            <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx"
                onChange={(e) => onFileSelect(e.target.files)}
                className="file-input"
            />
            <div className="upload-icon">
                <svg
                    width="64"
                    height="64"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="1.5"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                >
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                    <polyline points="17 8 12 3 7 8" />
                    <line x1="12" y1="3" x2="12" y2="15" />
                </svg>
            </div>
            <h2>拖拽文件到此处</h2>
            <p>或点击选择文件</p>
            <div className="file-info">
                <span>支持格式: .xlsx</span>
                <span>最大大小: 50MB</span>
            </div>
        </div>
    );
}

export default UploadZone;