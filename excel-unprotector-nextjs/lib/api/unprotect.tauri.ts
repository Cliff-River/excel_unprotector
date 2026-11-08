import { invoke } from "@tauri-apps/api/core";
import type { UploadResult } from "@/lib/types";
import type { UploadOptions } from "./types";

const XLSX_MIME =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export async function uploadFile(
    file: File,
    options?: UploadOptions,
): Promise<UploadResult> {
    void options;
    const data = new Uint8Array(await file.arrayBuffer());
    try {
        const response = await invoke<ArrayBuffer | Uint8Array>(
            "unprotect_excel",
            { fileName: file.name, data },
        );
        const buffer =
            response instanceof Uint8Array
                ? (response.buffer as ArrayBuffer)
                : response;
        return {
            blob: new Blob([buffer], { type: XLSX_MIME }),
            fileName: `unprotected_${file.name}`,
        };
    } catch (error) {
        if (typeof error === "string") {
            throw new Error(error);
        }
        if (error instanceof Error) {
            throw error;
        }
        throw new Error("文件处理失败");
    }
}

export async function saveResult(
    blob: Blob,
    fileName: string,
): Promise<"saved" | "cancelled"> {
    const data = new Uint8Array(await blob.arrayBuffer());
    try {
        const path = await invoke<string | null>("save_file", {
            defaultFileName: fileName,
            data,
        });
        return path ? "saved" : "cancelled";
    } catch {
        throw new Error("保存文件失败，请重试");
    }
}
