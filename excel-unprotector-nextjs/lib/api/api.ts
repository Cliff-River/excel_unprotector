import { isTauri } from "@tauri-apps/api/core";
import type { UploadResult } from "@/lib/types";
import type { UploadOptions } from "./types";
import * as web from "./unprotect.web";
import * as tauri from "./unprotect.tauri";

export function isDesktopRuntime(): boolean {
    return isTauri();
}

export function uploadFile(
    file: File,
    options?: UploadOptions,
): Promise<UploadResult> {
    return (isTauri() ? tauri.uploadFile : web.uploadFile)(file, options);
}

export function saveResult(
    blob: Blob,
    fileName: string,
): Promise<"saved" | "cancelled"> {
    return (isTauri() ? tauri.saveResult : web.saveResult)(blob, fileName);
}
