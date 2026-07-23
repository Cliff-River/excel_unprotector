import axios from "axios";
import type { UploadResult } from "../types";

const API_BASE_URL = process.env.NEXT_PUBLIC_API_BASE_URL || "/api/";

interface UploadOptions {
    onProgress?: (progress: number) => void;
    onStatusChange?: (status: "uploading" | "processing") => void;
}

export async function uploadFile(
    file: File,
    options?: UploadOptions,
): Promise<UploadResult> {
    const formData = new FormData();
    formData.append("file", file);

    try {
        const response = await axios.post<Blob>(`${API_BASE_URL}unprotect`, formData, {
            responseType: "blob",
            timeout: 120000,
            onUploadProgress: (e) => {
                if (e.lengthComputable && e.total) {
                    const progress = (e.loaded / e.total) * 100;
                    options?.onProgress?.(Math.round(progress));
                    if (e.loaded === e.total) {
                        options?.onStatusChange?.("processing");
                    }
                }
            },
        });

        const contentDisposition = response.headers["content-disposition"];
        let downloadedFileName = "unprotected.xlsx";
        if (contentDisposition) {
            const match = contentDisposition.match(/filename=(.+)/);
            if (match) {
                downloadedFileName = match[1];
            }
        }

        return { blob: response.data, fileName: downloadedFileName };
    } catch (error) {
        if (axios.isAxiosError(error) && error.response && error.response.data instanceof Blob) {
            const reader = new FileReader();
            const { data } = error.response;
            return new Promise((_, reject) => {
                reader.onload = () => {
                    let errorMessage = "文件处理失败";
                    try {
                        const response = JSON.parse(reader.result as string);
                        errorMessage = response.detail || errorMessage;
                    } catch {}
                    reject(new Error(errorMessage));
                };
                reader.onerror = () => {
                    reject(new Error("文件处理失败"));
                };
                reader.readAsText(data);
            });
        }
        if (axios.isAxiosError(error)) {
            if (error.code === "ERR_NETWORK") {
                throw new Error("网络错误，请检查后端服务是否运行");
            }
            if (error.code === "ECONNABORTED") {
                throw new Error("请求超时");
            }
        }
        throw new Error("文件处理失败");
    }
}