import type { UploadResult } from "../types";

interface UploadOptions {
    onProgress?: (progress: number) => void;
    onStatusChange?: (status: "uploading" | "processing") => void;
}

export async function uploadFile(
    file: File,
    options?: UploadOptions,
): Promise<UploadResult> {
    return new Promise((resolve, reject) => {
        const formData = new FormData();
        formData.append("file", file);

        const xhr = new XMLHttpRequest();

        xhr.upload.onprogress = (e) => {
            if (e.lengthComputable) {
                const progress = (e.loaded / e.total) * 100;
                options?.onProgress?.(Math.round(progress));
            }
        };

        xhr.onloadstart = () => {
            options?.onStatusChange?.("uploading");
        };

        xhr.upload.onload = () => {
            options?.onStatusChange?.("processing");
        };

        xhr.onload = () => {
            if (xhr.status === 200) {
                const blob = xhr.response as Blob;
                const contentDisposition = xhr.getResponseHeader(
                    "Content-Disposition",
                );
                let downloadedFileName = "unprotected.xlsx";
                if (contentDisposition) {
                    const match = contentDisposition.match(/filename=(.+)/);
                    if (match) {
                        downloadedFileName = match[1];
                    }
                }
                resolve({ blob, fileName: downloadedFileName });
            } else {
                const blob = xhr.response as Blob;
                const reader = new FileReader();
                reader.onload = () => {
                    let errorMessage = "文件处理失败";
                    try {
                        const response = JSON.parse(reader.result as string);
                        errorMessage = response.detail || errorMessage;
                    } catch {}
                    reject(new Error(errorMessage));
                };
                reader.readAsText(blob);
            }
        };

        xhr.onerror = () => {
            reject(new Error("网络错误，请检查后端服务是否运行"));
        };

        xhr.ontimeout = () => {
            reject(new Error("请求超时"));
        };

        xhr.open("POST", "/unprotect");
        xhr.responseType = "blob";
        xhr.timeout = 120000;
        xhr.send(formData);
    });
}
