import { useState, useCallback, useEffect } from "react";
import type { UploadState } from "../types";
import { validateFile } from "../utils/fileUtils";
import { uploadFile } from "../utils/api";

const initialState: UploadState = {
    status: "idle",
    progress: 0,
    fileName: "",
    fileSize: 0,
    errorMessage: "",
    downloadUrl: "",
};

export function useUpload() {
    const [uploadState, setUploadState] = useState<UploadState>(initialState);

    useEffect(() => {
        return () => {
            if (uploadState.downloadUrl) {
                URL.revokeObjectURL(uploadState.downloadUrl);
            }
        };
    }, [uploadState.downloadUrl]);

    const handleFileSelect = useCallback(async (files: FileList | null) => {
        if (!files || files.length === 0) return;

        const file = files[0];
        const validationError = validateFile(file);

        if (validationError) {
            setUploadState({
                ...initialState,
                status: "error",
                errorMessage: validationError,
            });
            return;
        }

        setUploadState({
            status: "uploading",
            progress: 0,
            fileName: file.name,
            fileSize: file.size,
            errorMessage: "",
            downloadUrl: "",
        });

        try {
            const result = await uploadFile(file, {
                onProgress: (progress) => {
                    setUploadState((prev) => ({ ...prev, progress }));
                },
                onStatusChange: (status) => {
                    setUploadState((prev) => ({ ...prev, status }));
                },
            });

            const downloadUrl = URL.createObjectURL(result.blob);
            setUploadState({
                status: "completed",
                progress: 100,
                fileName: result.fileName,
                fileSize: result.blob.size,
                errorMessage: "",
                downloadUrl,
            });
        } catch (error) {
            setUploadState({
                ...initialState,
                status: "error",
                errorMessage:
                    error instanceof Error ? error.message : "文件处理失败",
            });
        }
    }, []);

    const handleDownload = useCallback(() => {
        if (!uploadState.downloadUrl) return;
        const link = document.createElement("a");
        link.href = uploadState.downloadUrl;
        link.download = uploadState.fileName;
        link.click();
    }, [uploadState.downloadUrl, uploadState.fileName]);

    const handleReset = useCallback(() => {
        if (uploadState.downloadUrl) {
            URL.revokeObjectURL(uploadState.downloadUrl);
        }
        setUploadState(initialState);
    }, [uploadState.downloadUrl]);

    return {
        uploadState,
        handleFileSelect,
        handleDownload,
        handleReset,
    };
}
