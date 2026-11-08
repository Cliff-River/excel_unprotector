"use client";

import { useState, useCallback, useEffect, useRef } from "react";
import type { UploadState } from "@/lib/types";
import { validateFile } from "@/lib/utils/fileUtils";
import { isDesktopRuntime, saveResult, uploadFile } from "@/lib/api/api";

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
    const blobRef = useRef<Blob | null>(null);

    useEffect(() => {
        return () => {
            if (uploadState.downloadUrl) {
                URL.revokeObjectURL(uploadState.downloadUrl);
            }
            blobRef.current = null;
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

        const desktop = isDesktopRuntime();

        setUploadState({
            status: desktop ? "processing" : "uploading",
            progress: 0,
            fileName: file.name,
            fileSize: file.size,
            errorMessage: "",
            downloadUrl: "",
        });

        try {
            const result = await uploadFile(file, {
                onProgress: (progress) => {
                    if (desktop) return;
                    setUploadState((prev) => ({ ...prev, progress }));
                },
                onStatusChange: (status) => {
                    setUploadState((prev) => ({ ...prev, status }));
                },
            });

            blobRef.current = result.blob;
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
            blobRef.current = null;
            setUploadState({
                ...initialState,
                status: "error",
                errorMessage:
                    error instanceof Error ? error.message : "文件处理失败",
            });
        }
    }, []);

    const handleDownload = useCallback(async () => {
        if (!blobRef.current) return;
        try {
            await saveResult(blobRef.current, uploadState.fileName);
        } catch (error) {
            setUploadState({
                ...initialState,
                status: "error",
                errorMessage:
                    error instanceof Error ? error.message : "保存文件失败，请重试",
            });
        }
    }, [uploadState.fileName]);

    const handleReset = useCallback(() => {
        if (uploadState.downloadUrl) {
            URL.revokeObjectURL(uploadState.downloadUrl);
        }
        blobRef.current = null;
        setUploadState(initialState);
    }, [uploadState.downloadUrl]);

    return {
        uploadState,
        handleFileSelect,
        handleDownload,
        handleReset,
    };
}
