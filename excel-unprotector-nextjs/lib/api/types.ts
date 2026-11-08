export interface UploadOptions {
    onProgress?: (progress: number) => void;
    onStatusChange?: (status: "uploading" | "processing") => void;
}
