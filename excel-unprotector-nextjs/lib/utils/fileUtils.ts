const MAX_FILE_SIZE = 50 * 1024 * 1024;
const ALLOWED_EXTENSIONS = [".xlsx"];

export function validateFile(file: File): string | null {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (!ext || !ALLOWED_EXTENSIONS.includes(`.${ext}`)) {
        return "只支持 .xlsx 格式的 Excel 文件";
    }
    if (file.size > MAX_FILE_SIZE) {
        return `文件大小不能超过 ${MAX_FILE_SIZE / 1024 / 1024}MB`;
    }
    return null;
}

export function formatFileSize(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(2)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}