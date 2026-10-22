import { Chip, ProgressBar, Spinner } from "@heroui/react";
import type { UploadStatus } from "@/lib/types";
import { formatFileSize } from "@/lib/utils/fileUtils";

interface ProgressContainerProps {
    status: "uploading" | "processing";
    fileName: string;
    fileSize: number;
    progress: number;
}

function ProgressContainer({
    status,
    fileName,
    fileSize,
    progress,
}: ProgressContainerProps) {
    const getStatusText = (status: UploadStatus): string => {
        switch (status) {
            case "uploading":
                return "正在上传...";
            case "processing":
                return "正在处理文件...";
            default:
                return "";
        }
    };

    return (
        <div className="progress-container">
            <div className="progress-header">
                <div className="file-info-display">
                    <svg
                        width="24"
                        height="24"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                    >
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                        <polyline points="14 2 14 8 20 8" />
                        <line x1="16" y1="13" x2="8" y2="13" />
                        <line x1="16" y1="17" x2="8" y2="17" />
                        <polyline points="10 9 9 9 8 9" />
                    </svg>
                    <div>
                        <p className="file-name">{fileName}</p>
                        <p className="file-size">{formatFileSize(fileSize)}</p>
                    </div>
                </div>
                <Chip
                    className={`rounded-[20px] px-3 py-1.5 text-[13px] font-medium ${
                        status === "uploading"
                            ? "bg-[rgba(59,130,246,0.1)] text-[#3b82f6]"
                            : "bg-[rgba(245,158,11,0.1)] text-[#f59e0b]"
                    }`}
                >
                    {getStatusText(status)}
                </Chip>
            </div>

            {status === "uploading" && (
                <>
                    <ProgressBar aria-label="上传进度" value={progress}>
                        <ProgressBar.Track className="bg-[var(--border)]">
                            <ProgressBar.Fill className="bg-[linear-gradient(90deg,var(--accent)_0%,#8b5cf6_100%)]" />
                        </ProgressBar.Track>
                    </ProgressBar>
                    <p className="progress-text">{progress}%</p>
                </>
            )}

            {status === "processing" && (
                <>
                    <div className="spinner">
                        <Spinner size="lg" className="size-12" />
                    </div>
                    <p className="processing-text">
                        正在移除工作表保护，请稍候...
                    </p>
                </>
            )}
        </div>
    );
}

export default ProgressContainer;
