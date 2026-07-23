import type { UploadStatus } from "../types";
import { formatFileSize } from "../utils/fileUtils";

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
                <div className={`status-badge ${status}`}>
                    {getStatusText(status)}
                </div>
            </div>

            {status === "uploading" && (
                <>
                    <div className="progress-bar-wrapper">
                        <div
                            className="progress-bar"
                            style={{ width: `${progress}%` }}
                        />
                    </div>
                    <p className="progress-text">{progress}%</p>
                </>
            )}

            {status === "processing" && (
                <>
                    <div className="spinner">
                        <svg
                            width="48"
                            height="48"
                            viewBox="0 0 24 24"
                            fill="none"
                        >
                            <circle
                                className="spinner-circle"
                                cx="12"
                                cy="12"
                                r="10"
                                stroke="currentColor"
                                strokeWidth="3"
                            />
                        </svg>
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