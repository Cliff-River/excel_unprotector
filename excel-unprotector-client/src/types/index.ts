export type UploadStatus = 'idle' | 'uploading' | 'processing' | 'completed' | 'error'

export interface UploadState {
  status: UploadStatus
  progress: number
  fileName: string
  fileSize: number
  errorMessage: string
  downloadUrl: string
}

export interface UploadResult {
  blob: Blob
  fileName: string
}