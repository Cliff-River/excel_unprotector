import { useState, useCallback, useRef, useEffect } from 'react'
import './App.css'

type UploadStatus = 'idle' | 'uploading' | 'processing' | 'completed' | 'error'

interface UploadState {
  status: UploadStatus
  progress: number
  fileName: string
  fileSize: number
  errorMessage: string
  downloadUrl: string
}

const MAX_FILE_SIZE = 50 * 1024 * 1024
const ALLOWED_EXTENSIONS = ['.xlsx']

function App() {
  const [uploadState, setUploadState] = useState<UploadState>({
    status: 'idle',
    progress: 0,
    fileName: '',
    fileSize: 0,
    errorMessage: '',
    downloadUrl: '',
  })

  const [isDragging, setIsDragging] = useState(false)

  const fileInputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    return () => {
      if (uploadState.downloadUrl) {
        URL.revokeObjectURL(uploadState.downloadUrl)
      }
    }
  }, [uploadState.downloadUrl])

  const validateFile = useCallback((file: File): string | null => {
    const ext = file.name.split('.').pop()?.toLowerCase()
    if (!ext || !ALLOWED_EXTENSIONS.includes(`.${ext}`)) {
      return '只支持 .xlsx 格式的 Excel 文件'
    }
    if (file.size > MAX_FILE_SIZE) {
      return `文件大小不能超过 ${MAX_FILE_SIZE / 1024 / 1024}MB`
    }
    return null
  }, [])

  const formatFileSize = useCallback((bytes: number): string => {
    if (bytes < 1024) return `${bytes} B`
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(2)} KB`
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`
  }, [])

  const handleFileSelect = useCallback((files: FileList | null) => {
    if (!files || files.length === 0) return

    const file = files[0]
    const validationError = validateFile(file)

    if (validationError) {
      setUploadState({
        ...uploadState,
        status: 'error',
        errorMessage: validationError,
      })
      return
    }

    setUploadState({
      status: 'uploading',
      progress: 0,
      fileName: file.name,
      fileSize: file.size,
      errorMessage: '',
      downloadUrl: '',
    })

    const formData = new FormData()
    formData.append('file', file)

    const xhr = new XMLHttpRequest()

    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable) {
        const progress = (e.loaded / e.total) * 100
        setUploadState((prev) => ({ ...prev, progress: Math.round(progress) }))
      }
    }

    xhr.onloadstart = () => {
      setUploadState((prev) => ({ ...prev, status: 'uploading' }))
    }

    xhr.upload.onload = () => {
      setUploadState((prev) => ({ ...prev, status: 'processing' }))
    }

    xhr.onload = () => {
      if (xhr.status === 200) {
        const blob = xhr.response
        const downloadUrl = URL.createObjectURL(blob)
        const contentDisposition = xhr.getResponseHeader('Content-Disposition')
        let downloadedFileName = 'unprotected.xlsx'
        if (contentDisposition) {
          const match = contentDisposition.match(/filename=(.+)/)
          if (match) {
            downloadedFileName = match[1]
          }
        }

        setUploadState({
          status: 'completed',
          progress: 100,
          fileName: downloadedFileName,
          fileSize: blob.size,
          errorMessage: '',
          downloadUrl,
        })
      } else {
        const blob = xhr.response as Blob
        const reader = new FileReader()
        reader.onload = () => {
          let errorMessage = '文件处理失败'
          try {
            const response = JSON.parse(reader.result as string)
            errorMessage = response.detail || errorMessage
          } catch {
          }
          setUploadState({
            status: 'error',
            progress: 0,
            fileName: '',
            fileSize: 0,
            errorMessage,
            downloadUrl: '',
          })
        }
        reader.readAsText(blob)
      }
    }

    xhr.onerror = () => {
      setUploadState({
        status: 'error',
        progress: 0,
        fileName: '',
        fileSize: 0,
        errorMessage: '网络错误，请检查后端服务是否运行',
        downloadUrl: '',
      })
    }

    xhr.ontimeout = () => {
      setUploadState({
        status: 'error',
        progress: 0,
        fileName: '',
        fileSize: 0,
        errorMessage: '请求超时',
        downloadUrl: '',
      })
    }

    xhr.open('POST', '/unprotect')
    xhr.responseType = 'blob'
    xhr.timeout = 120000
    xhr.send(formData)
  }, [validateFile, uploadState])

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(true)
  }, [])

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
  }, [])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
    handleFileSelect(e.dataTransfer.files)
  }, [handleFileSelect])

  const handleClick = useCallback(() => {
    fileInputRef.current?.click()
  }, [])

  const handleDownload = useCallback(() => {
    if (!uploadState.downloadUrl) return
    const link = document.createElement('a')
    link.href = uploadState.downloadUrl
    link.download = uploadState.fileName
    link.click()
  }, [uploadState.downloadUrl, uploadState.fileName])

  const handleReset = useCallback(() => {
    if (uploadState.downloadUrl) {
      URL.revokeObjectURL(uploadState.downloadUrl)
    }
    setUploadState({
      status: 'idle',
      progress: 0,
      fileName: '',
      fileSize: 0,
      errorMessage: '',
      downloadUrl: '',
    })
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }, [uploadState.downloadUrl])

  const getStatusText = () => {
    switch (uploadState.status) {
      case 'uploading':
        return '正在上传...'
      case 'processing':
        return '正在处理文件...'
      case 'completed':
        return '处理完成'
      case 'error':
        return '处理失败'
      default:
        return ''
    }
  }

  return (
    <div className="app-container">
      <header className="header">
        <div className="header-content">
          <div className="logo">
            <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M3 12h18" />
              <path d="M19 8v8" />
              <path d="M5 8v8" />
              <path d="M12 3v18" />
              <path d="M8 3v18" />
              <path d="M16 3v18" />
            </svg>
          </div>
          <div className="title-section">
            <h1>Excel Unprotector</h1>
            <p>快速解除 Excel 文件工作表保护</p>
          </div>
        </div>
      </header>

      <main className="main-content">
        <div className="upload-container">
          {uploadState.status === 'idle' && (
            <div
              className={`upload-zone ${isDragging ? 'dragover' : ''}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              onClick={handleClick}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx"
                onChange={(e) => handleFileSelect(e.target.files)}
                className="file-input"
              />
              <div className="upload-icon">
                <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <polyline points="17 8 12 3 7 8" />
                  <line x1="12" y1="3" x2="12" y2="15" />
                </svg>
              </div>
              <h2>拖拽文件到此处</h2>
              <p>或点击选择文件</p>
              <div className="file-info">
                <span>支持格式: .xlsx</span>
                <span>最大大小: 50MB</span>
              </div>
            </div>
          )}

          {uploadState.status === 'uploading' && (
            <div className="progress-container">
              <div className="progress-header">
                <div className="file-info-display">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                    <polyline points="14 2 14 8 20 8" />
                    <line x1="16" y1="13" x2="8" y2="13" />
                    <line x1="16" y1="17" x2="8" y2="17" />
                    <polyline points="10 9 9 9 8 9" />
                  </svg>
                  <div>
                    <p className="file-name">{uploadState.fileName}</p>
                    <p className="file-size">{formatFileSize(uploadState.fileSize)}</p>
                  </div>
                </div>
                <div className="status-badge uploading">{getStatusText()}</div>
              </div>
              <div className="progress-bar-wrapper">
                <div className="progress-bar" style={{ width: `${uploadState.progress}%` }} />
              </div>
              <p className="progress-text">{uploadState.progress}%</p>
            </div>
          )}

          {uploadState.status === 'processing' && (
            <div className="progress-container">
              <div className="progress-header">
                <div className="file-info-display">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                    <polyline points="14 2 14 8 20 8" />
                    <line x1="16" y1="13" x2="8" y2="13" />
                    <line x1="16" y1="17" x2="8" y2="17" />
                    <polyline points="10 9 9 9 8 9" />
                  </svg>
                  <div>
                    <p className="file-name">{uploadState.fileName}</p>
                    <p className="file-size">{formatFileSize(uploadState.fileSize)}</p>
                  </div>
                </div>
                <div className="status-badge processing">{getStatusText()}</div>
              </div>
              <div className="spinner">
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
                  <circle className="spinner-circle" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" />
                </svg>
              </div>
              <p className="processing-text">正在移除工作表保护，请稍候...</p>
            </div>
          )}

          {uploadState.status === 'completed' && (
            <div className="result-container">
              <div className="success-icon">
                <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="20 6 9 17 4 12" />
                </svg>
              </div>
              <h2>{getStatusText()}</h2>
              <div className="result-file-info">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="16" y1="13" x2="8" y2="13" />
                  <line x1="16" y1="17" x2="8" y2="17" />
                  <polyline points="10 9 9 9 8 9" />
                </svg>
                <div>
                  <p className="file-name">{uploadState.fileName}</p>
                  <p className="file-size">{formatFileSize(uploadState.fileSize)}</p>
                </div>
              </div>
              <div className="action-buttons">
                <button className="download-button" onClick={handleDownload}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                    <polyline points="7 10 12 15 17 10" />
                    <line x1="12" y1="15" x2="12" y2="3" />
                  </svg>
                  下载文件
                </button>
                <button className="reset-button" onClick={handleReset}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                    <path d="M3 3v5h5" />
                  </svg>
                  处理其他文件
                </button>
              </div>
            </div>
          )}

          {uploadState.status === 'error' && (
            <div className="error-container">
              <div className="error-icon">
                <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="10" />
                  <line x1="15" y1="9" x2="9" y2="15" />
                  <line x1="9" y1="9" x2="15" y2="15" />
                </svg>
              </div>
              <h2>{getStatusText()}</h2>
              <p className="error-message">{uploadState.errorMessage}</p>
              <button className="reset-button" onClick={handleReset}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                  <path d="M3 3v5h5" />
                </svg>
                重试
              </button>
            </div>
          )}
        </div>

        <section className="features">
          <div className="feature-card">
            <div className="feature-icon">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                <polyline points="17 8 12 3 7 8" />
                <line x1="12" y1="3" x2="12" y2="15" />
              </svg>
            </div>
            <h3>简单上传</h3>
            <p>支持拖拽上传，一键选择文件，操作简单直观</p>
          </div>
          <div className="feature-card">
            <div className="feature-icon">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
              </svg>
            </div>
            <h3>安全处理</h3>
            <p>文件在服务器端处理，本地不保存任何数据</p>
          </div>
          <div className="feature-card">
            <div className="feature-icon">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <circle cx="12" cy="12" r="10" />
                <polyline points="12 6 12 12 16 14" />
              </svg>
            </div>
            <h3>快速高效</h3>
            <p>批量移除工作表保护，处理速度快，省时省力</p>
          </div>
        </section>
      </main>

      <footer className="footer">
        <p>Excel Unprotector - 解除 Excel 文件工作表保护工具</p>
      </footer>
    </div>
  )
}

export default App
