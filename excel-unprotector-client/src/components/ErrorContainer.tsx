interface ErrorContainerProps {
  errorMessage: string
  onReset: () => void
}

function ErrorContainer({ errorMessage, onReset }: ErrorContainerProps) {
  return (
    <div className="error-container">
      <div className="error-icon">
        <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
          <circle cx="12" cy="12" r="10" />
          <line x1="15" y1="9" x2="9" y2="15" />
          <line x1="9" y1="9" x2="15" y2="15" />
        </svg>
      </div>
      <h2>处理失败</h2>
      <p className="error-message">{errorMessage}</p>
      <button className="reset-button" onClick={onReset}>
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
          <path d="M3 3v5h5" />
        </svg>
        重试
      </button>
    </div>
  )
}

export default ErrorContainer