function Header() {
  return (
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
  )
}

export default Header