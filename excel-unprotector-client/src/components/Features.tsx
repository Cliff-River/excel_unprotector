function Features() {
  return (
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
  )
}

export default Features