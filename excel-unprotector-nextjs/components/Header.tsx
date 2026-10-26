import { Link } from "@heroui/react";

const GITHUB_URL = "https://github.com/Cliff-River/excel_unprotector";

function Header() {
    return (
        <header className="header">
            <div className="header-content">
                <div className="logo">
                    <svg
                        width="36"
                        height="36"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                    >
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
                <Link
                    href={GITHUB_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    aria-label="在 GitHub 上查看项目"
                    className="ml-auto flex size-10 shrink-0 items-center justify-center rounded-xl bg-white/15 text-white no-underline transition-colors hover:bg-white/25 hover:text-white"
                >
                    <svg
                        className="size-6"
                        viewBox="0 0 24 24"
                        fill="currentColor"
                        aria-hidden="true"
                    >
                        <path d="M12 .5C5.73.5.5 5.73.5 12c0 5.3 3.44 9.8 8.21 11.39.6.11.82-.26.82-.58 0-.29-.01-1.04-.02-2.05-3.34.73-4.04-1.61-4.04-1.61-.55-1.39-1.34-1.76-1.34-1.76-1.09-.75.08-.73.08-.73 1.2.08 1.84 1.24 1.84 1.24 1.07 1.84 2.81 1.31 3.5 1 .11-.78.42-1.31.76-1.61-2.67-.3-5.47-1.33-5.47-5.93 0-1.31.47-2.38 1.24-3.22-.12-.3-.54-1.52.12-3.17 0 0 1.01-.32 3.3 1.23a11.5 11.5 0 0 1 6 0c2.29-1.55 3.3-1.23 3.3-1.23.66 1.65.24 2.87.12 3.17.77.84 1.24 1.91 1.24 3.22 0 4.61-2.81 5.63-5.48 5.92.43.37.81 1.1.81 2.22 0 1.61-.01 2.9-.01 3.29 0 .32.22.7.83.58A12.01 12.01 0 0 0 23.5 12C23.5 5.73 18.27.5 12 .5Z" />
                    </svg>
                </Link>
            </div>
        </header>
    );
}

export default Header;