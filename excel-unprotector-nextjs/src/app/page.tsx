"use client";

import Header from "../components/Header";
import UploadZone from "../components/UploadZone";
import ProgressContainer from "../components/ProgressContainer";
import ResultContainer from "../components/ResultContainer";
import ErrorContainer from "../components/ErrorContainer";
import Features from "../components/Features";
import Footer from "../components/Footer";
import { useUpload } from "../hooks/useUpload";

export default function HomePage() {
    const { uploadState, handleFileSelect, handleDownload, handleReset } =
        useUpload();

    return (
        <div className="app-container">
            <Header />

            <main className="main-content">
                <div className="upload-container">
                    {uploadState.status === "idle" && (
                        <UploadZone onFileSelect={handleFileSelect} />
                    )}

                    {(uploadState.status === "uploading" ||
                        uploadState.status === "processing") && (
                        <ProgressContainer
                            status={uploadState.status}
                            fileName={uploadState.fileName}
                            fileSize={uploadState.fileSize}
                            progress={uploadState.progress}
                        />
                    )}

                    {uploadState.status === "completed" && (
                        <ResultContainer
                            fileName={uploadState.fileName}
                            fileSize={uploadState.fileSize}
                            onDownload={handleDownload}
                            onReset={handleReset}
                        />
                    )}

                    {uploadState.status === "error" && (
                        <ErrorContainer
                            errorMessage={uploadState.errorMessage}
                            onReset={handleReset}
                        />
                    )}
                </div>

                <Features />
            </main>

            <Footer />
        </div>
    );
}