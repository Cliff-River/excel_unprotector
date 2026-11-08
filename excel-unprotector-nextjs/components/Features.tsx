"use client";

import { useSyncExternalStore } from "react";
import FeatureCard from "./FeatureCard";
import { isDesktopRuntime } from "@/lib/api/api";

const subscribe = () => () => {};

function Features() {
    // 服务端快照固定为 Web；客户端挂载后再按真实运行环境切换，避免水合不一致
    const isDesktop = useSyncExternalStore(
        subscribe,
        isDesktopRuntime,
        () => false,
    );

    return (
        <section className="features">
            <FeatureCard
                title={isDesktop ? "本地导入" : "简单上传"}
                description={
                    isDesktop
                        ? "支持拖拽或点击选择文件，全程无需联网"
                        : "支持拖拽上传，一键选择文件，操作简单直观"
                }
                icon={
                    <svg
                        width="40"
                        height="40"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                    >
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                        <polyline points="17 8 12 3 7 8" />
                        <line x1="12" y1="3" x2="12" y2="15" />
                    </svg>
                }
            />
            <FeatureCard
                title={isDesktop ? "本地处理" : "安全处理"}
                description={
                    isDesktop
                        ? "文件只在本机处理，不上传网络，也不保存任何数据"
                        : "文件在服务器端处理，本地不保存任何数据"
                }
                icon={
                    <svg
                        width="40"
                        height="40"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                    >
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                    </svg>
                }
            />
            <FeatureCard
                title="快速高效"
                description="免密解除，批量移除工作表保护，处理速度快，省时省力"
                icon={
                    <svg
                        width="40"
                        height="40"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                    >
                        <circle cx="12" cy="12" r="10" />
                        <polyline points="12 6 12 12 16 14" />
                    </svg>
                }
            />
        </section>
    );
}

export default Features;