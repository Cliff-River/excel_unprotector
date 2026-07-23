import type { NextConfig } from "next";

const nextConfig: NextConfig = {
    output: "standalone",
    sassOptions: {
        includePaths: ["./src"],
    },
    async rewrites() {
        const isProduction = process.env.NODE_ENV === "production";
        return [
            {
                source: "/api/:path*",
                destination: isProduction
                    ? "http://backend:8000/:path*"
                    : "http://localhost:8000/:path*",
            },
        ];
    },
};

export default nextConfig;