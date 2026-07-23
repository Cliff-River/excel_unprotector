import type { NextConfig } from "next";

const nextConfig: NextConfig = {
    sassOptions: {
        includePaths: ["./src"],
    },
    async rewrites() {
        return [
            {
                source: "/api/:path*",
                destination: "http://localhost:8000/:path*",
            },
        ];
    },
};

export default nextConfig;