import type { NextConfig } from "next";

const isDesktop = process.env.BUILD_TARGET === "desktop";

const nextConfig: NextConfig = {
    output: isDesktop ? "export" : "standalone",
    ...(isDesktop ? { images: { unoptimized: true } } : {}),
    sassOptions: {
        includePaths: ["./"],
    },
    ...(isDesktop
        ? {}
        : {
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
          }),
};

export default nextConfig;
