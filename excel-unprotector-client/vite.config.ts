import { defineConfig, loadEnv } from "vite";
import react, { reactCompilerPreset } from "@vitejs/plugin-react";
import babel from "@rolldown/plugin-babel";

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, process.cwd(), "");
    const backendHost = env.BACKEND_HOST || "localhost";
    const backendPort = env.BACKEND_PORT || "8000";
    const backendUrl = `http://${backendHost}:${backendPort}`;

    return {
        plugins: [react(), babel({ presets: [reactCompilerPreset()] })],
        server: {
            proxy: {
                "/unprotect": {
                    target: backendUrl,
                    changeOrigin: true,
                },
                "/health": {
                    target: backendUrl,
                    changeOrigin: true,
                },
            },
        },
    };
});
