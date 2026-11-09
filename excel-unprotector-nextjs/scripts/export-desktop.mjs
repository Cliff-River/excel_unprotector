import { spawnSync } from "node:child_process";
import {
    cpSync,
    existsSync,
    mkdirSync,
    readdirSync,
    rmSync,
} from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const nextjsRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outDir = path.join(nextjsRoot, "out");
const tauriDist = path.resolve(nextjsRoot, "../excel-unprotector-tauri/dist");

const pnpmCmd = process.platform === "win32" ? "pnpm.cmd" : "pnpm";

console.log("[export-desktop] Building Next.js static export (BUILD_TARGET=desktop)...");
const build = spawnSync(pnpmCmd, ["exec", "next", "build"], {
    cwd: nextjsRoot,
    stdio: "inherit",
    // Windows 上 spawnSync 无法直接启动 .cmd，需经 shell 转发
    shell: process.platform === "win32",
    env: { ...process.env, BUILD_TARGET: "desktop" },
});

if (build.status !== 0) {
    console.error("[export-desktop] next build failed");
    process.exit(build.status ?? 1);
}

if (!existsSync(path.join(outDir, "index.html"))) {
    console.error(`[export-desktop] missing ${outDir}/index.html — export did not produce expected output`);
    process.exit(1);
}

rmSync(tauriDist, { recursive: true, force: true });
mkdirSync(tauriDist, { recursive: true });
cpSync(outDir, tauriDist, { recursive: true });

const count = readdirSync(tauriDist).length;
console.log(`[export-desktop] synced ${count} top-level entries -> ${tauriDist}`);
