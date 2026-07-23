import * as esbuild from "esbuild";
import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { execSync } from "node:child_process";

const watch = process.argv.includes("--watch");
const prod = process.argv.includes("--prod");

const options = {
  entryPoints: ["src/taskpane/taskpane.ts"],
  bundle: true,
  outfile: "dist/taskpane.js",
  target: "es2019",
  sourcemap: !prod,
  minify: prod,
  logLevel: "info",
};

/**
 * Prefers GIT_COMMIT from the environment (set as a Docker build ARG, since
 * .git is excluded from the build context) and falls back to asking git
 * directly for local/non-Docker builds, where .git is available on disk.
 */
function getBuildStamp() {
  const commit =
    process.env.GIT_COMMIT ||
    (() => {
      try {
        return execSync("git rev-parse --short HEAD").toString().trim();
      } catch {
        return "unknown";
      }
    })();
  return `${commit} · built ${new Date().toISOString()}`;
}

async function copyStaticAssets() {
  await mkdir("dist/assets", { recursive: true });
  const html = await readFile("src/taskpane/taskpane.html", "utf8");
  await writeFile("dist/taskpane.html", html.replace("__BUILD_INFO__", getBuildStamp()));
  await cp("src/taskpane/taskpane.css", "dist/taskpane.css");
  await cp("assets", "dist/assets", { recursive: true });
  await cp(prod ? "manifest.prod.xml" : "manifest.xml", "dist/manifest.xml");
  console.log("Copied taskpane.html (with build stamp)/css, assets/, and manifest.xml into dist/");
}

if (watch) {
  const ctx = await esbuild.context(options);
  await ctx.watch();
  await copyStaticAssets();
  console.log("Watching for changes...");
} else {
  await esbuild.build(options);
  await copyStaticAssets();
}
