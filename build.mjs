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
function getCommit() {
  return (
    process.env.GIT_COMMIT ||
    (() => {
      try {
        return execSync("git rev-parse --short HEAD").toString().trim();
      } catch {
        return "unknown";
      }
    })()
  );
}

async function copyStaticAssets() {
  await mkdir("dist/assets", { recursive: true });
  const commit = getCommit();
  const buildStamp = `${commit} · built ${new Date().toISOString()}`;
  const html = await readFile("src/taskpane/taskpane.html", "utf8");
  await writeFile(
    "dist/taskpane.html",
    html.replace("__BUILD_INFO__", buildStamp).replaceAll("__CACHE_BUST__", commit)
  );
  await cp("src/taskpane/taskpane.css", "dist/taskpane.css");
  await cp("assets", "dist/assets", { recursive: true });
  await cp(prod ? "manifest.prod.xml" : "manifest.xml", "dist/manifest.xml");
  // Vendored fresh from node_modules on every build (not committed to the
  // repo) so it always matches package.json's pinned @azure/msal-browser
  // version — used by /admin's standalone browser sign-in page, which
  // deliberately isn't part of the esbuild bundle (see server.js).
  // msal-redirect-bridge is a separate sub-package (not re-exported by
  // msal-browser itself) that must run on the page the popup lands on:
  // loginPopup()'s opener waits on a BroadcastChannel for the response,
  // and this bridge script is what actually reads the redirect URL's auth
  // payload and posts it there — without it the opener waits forever.
  await mkdir("dist/vendor", { recursive: true });
  await cp("node_modules/@azure/msal-browser/lib/msal-browser.min.js", "dist/vendor/msal-browser.min.js");
  await cp(
    "node_modules/@azure/msal-browser/lib/redirect-bridge/msal-redirect-bridge.min.js",
    "dist/vendor/msal-redirect-bridge.min.js"
  );
  console.log("Copied taskpane.html (with build stamp)/css, assets/, vendor/, and manifest.xml into dist/");
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
