import * as esbuild from "esbuild";
import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { execSync } from "node:child_process";

const watch = process.argv.includes("--watch");
const prod = process.argv.includes("--prod");

const sharedOptions = {
  bundle: true,
  target: "es2019",
  sourcemap: !prod,
  minify: prod,
  logLevel: "info",
};
const options = { ...sharedOptions, entryPoints: ["src/taskpane/taskpane.ts"], outfile: "dist/taskpane.js" };
// Phase 5: the gallery dialog. A separate bundle, not a second entry in
// the same esbuild.build() call — esbuild rejects a single `outfile` with
// multiple entry points, and keeping two independent build() calls is
// simpler than switching to `outdir` (which would rename taskpane.js's
// output path and touch every reference to it) for one extra small page.
const galleryOptions = { ...sharedOptions, entryPoints: ["src/gallery/gallery.ts"], outfile: "dist/gallery.js" };
// Task Pane Phase 15: the account-creation dialog, same reasoning as
// galleryOptions above — its own small bundle, not a second entry point
// sharing taskpane.js's outfile.
const registerOptions = { ...sharedOptions, entryPoints: ["src/register/register.ts"], outfile: "dist/register.js" };
// Task Pane Phase 20: the template library dialog, same reasoning as
// galleryOptions/registerOptions above.
const templatesOptions = { ...sharedOptions, entryPoints: ["src/templates/templates.ts"], outfile: "dist/templates.js" };
// Library Upload: the bulk-import review dialog, same reasoning as the
// entry points above — its own small bundle, not a second entry point
// sharing taskpane.js's outfile.
const libraryUploadOptions = {
  ...sharedOptions,
  entryPoints: ["src/library-upload/library-upload.ts"],
  outfile: "dist/library-upload.js",
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
  const galleryHtml = await readFile("src/gallery/gallery.html", "utf8");
  await writeFile("dist/gallery.html", galleryHtml.replaceAll("__CACHE_BUST__", commit));
  await cp("src/gallery/gallery.css", "dist/gallery.css");
  const registerHtml = await readFile("src/register/register.html", "utf8");
  await writeFile("dist/register.html", registerHtml.replaceAll("__CACHE_BUST__", commit));
  await cp("src/register/register.css", "dist/register.css");
  const templatesHtml = await readFile("src/templates/templates.html", "utf8");
  await writeFile("dist/templates.html", templatesHtml.replaceAll("__CACHE_BUST__", commit));
  await cp("src/templates/templates.css", "dist/templates.css");
  const libraryUploadHtml = await readFile("src/library-upload/library-upload.html", "utf8");
  await writeFile("dist/library-upload.html", libraryUploadHtml.replaceAll("__CACHE_BUST__", commit));
  await cp("src/library-upload/library-upload.css", "dist/library-upload.css");
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
  console.log(
    "Copied taskpane.html (with build stamp)/css, gallery/register/templates html+css, assets/, vendor/, and manifest.xml into dist/"
  );
}

if (watch) {
  const ctx = await esbuild.context(options);
  const galleryCtx = await esbuild.context(galleryOptions);
  const registerCtx = await esbuild.context(registerOptions);
  const templatesCtx = await esbuild.context(templatesOptions);
  const libraryUploadCtx = await esbuild.context(libraryUploadOptions);
  await ctx.watch();
  await galleryCtx.watch();
  await registerCtx.watch();
  await templatesCtx.watch();
  await libraryUploadCtx.watch();
  await copyStaticAssets();
  console.log("Watching for changes...");
} else {
  await esbuild.build(options);
  await esbuild.build(galleryOptions);
  await esbuild.build(registerOptions);
  await esbuild.build(templatesOptions);
  await esbuild.build(libraryUploadOptions);
  await copyStaticAssets();
}
