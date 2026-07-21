import * as esbuild from "esbuild";
import { cp, mkdir } from "node:fs/promises";
import { existsSync } from "node:fs";

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

async function copyStaticAssets() {
  await mkdir("dist/assets", { recursive: true });
  await cp("src/taskpane/taskpane.html", "dist/taskpane.html");
  await cp("src/taskpane/taskpane.css", "dist/taskpane.css");
  await cp("assets", "dist/assets", { recursive: true });
  await cp(prod ? "manifest.prod.xml" : "manifest.xml", "dist/manifest.xml");
  console.log("Copied taskpane.html/css, assets/, and manifest.xml into dist/");
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
