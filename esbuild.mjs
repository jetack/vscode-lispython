import fs from "node:fs";
import esbuild from "esbuild";

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");
const outfileArg = process.argv.find((a) => a.startsWith("--outfile="));

/** @type {import("esbuild").BuildOptions} */
const options = {
  entryPoints: ["src/extension.ts"],
  outfile: outfileArg ? outfileArg.slice("--outfile=".length) : "out/extension.js",
  bundle: true,
  platform: "node",
  format: "cjs",
  external: ["vscode"],
  // VS Code 1.75 engine floor ships Electron 19 (Node 16.14).
  target: "node16",
  minify: production,
  sourcemap: !production,
  logLevel: "info",
};

// A production bundle must never sit beside a stale dev .map from an earlier build.
if (production) fs.rmSync(`${options.outfile}.map`, { force: true });

try {
  if (watch) {
    const ctx = await esbuild.context(options);
    await ctx.watch();
  } else {
    await esbuild.build(options);
  }
} catch {
  process.exit(1);
}
