import esbuild from "esbuild";

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");

/** @type {import("esbuild").BuildOptions} */
const options = {
  entryPoints: ["src/extension.ts"],
  outfile: "out/extension.js",
  bundle: true,
  platform: "node",
  format: "cjs",
  external: ["vscode"],
  // VS Code 1.75 engine floor = Electron 22 = Node 16.
  target: "node16",
  minify: production,
  sourcemap: !production,
  logLevel: "info",
};

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
