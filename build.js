const esbuild = require("esbuild");

const watch = process.argv.includes("--watch");

const options = {
  entryPoints: ["src/app.jsx"],
  bundle: true,
  minify: true,
  target: ["ios14", "chrome100"],
  outfile: "app.js",
  loader: { ".jsx": "jsx" },
  jsx: "automatic",
  logLevel: "info",
};

(async () => {
  if (watch) {
    const ctx = await esbuild.context(options);
    await ctx.watch();
    console.log("Watching for changes...");
  } else {
    await esbuild.build(options);
  }
})().catch(() => process.exit(1));
