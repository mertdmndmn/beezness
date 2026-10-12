const esbuild = require("esbuild");
const fs = require("fs");
const crypto = require("crypto");

// Every build gets its own version, stamped where browsers look for it:
// index.html loads app.js?v=<hash>, so a normal refresh can never pair a
// fresh page with a cached old app.js; sw.js's cache name changes, which is
// what the page's foreground check (and the browser) use to spot a deploy.
// Numeric on purpose: older pages match /beezness-v\d+/.
function stamp() {
  const hash = crypto.createHash("sha1").update(fs.readFileSync("app.js")).digest("hex");
  const version = parseInt(hash.slice(0, 8), 16);
  for (const file of ["index.html"]) {
    const html = fs.readFileSync(file, "utf8");
    fs.writeFileSync(file, html.replace(/<script src="app\.js(\?v=\w+)?"><\/script>/, `<script src="app.js?v=${hash.slice(0, 10)}"></script>`));
  }
  const sw = fs.readFileSync("sw.js", "utf8");
  fs.writeFileSync("sw.js", sw.replace(/const CACHE = "beezness-v\d+";/, `const CACHE = "beezness-v${version}";`));
}

// The Excel library (~1 MB) is only needed to download a Z report, so it
// isn't bundled into app.js: its ready-made browser build is copied next to
// it as excel.js and fetched on first use. The hash in its URL makes a new
// library version a new file for every cache on the way.
const excelSrc = fs.readFileSync("node_modules/exceljs/dist/exceljs.min.js");
fs.writeFileSync("excel.js", excelSrc);
const excelVersion = crypto.createHash("sha1").update(excelSrc).digest("hex").slice(0, 10);

const watch = process.argv.includes("--watch");

const options = {
  entryPoints: ["src/app.jsx"],
  bundle: true,
  minify: true,
  target: ["ios14", "chrome100"],
  outfile: "app.js",
  loader: { ".jsx": "jsx" },
  jsx: "automatic",
  define: { __EXCEL_URL__: JSON.stringify(`excel.js?v=${excelVersion}`) },
  logLevel: "info",
  plugins: [{ name: "stamp", setup: (b) => b.onEnd((r) => r.errors.length || stamp()) }],
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
