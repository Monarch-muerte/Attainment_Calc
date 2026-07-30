const esbuild = require("esbuild");
const path = require("path");

esbuild
  .build({
    entryPoints: [path.join(__dirname, "..", "frontend", "src", "index.jsx")],
    bundle: true,
    minify: true,
    sourcemap: false,
    outfile: path.join(__dirname, "..", "public", "bundle.js"),
    loader: { ".js": "jsx" },
    define: { "process.env.NODE_ENV": '"production"' },
    logLevel: "info",
  })
  .then(() => console.log("Frontend build complete -> public/bundle.js"))
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
