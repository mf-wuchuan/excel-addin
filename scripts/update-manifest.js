/**
 * Updates manifest.xml URLs to point to a given base URL.
 * Usage: node scripts/update-manifest.js <base-url>
 * Example: node scripts/update-manifest.js https://your-username.github.io/excel-addin/dist
 */
const fs = require("fs");
const path = require("path");

const baseUrl = process.argv[2];
if (!baseUrl) {
  console.error("Usage: node scripts/update-manifest.js <base-url>");
  console.error("Example: node scripts/update-manifest.js https://user.github.io/excel-addin/dist");
  process.exit(1);
}

const url = baseUrl.replace(/\/$/, "");
const manifestPath = path.resolve(__dirname, "../manifest.xml");
let manifest = fs.readFileSync(manifestPath, "utf8");

// Replace all localhost:3000 references with the new base URL
manifest = manifest.replace(/https:\/\/localhost:3000/g, url);

const outPath = path.resolve(__dirname, "../manifest-prod.xml");
fs.writeFileSync(outPath, manifest, "utf8");
console.log(`Production manifest written to: ${outPath}`);
console.log(`Base URL: ${url}`);
