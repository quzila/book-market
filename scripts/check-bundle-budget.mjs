import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";

const root = process.cwd();
const assetsDir = path.join(root, "dist", "assets");

if (!fs.existsSync(assetsDir)) {
  console.error("dist/assets が見つかりません。先に `npm run build` を実行してください。");
  process.exit(1);
}

const files = fs.readdirSync(assetsDir);
const jsFiles = files.filter((name) => name.endsWith(".js"));
const cssFiles = files.filter((name) => name.endsWith(".css"));

if (!jsFiles.length) {
  console.error("JavaScript アセットが見つかりません。");
  process.exit(1);
}

const gzipSize = (absolutePath) => zlib.gzipSync(fs.readFileSync(absolutePath), { level: 9 }).length;

const jsSizes = jsFiles.map((name) => {
  const absolutePath = path.join(assetsDir, name);
  return {
    name,
    raw: fs.statSync(absolutePath).size,
    gzip: gzipSize(absolutePath),
  };
});

const cssSizes = cssFiles.map((name) => {
  const absolutePath = path.join(assetsDir, name);
  return {
    name,
    raw: fs.statSync(absolutePath).size,
    gzip: gzipSize(absolutePath),
  };
});

const totalJsGzip = jsSizes.reduce((sum, row) => sum + row.gzip, 0);
const totalCssGzip = cssSizes.reduce((sum, row) => sum + row.gzip, 0);
const mainJs = jsSizes.slice().sort((a, b) => b.gzip - a.gzip)[0];

const budget = {
  mainJsGzip: 18000,
  totalJsGzip: 26000,
  totalCssGzip: 7000,
};

console.log("Bundle budget check");
console.log(`- main js gzip: ${mainJs.gzip} bytes (${mainJs.name}) / budget ${budget.mainJsGzip}`);
console.log(`- total js gzip: ${totalJsGzip} bytes / budget ${budget.totalJsGzip}`);
console.log(`- total css gzip: ${totalCssGzip} bytes / budget ${budget.totalCssGzip}`);

const failures = [];
if (mainJs.gzip > budget.mainJsGzip) failures.push("main js gzip exceeds budget");
if (totalJsGzip > budget.totalJsGzip) failures.push("total js gzip exceeds budget");
if (totalCssGzip > budget.totalCssGzip) failures.push("total css gzip exceeds budget");

if (failures.length) {
  console.error(`Budget check failed: ${failures.join(", ")}`);
  process.exit(1);
}
