// A SheetJS (xlsx) könyvtárat bemásolja a www/lib mappába,
// hogy az app internet nélkül is be tudja olvasni az Excel-fájlokat.
const fs = require("fs");
const path = require("path");

const src = path.join(__dirname, "..", "node_modules", "xlsx", "dist", "xlsx.full.min.js");
const destDir = path.join(__dirname, "..", "www", "lib");

if (!fs.existsSync(src)) {
  console.error("Nem találom: " + src + "\nFuttasd előbb: npm install");
  process.exit(1);
}
fs.mkdirSync(destDir, { recursive: true });
fs.copyFileSync(src, path.join(destDir, "xlsx.full.min.js"));
console.log("xlsx.full.min.js bemásolva a www/lib mappába.");
