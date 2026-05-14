import fs from "node:fs";
const required = ["public/index.html", "public/style.css", "public/app.js", "public/_headers", "public/_redirects"];
const missing = required.filter((file) => !fs.existsSync(file));
if (missing.length) {
  console.error("Missing files:\n" + missing.map((file) => `- ${file}`).join("\n"));
  process.exit(1);
}
const html = fs.readFileSync("public/index.html", "utf8");
for (const token of ["xlsx.full.min.js", "pdf.min.js", "app.js", "excelFile", "pdfFile"]) {
  if (!html.includes(token)) {
    console.error(`index.html does not include ${token}`);
    process.exit(1);
  }
}
console.log("Smoke test passed");
