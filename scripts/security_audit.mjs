import fs from "node:fs";
import path from "node:path";
const ROOT = process.cwd();
const PUBLIC_DIR = path.join(ROOT, "public");
const HEADERS_FILE = path.join(PUBLIC_DIR, "_headers");
const patterns = [
  ["email", /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i],
  ["phone", /01[016789][-\s]?\d{3,4}[-\s]?\d{4}/],
  ["resident-registration-number", /\d{6}-[1-4]\d{6}/],
  ["secret-keyword", /(password|passwd|secret|api[_-]?key|access[_-]?token|refresh[_-]?token|github_pat|cloudflare)/i],
];
const requiredHeaders = ["Strict-Transport-Security","X-Frame-Options","X-Content-Type-Options","Referrer-Policy","Permissions-Policy","Cross-Origin-Opener-Policy","Cross-Origin-Resource-Policy","Content-Security-Policy"];
function walk(dir){return fs.readdirSync(dir,{withFileTypes:true}).flatMap(e=>{const p=path.join(dir,e.name); return e.isDirectory()?walk(p):[p];});}
const findings=[];
for (const file of walk(PUBLIC_DIR)) {
  if (!/\.(html|js|css|json|txt)$/i.test(file)) continue;
  const content = fs.readFileSync(file,"utf8");
  for (const [label, pattern] of patterns) {
    const m = content.match(pattern);
    if (m) findings.push(`${path.relative(ROOT,file)}: ${label} (${m[0].slice(0,80)})`);
  }
}
const headers = fs.readFileSync(HEADERS_FILE,"utf8");
const missing = requiredHeaders.filter(h=>!headers.includes(`${h}:`));
if (findings.length || missing.length) {
  if (findings.length) console.error("Sensitive data pattern findings:\n" + findings.map(x=>`- ${x}`).join("\n"));
  if (missing.length) console.error("Missing security headers:\n" + missing.map(x=>`- ${x}`).join("\n"));
  process.exit(1);
}
console.log("Security audit passed");
