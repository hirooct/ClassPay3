import fs from "node:fs";
import crypto from "node:crypto";

const manifest = JSON.parse(fs.readFileSync("release/manifest.json", "utf8"));
let failed = false;
const sha256 = path => crypto.createHash("sha256").update(fs.readFileSync(path)).digest("hex");

for (const file of manifest.files) {
  const path = file.name === "appsscript" ? "appsscript.json" : file.name + (file.type === "HTML" ? ".html" : ".js");
  if (!fs.existsSync(path) || sha256(path) !== file.sha256) {
    console.error(`release hash mismatch: ${path}`);
    failed = true;
  }
}

for (const migration of manifest.migrations || []) {
  const path = `release/migrations/${migration.from}-to-${migration.to}.json`;
  if (!fs.existsSync(path) || sha256(path) !== migration.sha256) {
    console.error(`migration hash mismatch: ${path}`);
    failed = true;
  }
}

if (failed) process.exit(1);
console.log(`verified ${manifest.files.length} release files and ${(manifest.migrations || []).length} migrations`);
