/* eslint-disable no-undef */
/**
 * 公開用リポジトリに送る README を、利用者向けの内容だけにする。
 *   node scripts/public-readme.js <README.md> [配信元 URL]
 * 開発用リポジトリの話（設計メモ、開発版の試し方、開発・配布の手順）を除き、導入はサイトに案内する。
 */
const fs = require("fs");
const { DEFAULT_BASE_URL } = require("./build-manifest.js");

function publicReadme(text, baseUrl = DEFAULT_BASE_URL) {
  const base = baseUrl.replace(/\/+$/, "");
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  const out = [];
  let skipping = false;
  for (const line of lines) {
    if (/^## /.test(line)) {
      skipping = ["## Windows で試す", "## 開発", "## 配布版の公開", "## 導入"].includes(line);
      if (line === "## ライセンス") out.push("## 導入", "", `導入の手順は <${base}/> をご覧ください。`, "");
    }
    if (skipping || line.startsWith("> 設計:")) continue;
    out.push(line);
  }
  if (out[0].startsWith("# ")) out[0] = "# Mermaid on PowerPoint";
  return out.join("\n").replace(/\n{3,}/g, "\n\n");
}

module.exports = { publicReadme };

if (require.main === module) {
  const [file, base] = process.argv.slice(2);
  fs.writeFileSync(file, publicReadme(fs.readFileSync(file, "utf8"), base || DEFAULT_BASE_URL));
}
