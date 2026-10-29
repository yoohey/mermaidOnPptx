/* eslint-disable no-undef */
/**
 * 開発用 manifest.xml（localhost）から配布用のマニフェストを作り、利用者向けページ（site/）を添える。
 *   node scripts/build-manifest.js [出力先フォルダー]   既定は dist/
 * 配信元は環境変数 ADDIN_BASE_URL で変えられる（末尾の / は不要）。
 */
const fs = require("fs");
const path = require("path");

const DEV_ORIGIN = "https://localhost:3000";
const DEV_REPO_URL = "https://github.com/yoohey/mermaid4pptx";
const DEFAULT_BASE_URL = "https://yoohey.github.io/mermaidOnPptx";
// 開発版と並べて入れられるよう、配布版は別の Id にする。一度配ったら変えないこと
const RELEASE_ID = "5d0c6a3e-8f2b-4c71-9a46-2e7b1f3c9d58";

function buildReleaseManifest(devXml, baseUrl = DEFAULT_BASE_URL) {
  const base = baseUrl.replace(/\/+$/, "");
  if (!/^https:\/\//.test(base)) throw new Error(`配信元は https で指定してください: ${baseUrl}`);
  const origin = new URL(base).origin;

  let xml = devXml
    .replace(/<Id>[^<]*<\/Id>/, `<Id>${RELEASE_ID}</Id>`)
    .replace(/(<DisplayName DefaultValue="[^"]*?)\s*\(dev\)"/, '$1"')
    // AppDomain はパスを含まないオリジンで書く
    .replace(`<AppDomain>${DEV_ORIGIN}</AppDomain>`, `<AppDomain>${origin}</AppDomain>`)
    // 開発用リポジトリは非公開なので、サポートと「詳細」は配布サイトのページに向ける
    .replace(`<SupportUrl DefaultValue="${DEV_REPO_URL}"/>`, `<SupportUrl DefaultValue="${base}/support.html"/>`)
    .split(`DefaultValue="${DEV_REPO_URL}"`)
    .join(`DefaultValue="${base}/"`)
    .split(DEV_ORIGIN)
    .join(base);

  if (xml.includes("localhost")) throw new Error("localhost が残っています");
  if (xml.includes(DEV_REPO_URL)) throw new Error("非公開リポジトリの URL が残っています");
  return xml;
}

/** https://<owner>.github.io/<repo> の配信元から、公開用リポジトリの URL を作る */
function repoUrlOf(baseUrl) {
  const m = /^https:\/\/([^.]+)\.github\.io\/([^/]+)/.exec(baseUrl);
  return m ? `https://github.com/${m[1]}/${m[2]}` : baseUrl;
}

/** site/ のページの __BASE_URL__ と __REPO_URL__ を埋める */
function fillSiteTemplate(text, baseUrl = DEFAULT_BASE_URL) {
  const base = baseUrl.replace(/\/+$/, "");
  return text.split("__BASE_URL__").join(base).split("__REPO_URL__").join(repoUrlOf(base));
}

module.exports = { buildReleaseManifest, fillSiteTemplate, repoUrlOf, DEFAULT_BASE_URL, RELEASE_ID };

if (require.main === module) {
  const root = path.join(__dirname, "..");
  const outDir = path.resolve(root, process.argv[2] || "dist");
  const base = process.env.ADDIN_BASE_URL || DEFAULT_BASE_URL;
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(
    path.join(outDir, "manifest.xml"),
    buildReleaseManifest(fs.readFileSync(path.join(root, "manifest.xml"), "utf8"), base),
  );
  const siteDir = path.join(root, "site");
  for (const name of fs.readdirSync(siteDir)) {
    const text = fillSiteTemplate(fs.readFileSync(path.join(siteDir, name), "utf8"), base);
    // Windows PowerShell 5.1 は BOM の無いスクリプトを UTF-8 として読まず、日本語が化ける
    fs.writeFileSync(path.join(outDir, name), name.endsWith(".ps1") ? "\uFEFF" + text : text);
  }
  console.log(`配布用マニフェストと利用者向けページを書き出しました: ${path.relative(root, outDir)}（配信元 ${base}）`);
}
