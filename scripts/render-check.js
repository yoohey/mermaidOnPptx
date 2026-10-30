/* eslint-disable no-undef */
/**
 * 全サンプル × 全テーマを、アドインと同じ設定で Chromium 上に描画して確かめる。
 *   node scripts/render-check.js [結果の JSON]
 * 失敗（描画エラー、foreignObject や dominant-baseline が残る、%%{init}%% が効かない）があると終了コード 1。
 * 枠を広げて作り直したとき、文字の大きさを変えずにどれだけ枠を埋められるかも記録する。
 * 図形の色が CSS だけで付いているか（PowerPoint 以外での表示に関係）は数を記録するだけで、失敗にはしない。
 * Chromium の場所は環境変数 CHROMIUM_PATH で指定できる。
 */
const fs = require("fs");
const os = require("os");
const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const { chromium } = require("playwright");

const root = path.join(__dirname, "..");

async function buildHarness(outDir) {
  const base = await require(path.join(root, "webpack.config.js"))({ WEBPACK_BUILD: true }, { mode: "development" });
  const config = {
    ...base,
    mode: "development",
    context: root,
    entry: { harness: "./test/render/harness.ts" },
    output: { path: outDir, filename: "[name].js", clean: true },
    plugins: [
      new webpack.DefinePlugin({ __MERMAID_VERSION__: JSON.stringify("render-check") }),
      new HtmlWebpackPlugin({ filename: "index.html", template: "./test/render/harness.html" }),
    ],
    devServer: undefined,
  };
  await new Promise((resolve, reject) =>
    webpack(config, (err, stats) => {
      if (err) return reject(err);
      if (stats.hasErrors()) return reject(new Error(stats.toString("errors-only")));
      resolve();
    }),
  );
}

// ブロック先頭の %%{init}%% が、パネルで選んだテーマより優先されること。forest の primaryColor は #cde498
const INIT_CASE = {
  id: "init-forest",
  code: '%%{init: {"theme": "forest"}}%%\nflowchart LR\n  A --> B',
  expect: /#cde498/i,
};

async function main() {
  const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "m4p-render-"));
  await buildHarness(outDir);

  const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
  const page = await browser.newPage();
  await page.goto("file://" + path.join(outDir, "index.html"));
  await page.waitForFunction(() => "m4pRender" in window);

  const { samples, themes } = await page.evaluate(() => ({ samples: window.m4pSamples, themes: window.m4pThemes }));
  const results = [];
  const failures = [];

  for (const theme of themes) {
    for (const sample of samples) {
      const t0 = Date.now();
      const r = await page.evaluate(([code, th]) => window.m4pRender(code, th), [sample.code, theme]);
      const row = { theme, sample: sample.id, ms: Date.now() - t0, ...r, svg: undefined };
      results.push(row);
      if (!r.ok) failures.push(`${theme} / ${sample.id}: ${r.error}`);
      else if (r.foreignObjects > 0) failures.push(`${theme} / ${sample.id}: foreignObject が ${r.foreignObjects} 個`);
      // PowerPoint は dominant-baseline が効かず文字が上にずれるため、dy に置き換えてあること
      else if (/<text[^>]*dominant-baseline="(central|middle)"/.test(r.svg))
        failures.push(`${theme} / ${sample.id}: dominant-baseline が残っている`);
    }
  }

  // 枠を広げて作り直したとき、文字の大きさを変えずにどれだけ枠を埋められるか（記録だけ）
  const FILL_CASES = [
    { label: "幅2倍", w: 2, h: 1 },
    { label: "高さ2倍", w: 1, h: 2 },
    { label: "幅高さ2倍", w: 2, h: 2 },
    { label: "0.7倍", w: 0.7, h: 0.7 },
  ];
  const fills = [];
  for (const sample of samples) {
    for (const c of FILL_CASES) {
      const r = await page.evaluate(([code, w, h]) => window.m4pFill(code, w, h), [sample.code, c.w, c.h]);
      fills.push({ sample: sample.id, case: c.label, ...r });
      if (!r.ok) failures.push(`${sample.id} / ${c.label}: ${r.error}`);
    }
  }

  const init = await page.evaluate(([code]) => window.m4pRender(code, "modern"), [INIT_CASE.code]);
  if (!init.ok || !INIT_CASE.expect.test(init.svg)) failures.push(`${INIT_CASE.id}: %%{init}%% のテーマが効いていない`);

  await browser.close();
  fs.rmSync(outDir, { recursive: true, force: true });

  console.log("theme      sample     ms    size   色が属性の図形/図形");
  for (const r of results) {
    const color = r.ok ? `${r.shapesWithColorAttr}/${r.shapes}` : "-";
    console.log(
      `${r.theme.padEnd(10)} ${r.sample.padEnd(9)} ${String(r.ms).padStart(5)} ${String(r.bytes ?? "-").padStart(7)}  ${color}${r.ok ? "" : "  ERROR"}`,
    );
  }
  console.log("\nsample     枠          幅を埋めた  高さを埋めた  文字の縮小");
  const pct = (v) => `${Math.round(v * 100)}%`.padStart(6);
  for (const r of fills) {
    const v = r.ok ? `${pct(r.fillW)}      ${pct(r.fillH)}        ${pct(r.shrink)}` : "ERROR";
    console.log(`${r.sample.padEnd(10)} ${r.case.padEnd(10)} ${v}`);
  }
  if (process.argv[2])
    fs.writeFileSync(process.argv[2], JSON.stringify({ results, fills, init: { ...init, svg: undefined } }, null, 2));

  if (failures.length > 0) {
    console.error(`\n失敗 ${failures.length} 件:\n` + failures.map((f) => "  " + f).join("\n"));
    process.exit(1);
  }
  console.log(`\n${results.length} 件と %%{init}%% の優先をすべて描画できました`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
