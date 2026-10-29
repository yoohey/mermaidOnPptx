/* eslint-disable no-undef */
const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { LicenseWebpackPlugin } = require("license-webpack-plugin");

const mermaidVersion = JSON.parse(
  fs.readFileSync(path.join(__dirname, "node_modules/mermaid/package.json"), "utf8"),
).version;

// langium（mermaid の依存）は vscode-jsonrpc を宣言せずに import している。
// 開発ツール側が別版（4.x）を最上位に置くと誤解決されるため、langium 系が本来使う版に固定する。
// package.json を exports で隠しているパッケージがあるため、Node の探索順を自前でたどる
const findPkg = (name, from) => {
  for (let dir = from; ; dir = path.dirname(dir)) {
    const candidate = path.join(dir, "node_modules", name);
    if (fs.existsSync(path.join(candidate, "package.json"))) return candidate;
    if (path.dirname(dir) === dir) throw new Error(`${name} が見つかりません（${from} から探索）`);
  }
};
const langiumDir = findPkg("langium", findPkg("@mermaid-js/parser", findPkg("mermaid", __dirname)));
const jsonrpcDir = findPkg(
  "vscode-jsonrpc",
  findPkg("vscode-languageserver-protocol", findPkg("vscode-languageserver", langiumDir)),
);

// 「MIT OR GPL」のような選択式は許可し、GPL 系だけのライセンスはビルドを止める
const isUnacceptable = (license) => /GPL/i.test(license) && !/\bOR\b/i.test(license);

async function httpsOptions() {
  const devCerts = require("office-addin-dev-certs");
  return devCerts.getHttpsServerOptions();
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      // 配布版はファイル名に内容のハッシュを付け、更新後に古いコードがキャッシュから読まれないようにする
      filename: dev ? "[name].js" : "[name].[contenthash:8].js",
      chunkFilename: dev ? "[id].js" : "[id].[contenthash:8].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
      alias: { "vscode-jsonrpc": jsonrpcDir },
    },
    module: {
      rules: [
        { test: /\.ts$/, exclude: /node_modules/, use: { loader: "ts-loader", options: { transpileOnly: true } } },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({ __MERMAID_VERSION__: JSON.stringify(mermaidVersion) }),
      new HtmlWebpackPlugin({ filename: "taskpane.html", template: "./src/taskpane/taskpane.html", chunks: ["taskpane"] }),
      new CopyWebpackPlugin({ patterns: [
          { from: "assets", to: "assets" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
        ] }),
      // 第三者ライセンス一覧は配布物にだけ必要。開発サーバー（Windows）ではパッケージ探索が
      // 空パスで失敗することがあったため、本番ビルドでのみ生成する。
      !dev &&
        new LicenseWebpackPlugin({
          outputFilename: "THIRD-PARTY-NOTICES.txt",
          perChunkOutput: false,
          // khroma は package.json に license 欄が無いが、同梱の license ファイルは MIT
          licenseTypeOverrides: { khroma: "MIT" },
          unacceptableLicenseTest: isUnacceptable,
          handleUnacceptableLicense: (pkg, license) => {
            throw new Error(`許可していないライセンスです: ${pkg} (${license})`);
          },
          handleMissingLicenseType: (pkg) => {
            console.warn(`[license] ライセンス種別が不明: ${pkg}`);
            return null;
          },
        }),
    ].filter(Boolean),
    // モジュール連結を有効にすると、連結された d3-* などが第三者ライセンス一覧から漏れるため無効にする
    optimization: { concatenateModules: false },
    performance: { hints: false },
    devServer: {
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: env.WEBPACK_BUILD || !dev ? {} : await httpsOptions() },
    },
  };
};
