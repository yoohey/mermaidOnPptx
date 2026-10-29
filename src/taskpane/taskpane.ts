/**
 * タスクペインとリボンの関数コマンド。共有ランタイムで同じページとして動く。
 */
import { runSync, type ItemReport, type SyncOptions, type SyncReport } from "../app/syncRunner";
import { SAMPLES, fencedSample } from "../core/samples";
import {
  DEFAULT_CUSTOM_COLORS,
  THEMES,
  THEME_LABELS,
  customFromTemplate,
  customFromThemeKey,
  isTheme,
  paletteFor,
  type CustomColors,
} from "../core/theme";
import { renderMermaid } from "../office/renderer";
import { readPresentationNotes } from "../office/notesSource";
import { getCustomColors, getTheme, hasCustomColors, setCustomColors, setTheme } from "../office/settings";
import { findDiagramThemeKey } from "../office/diagrams";
import { getCurrentSlide } from "../office/notesSource";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const STATUS_LABEL: Record<ItemReport["status"], string> = {
  inserted: "新規",
  replaced: "更新",
  upToDate: "最新",
  orphan: "対応なし",
  error: "エラー",
};

let running = false;

function setBusy(busy: boolean) {
  for (const id of ["sync-current", "sync-all", "theme", "color-from-template"]) {
    $<HTMLButtonElement>(id).disabled = busy;
  }
}

function summaryText(report: SyncReport): string {
  const c = report.counts;
  if (report.slides.length === 0) return "mermaid ブロックが見つかりませんでした。";
  const parts = [`新規 ${c.inserted}`, `更新 ${c.replaced}`, `最新 ${c.upToDate}`, `エラー ${c.error}`];
  if (c.orphan > 0) parts.push(`対応なし ${c.orphan}`);
  return parts.join("、");
}

function renderItem(item: ItemReport): HTMLElement {
  const el = document.createElement("div");
  el.className = `item ${item.status}`;
  const head = document.createElement("div");
  const status = document.createElement("span");
  status.className = "status";
  status.textContent = STATUS_LABEL[item.status];
  head.append(status, item.blockId ? `id=${item.blockId}` : "");
  if (item.line !== undefined) head.append(`（ノート ${item.line} 行目）`);
  el.append(head);
  if (item.message) {
    const pre = document.createElement("pre");
    pre.textContent = item.message;
    el.append(pre);
  }
  if (item.svg) el.insertAdjacentHTML("beforeend", item.svg);
  return el;
}

function showReport(report: SyncReport) {
  $("summary").textContent = summaryText(report);
  const results = $("results");
  results.replaceChildren();
  for (const slide of report.slides) {
    const h = document.createElement("h3");
    h.textContent = `スライド ${slide.number}`;
    results.append(h, ...slide.items.map(renderItem));
  }
}

function showFailure(e: unknown) {
  $("summary").textContent = "実行できませんでした。";
  const el = document.createElement("div");
  el.className = "item error";
  el.textContent = e instanceof Error ? e.message : String(e);
  $("results").replaceChildren(el);
}

/** 二重実行を防いで同期を走らせ、結果をパネルに出す。結果を返す（失敗時は null） */
async function sync(options: SyncOptions): Promise<SyncReport | null> {
  if (running) return null;
  running = true;
  setBusy(true);
  $("summary").textContent = "実行中…";
  try {
    const report = await runSync(options);
    showReport(report);
    return report;
  } catch (e) {
    showFailure(e);
    return null;
  } finally {
    running = false;
    setBusy(false);
  }
}

/** リボンのメニューから実行する。エラーがあればパネルを開いて内容を見せる */
function commandFor(scope: SyncOptions["scope"]) {
  return async (event: Office.AddinCommands.Event) => {
    try {
      const report = await sync({ scope, force: false });
      if (!report || report.counts.error > 0 || report.slides.length === 0) await Office.addin.showAsTaskpane();
    } finally {
      event.completed();
    }
  };
}

function initThemeSelect() {
  const select = $<HTMLSelectElement>("theme");
  for (const theme of THEMES) select.add(new Option(THEME_LABELS[theme], theme));
  select.value = getTheme();
  $("custom-colors").hidden = select.value !== "custom";
  select.onchange = async () => {
    if (!isTheme(select.value)) return;
    // 初めてカスタムにするときは、今の図の色合い（図が無ければテンプレートの色）から始める
    if (select.value === "custom" && !hasCustomColors()) {
      await setCustomColors(await initialCustomColors());
      showCustomColors(getCustomColors());
    }
    await setTheme(select.value);
    $("custom-colors").hidden = select.value !== "custom";
    themeChanged();
  };
}

function themeChanged() {
  if ($<HTMLDetailsElement>("samples").open) void showSample();
  $("summary").textContent = "テーマを変えました。「全スライドを更新」で既存の図にも反映されます。";
}

const COLOR_INPUTS: Record<keyof CustomColors, string> = { main: "color-main", sub: "color-sub", text: "color-text" };

function showCustomColors(colors: CustomColors) {
  for (const [key, id] of Object.entries(COLOR_INPUTS)) {
    $<HTMLInputElement>(id).value = colors[key as keyof CustomColors].toLowerCase();
  }
}

/** カスタムの初期値。既にある図のテーマの色、図が無ければテンプレートの色、どちらも無理なら既定の色 */
async function initialCustomColors(): Promise<CustomColors> {
  try {
    const current = await getCurrentSlide().catch(() => null);
    const key = await findDiagramThemeKey(current?.id);
    const fromDiagram = key !== null ? customFromThemeKey(key) : null;
    if (fromDiagram) return fromDiagram;
    const template = (await readPresentationNotes()).themeColors;
    if (template) return customFromTemplate(template);
  } catch {
    // 読めなければ既定の色にする
  }
  return DEFAULT_CUSTOM_COLORS;
}

async function saveCustomColors(colors: CustomColors) {
  await setCustomColors(colors);
  showCustomColors(colors);
  themeChanged();
}

/** カスタムの色。選んだ色はファイルに保存され、ファイルと一緒に持ち運べる */
function initCustomColors() {
  showCustomColors(getCustomColors());
  for (const [key, id] of Object.entries(COLOR_INPUTS)) {
    // input は選んでいる間ずっと発生するので、決めたとき（change）だけ保存する
    $<HTMLInputElement>(id).onchange = (e) =>
      void saveCustomColors({ ...getCustomColors(), [key]: (e.target as HTMLInputElement).value.toUpperCase() });
  }
  $("color-from-template").onclick = async () => {
    const template = (await readPresentationNotes()).themeColors;
    if (!template) {
      $("summary").textContent = "このファイルのテーマの配色を読めませんでした。";
      return;
    }
    await saveCustomColors(customFromTemplate(template));
  };
}

/** 選んだサンプルの全文とプレビューを出す。プレビューは今のテーマで描く */
async function showSample() {
  const sample = SAMPLES.find((s) => s.id === $<HTMLSelectElement>("sample-select").value) ?? SAMPLES[0];
  $<HTMLTextAreaElement>("sample-code").value = fencedSample(sample);
  $("sample-status").textContent = "";
  const preview = $("sample-preview");
  try {
    const theme = getTheme();
    // 「テンプレートの色」はプレゼンの配色が要るので、そのときだけ pptx を読む
    const template = theme === "template" ? (await readPresentationNotes()).themeColors : null;
    const { svg } = await renderMermaid(sample.code, theme, paletteFor(theme, template, getCustomColors()));
    preview.innerHTML = svg;
  } catch (e) {
    preview.textContent = e instanceof Error ? e.message : String(e);
  }
}

/** クリップボードに書けない環境では、全文を選択して Ctrl+C を案内する */
async function copySample() {
  const area = $<HTMLTextAreaElement>("sample-code");
  const status = $("sample-status");
  try {
    await navigator.clipboard.writeText(area.value);
    status.textContent = "コピーしました。ノートに貼り付けてください。";
    return;
  } catch {
    // 下の execCommand に任せる
  }
  area.focus();
  area.select();
  status.textContent = document.execCommand("copy")
    ? "コピーしました。ノートに貼り付けてください。"
    : "コピーできませんでした。選択された状態なので Ctrl+C でコピーしてください。";
}

function initSamples() {
  const select = $<HTMLSelectElement>("sample-select");
  for (const sample of SAMPLES) select.add(new Option(sample.label, sample.id));
  select.onchange = () => void showSample();
  $("sample-copy").onclick = () => void copySample();
  // 開いたときだけ描く（パネルを開くたびに mermaid を動かさない）
  $<HTMLDetailsElement>("samples").ontoggle = () => {
    if ($<HTMLDetailsElement>("samples").open) void showSample();
  };
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    $("summary").textContent = "PowerPoint 以外では動作しません";
    return;
  }
  Office.actions.associate("syncCurrentSlide", commandFor("current"));
  Office.actions.associate("syncAllSlides", commandFor("all"));
  $("sync-current").onclick = () => sync({ scope: "current", force: $<HTMLInputElement>("force").checked });
  $("sync-all").onclick = () => sync({ scope: "all", force: $<HTMLInputElement>("force").checked });
  initThemeSelect();
  initCustomColors();
  initSamples();
  setBusy(false);
});
