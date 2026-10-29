/**
 * pptx（zip）からスライドごとのノートのテキストを取り出す。
 * Office.js にノート読み取りの公式 API が無いため、getFileAsync で得た pptx をここで解析する。
 */
import JSZip from "jszip";
import { DOMParser, type Element as XmlElement } from "@xmldom/xmldom";
import type { TemplateColors } from "./theme";

export interface SlideNotes {
  /** presentation.xml の <p:sldId id> */
  sldId: number;
  /** 表示順（0 始まり） */
  index: number;
  /** zip 内のスライドのパス（例 ppt/slides/slide1.xml） */
  slidePath: string;
  /** ノートが無いスライドは null */
  notesText: string | null;
}

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NOTES_REL_SUFFIX = "/notesSlide";

function parseXml(text: string): XmlElement {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (!doc.documentElement) throw new Error("XML を解析できません");
  return doc.documentElement;
}

/** 名前空間接頭辞に依存せず、localName で子孫要素を集める */
function descendants(node: XmlElement, localName: string): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (n: XmlElement) => {
    for (let c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1) {
        const el = c as XmlElement;
        if (el.localName === localName) out.push(el);
        walk(el);
      }
    }
  };
  walk(node);
  return out;
}

function children(node: XmlElement, localName: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (let c = node.firstChild; c; c = c.nextSibling) {
    if (c.nodeType === 1 && (c as XmlElement).localName === localName) out.push(c as XmlElement);
  }
  return out;
}

/** パート内の相対 Target を zip 内の絶対パスに解決する */
export function resolvePartPath(basePartPath: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = basePartPath.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== "." && seg !== "") parts.push(seg);
  }
  return parts.join("/");
}

function relsPathOf(partPath: string): string {
  const dir = partPath.split("/").slice(0, -1).join("/");
  const name = partPath.split("/").pop();
  return `${dir}/_rels/${name}.rels`;
}

async function readRels(zip: JSZip, partPath: string): Promise<Map<string, { type: string; target: string }>> {
  const file = zip.file(relsPathOf(partPath));
  const rels = new Map<string, { type: string; target: string }>();
  if (!file) return rels;
  const root = parseXml(await file.async("string"));
  for (const rel of descendants(root, "Relationship")) {
    rels.set(rel.getAttribute("Id") ?? "", {
      type: rel.getAttribute("Type") ?? "",
      target: rel.getAttribute("Target") ?? "",
    });
  }
  return rels;
}

/** ノートスライドの本文プレースホルダ（type="body"）のテキストを、段落ごとに改行で連結して返す */
export function extractNotesBodyText(notesXml: string): string {
  const root = parseXml(notesXml);
  const texts: string[] = [];
  for (const sp of descendants(root, "sp")) {
    const ph = descendants(sp, "ph")[0];
    if (!ph || ph.getAttribute("type") !== "body") continue;
    const txBody = descendants(sp, "txBody")[0];
    if (!txBody) continue;
    const paragraphs = children(txBody, "p").map((p) => {
      let line = "";
      for (let c = p.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const el = c as XmlElement;
        if (el.localName === "r" || el.localName === "fld") {
          for (const t of children(el, "t")) line += t.textContent ?? "";
        } else if (el.localName === "br") {
          line += "\n";
        }
      }
      return line;
    });
    texts.push(paragraphs.join("\n"));
  }
  return texts.join("\n");
}

export interface PresentationInfo {
  slides: SlideNotes[];
  /** スライドの大きさ（pt）。presentation.xml の <p:sldSz> から求める */
  slideSize: { width: number; height: number };
  /** テーマの配色。読めなければ null */
  themeColors: TemplateColors | null;
}

const THEME_REL_SUFFIX = "/theme";
const MASTER_REL_SUFFIX = "/slideMaster";

/** <a:dk1> などの中の色（srgbClr の val、sysClr の lastClr）を "#RRGGBB" で返す */
function schemeColor(scheme: XmlElement, name: string): string | null {
  const slot = children(scheme, name)[0];
  if (!slot) return null;
  for (let c = slot.firstChild; c; c = c.nextSibling) {
    if (c.nodeType !== 1) continue;
    const el = c as XmlElement;
    const value = el.localName === "srgbClr" ? el.getAttribute("val") : el.getAttribute("lastClr");
    if (value && /^[0-9a-fA-F]{6}$/.test(value)) return "#" + value.toUpperCase();
  }
  return null;
}

/** テーマ XML の <a:clrScheme> を読む。足りない色があれば null */
export function parseThemeColors(themeXml: string): TemplateColors | null {
  const scheme = descendants(parseXml(themeXml), "clrScheme")[0];
  if (!scheme) return null;
  const [dk1, lt1, dk2, lt2] = ["dk1", "lt1", "dk2", "lt2"].map((n) => schemeColor(scheme, n));
  const accents = [1, 2, 3, 4, 5, 6].map((i) => schemeColor(scheme, `accent${i}`));
  if (!dk1 || !lt1 || !dk2 || !lt2 || accents.some((a) => !a)) return null;
  return { dk1, lt1, dk2, lt2, accents: accents as string[] };
}

/**
 * スライドが使うテーマの配色を読む。マスターが複数あるときは最初のマスターのテーマを使う。
 * マスターからたどれなければ presentation.xml に直接ひも付くテーマを使う。
 */
async function readThemeColors(
  zip: JSZip,
  presPath: string,
  presRels: Map<string, { type: string; target: string }>,
): Promise<TemplateColors | null> {
  const candidates: string[] = [];
  const master = [...presRels.values()].find((r) => r.type.endsWith(MASTER_REL_SUFFIX));
  if (master) {
    const masterPath = resolvePartPath(presPath, master.target);
    const theme = [...(await readRels(zip, masterPath)).values()].find((r) => r.type.endsWith(THEME_REL_SUFFIX));
    if (theme) candidates.push(resolvePartPath(masterPath, theme.target));
  }
  for (const r of presRels.values()) {
    if (r.type.endsWith(THEME_REL_SUFFIX)) candidates.push(resolvePartPath(presPath, r.target));
  }
  for (const path of candidates) {
    const file = zip.file(path);
    const colors = file ? parseThemeColors(await file.async("string")) : null;
    if (colors) return colors;
  }
  return null;
}

const EMU_PER_PT = 12700;
/** <p:sldSz> が無いときの既定（16:9 標準） */
const DEFAULT_SLIDE_SIZE = { width: 960, height: 540 };

export async function readSlideNotes(pptx: Uint8Array | ArrayBuffer): Promise<SlideNotes[]> {
  return (await readPresentationInfo(pptx)).slides;
}

export async function readPresentationInfo(pptx: Uint8Array | ArrayBuffer): Promise<PresentationInfo> {
  const zip = await JSZip.loadAsync(pptx);
  const presPath = "ppt/presentation.xml";
  const presFile = zip.file(presPath);
  if (!presFile) throw new Error("ppt/presentation.xml が見つかりません");

  const pres = parseXml(await presFile.async("string"));
  const sldSz = descendants(pres, "sldSz")[0];
  const cx = Number(sldSz?.getAttribute("cx"));
  const cy = Number(sldSz?.getAttribute("cy"));
  const slideSize = cx > 0 && cy > 0 ? { width: cx / EMU_PER_PT, height: cy / EMU_PER_PT } : DEFAULT_SLIDE_SIZE;
  const presRels = await readRels(zip, presPath);
  const result: SlideNotes[] = [];

  const sldIds = descendants(pres, "sldId");
  for (let index = 0; index < sldIds.length; index++) {
    const el = sldIds[index];
    const rId = el.getAttributeNS(REL_NS, "id") ?? el.getAttribute("r:id") ?? "";
    const rel = presRels.get(rId);
    if (!rel) continue;
    const slidePath = resolvePartPath(presPath, rel.target);

    let notesText: string | null = null;
    const slideRels = await readRels(zip, slidePath);
    for (const r of slideRels.values()) {
      if (!r.type.endsWith(NOTES_REL_SUFFIX)) continue;
      const notesFile = zip.file(resolvePartPath(slidePath, r.target));
      if (notesFile) notesText = extractNotesBodyText(await notesFile.async("string"));
    }

    result.push({ sldId: Number(el.getAttribute("id")), index, slidePath, notesText });
  }
  return { slides: result, slideSize, themeColors: await readThemeColors(zip, presPath, presRels) };
}

/**
 * Office.js の Slide.id（例 "267#" や "267#1234567"）から sldId の数値部分を取り出す。
 * 形式は公式に文書化されていないため、取れなければ null を返し、呼び出し側は表示順で対応付ける。
 */
export function sldIdFromOfficeSlideId(officeId: string): number | null {
  const m = /^(\d+)/.exec(officeId);
  return m ? Number(m[1]) : null;
}
