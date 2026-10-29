/**
 * スライドノートのテキストから ```mermaid フェンスブロックを取り出す。
 * Office に依存しない純粋関数で、仕様は設計書「ノートの記法仕様」に従う。
 */

export interface MermaidBlock {
  /** ブロックの id。`id=` 指定が無ければ出現順の "1", "2", ... */
  id: string;
  /** フェンスの内側のコード（正規化済み） */
  code: string;
  /** 開始フェンスの行番号（1 始まり、正規化後のテキスト基準） */
  line: number;
}

export interface ParseError {
  line: number;
  message: string;
}

export interface ParseResult {
  blocks: MermaidBlock[];
  errors: ParseError[];
}

const OPEN_FENCE = /^\s*(`{3,}|~{3,})\s*mermaid\b(.*)$/i;
const ID_ATTR = /(?:^|\s)id\s*=\s*["']?([A-Za-z0-9_-]+)["']?/;

/**
 * PowerPoint の自動修正やノート特有の文字を、Mermaid が解釈できる形に戻す。
 */
export function normalizeNotesText(text: string): string {
  return (
    text
      // Office.js の textRange は段落内改行を \v で返すことがある
      .replace(/\r\n?|\v|\u2028|\u2029/g, "\n")
      // 全角バッククォート・全角チルダ
      .replace(/\uFF40/g, "`")
      .replace(/\uFF5E/g, "~")
      // スマートクォート
      .replace(/[\u201C\u201D\u201E\u201F\u2033]/g, '"')
      .replace(/[\u2018\u2019\u201A\u201B\u2032]/g, "'")
      // ノーブレークスペース
      .replace(/\u00A0/g, " ")
  );
}

export function parseMermaidBlocks(rawText: string): ParseResult {
  const lines = normalizeNotesText(rawText).split("\n");
  const blocks: MermaidBlock[] = [];
  const errors: ParseError[] = [];

  let i = 0;
  while (i < lines.length) {
    const open = OPEN_FENCE.exec(lines[i]);
    if (!open) {
      i++;
      continue;
    }
    const fence = open[1];
    const startLine = i + 1;
    const close = new RegExp(`^\\s*${fence[0] === "`" ? "`" : "~"}{${fence.length},}\\s*$`);

    let j = i + 1;
    while (j < lines.length && !close.test(lines[j])) j++;
    if (j >= lines.length) {
      errors.push({ line: startLine, message: "mermaid ブロックの終了行（```）がありません" });
      break;
    }

    const idMatch = ID_ATTR.exec(open[2]);
    blocks.push({
      id: idMatch ? idMatch[1] : String(blocks.length + 1),
      code: lines.slice(i + 1, j).join("\n").trim(),
      line: startLine,
    });
    i = j + 1;
  }

  const seen = new Map<string, number>();
  for (const block of blocks) {
    const first = seen.get(block.id);
    if (first !== undefined) {
      errors.push({
        line: block.line,
        message: `id「${block.id}」が重複しています（${first} 行目と同じ）`,
      });
    } else {
      seen.set(block.id, block.line);
    }
  }

  for (const block of blocks) {
    if (block.code === "") {
      errors.push({ line: block.line, message: `mermaid ブロック「${block.id}」が空です` });
    }
  }

  return { blocks, errors };
}
