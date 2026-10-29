/**
 * SVG をスライドに挿入し、ソースをタグと代替テキストとして図形に保存する。
 */
import { TAG, DATA_FORMAT_VERSION } from "../core/tags";
import type { Box } from "../core/layout";

export interface WriteRequest {
  slideId: string;
  blockId: string;
  source: string;
  hash: string;
  theme: string;
  mermaidVersion: string;
  svg: string;
  box: Box;
}

function setSvg(svg: string, box: Box): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      svg,
      {
        coercionType: Office.CoercionType.XmlSvg,
        imageLeft: box.left,
        imageTop: box.top,
        imageWidth: box.width,
        imageHeight: box.height,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(`SVG の挿入に失敗: ${result.error.message}`));
      },
    );
  });
}

/**
 * setSelectedDataAsync は表示中のスライドに挿入するため、対象のスライドを選択しておく。
 * 図形の選択も外れるので、選択中の図形が置き換わることもない。
 */
async function showSlide(slideId: string): Promise<string[]> {
  return PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id");
    await context.sync();
    return shapes.items.map((s) => s.id);
  });
}

async function shapeIds(slideId: string): Promise<string[]> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id");
    await context.sync();
    return shapes.items.map((s) => s.id);
  });
}

/** 図を挿入し、挿入した図形の id を返す */
export async function insertDiagram(req: WriteRequest): Promise<string> {
  if (!Office.context.requirements.isSetSupported("ImageCoercion", "1.2")) {
    throw new Error("この PowerPoint は SVG 挿入（ImageCoercion 1.2）に対応していません");
  }
  const before = new Set(await showSlide(req.slideId));
  await setSvg(req.svg, req.box);
  // 挿入した図形は、挿入前後の図形 id の差分で特定する（実機で確認済み）
  const added = (await shapeIds(req.slideId)).filter((id) => !before.has(id));
  if (added.length !== 1) throw new Error(`挿入した図形を特定できません（増えた図形 ${added.length} 個）`);

  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides.getItem(req.slideId).shapes.getItem(added[0]);
    const tags: Record<string, string> = {
      [TAG.ID]: req.blockId,
      [TAG.SRC]: req.source,
      [TAG.HASH]: req.hash,
      [TAG.VER]: DATA_FORMAT_VERSION,
      [TAG.MERMAID]: req.mermaidVersion,
      [TAG.THEME]: req.theme,
    };
    for (const [key, value] of Object.entries(tags)) shape.tags.add(key, value);
    shape.name = `Mermaid ${req.blockId}`;
    if (Office.context.requirements.isSetSupported("PowerPointApi", "1.10")) {
      shape.altTextTitle = `Mermaid: ${req.blockId}`;
      shape.altTextDescription = req.source;
    }
    await context.sync();

    return added[0];
  });
}

export async function deleteShape(slideId: string, shapeId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.slides.getItem(slideId).shapes.getItem(shapeId).delete();
    await context.sync();
  });
}

export async function selectSlide(slideId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
  });
}
