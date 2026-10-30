/**
 * SVG をスライドに挿入し、ソースをタグと代替テキストとして図形に保存する。
 */
import { TAG, DATA_FORMAT_VERSION } from "../core/tags";
import { formatBox, type Box } from "../core/layout";
import type { ExistingDiagram } from "../core/sync";

export interface WriteRequest {
  slideId: string;
  blockId: string;
  source: string;
  hash: string;
  theme: string;
  mermaidVersion: string;
  svg: string;
  box: Box;
  /** 指定した文字の大きさ（pt） */
  fontPt: number;
  /** スライド上での実際の文字の大きさ（pt） */
  textPt: number;
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
      [TAG.BOX]: formatBox(req.box),
      [TAG.FONT_PT]: String(req.fontPt),
      [TAG.TEXT_PT]: req.textPt.toFixed(1),
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

/** 重なり順・グループを扱う API（PowerPointApi 1.8）が使えるか */
function canArrange(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
}

async function zOrderOf(slideId: string, shapeId: string): Promise<number> {
  return PowerPoint.run(async (context) => {
    const shape = context.presentation.slides.getItem(slideId).shapes.getItem(shapeId);
    shape.load("zOrderPosition");
    await context.sync();
    return shape.zOrderPosition;
  });
}

/** 図形を、重なり順が position になるまで背面へ移す。挿入した図形は最前面に来るため */
async function sendBackTo(slideId: string, shapeId: string, position: number): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id");
    const shape = shapes.getItem(shapeId);
    shape.load("zOrderPosition");
    await context.sync();
    if (position <= 0) {
      shape.setZOrder("SendToBack");
      await context.sync();
      return;
    }
    // 1 回で 1 つ下がる。図形の数より多くは回さない
    for (let i = 0; i < shapes.items.length && shape.zOrderPosition > position; i++) {
      shape.setZOrder("SendBackward");
      shape.load("zOrderPosition");
      await context.sync();
    }
  });
}

interface GroupInfo {
  name: string;
  zOrder: number;
  /** 図以外にグループに入っていた図形 */
  others: string[];
}

/** グループを解除し、作り直したあとに組み直すための情報を返す */
async function ungroup(slideId: string, groupId: string, diagramId: string): Promise<GroupInfo> {
  return PowerPoint.run(async (context) => {
    const group = context.presentation.slides.getItem(slideId).shapes.getItem(groupId);
    group.load("name,zOrderPosition");
    const members = group.group.shapes;
    members.load("items/id");
    await context.sync();
    const info = {
      name: group.name,
      zOrder: group.zOrderPosition,
      others: members.items.map((s) => s.id).filter((id) => id !== diagramId),
    };
    group.group.ungroup();
    await context.sync();
    return info;
  });
}

async function regroup(slideId: string, shapeIds: string[], info: GroupInfo): Promise<void> {
  const groupId = await PowerPoint.run(async (context) => {
    const group = context.presentation.slides.getItem(slideId).shapes.addGroup(shapeIds);
    group.name = info.name;
    group.load("id");
    await context.sync();
    return group.id;
  });
  await sendBackTo(slideId, groupId, info.zOrder);
}

/**
 * 既存の図を作り直した図に置き換える。重ねた図形が隠れないよう、前の図と同じ重なり順に戻す。
 * グループの中の図は、グループを一度解除して置き換え、組み直す。
 */
export async function replaceDiagram(req: WriteRequest, old: ExistingDiagram): Promise<string> {
  if (!canArrange()) {
    const id = await insertDiagram(req);
    await deleteShape(req.slideId, old.shapeId);
    return id;
  }
  const group = old.groupId ? await ungroup(req.slideId, old.groupId, old.shapeId) : null;
  let newId: string | null = null;
  try {
    const zOrder = await zOrderOf(req.slideId, old.shapeId);
    newId = await insertDiagram(req);
    await deleteShape(req.slideId, old.shapeId);
    await sendBackTo(req.slideId, newId, zOrder);
    return newId;
  } finally {
    // 置き換えに失敗しても、グループは元の図で組み直す
    if (group) await regroup(req.slideId, [...group.others, newId ?? old.shapeId], group);
  }
}

export async function selectSlide(slideId: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
  });
}
