/**
 * スライド上の、このアドインが管理する図（タグ M4P_ID を持つ図形）を読む。
 */
import { TAG } from "../core/tags";
import type { ExistingDiagram } from "../core/sync";
import { isResized, parseBox } from "../core/layout";

type ShapeTags = Record<"id" | "hash" | "theme" | "box" | "fontPt", PowerPoint.Tag>;

function loadTags(shape: PowerPoint.Shape): ShapeTags {
  const tags = {
    id: shape.tags.getItemOrNullObject(TAG.ID),
    hash: shape.tags.getItemOrNullObject(TAG.HASH),
    theme: shape.tags.getItemOrNullObject(TAG.THEME),
    box: shape.tags.getItemOrNullObject(TAG.BOX),
    fontPt: shape.tags.getItemOrNullObject(TAG.FONT_PT),
  };
  Object.values(tags).forEach((t) => t.load("value"));
  return tags;
}

const valueOf = (tag: PowerPoint.Tag) => (tag.isNullObject ? "" : tag.value);

/**
 * スライド直下の図と、スライド直下のグループの中の図を読む（グループの中のグループは見ない）。
 * グループの図形を読む API（PowerPointApi 1.8）が無ければ、スライド直下だけを読む
 */
export async function readDiagrams(slideId: string): Promise<ExistingDiagram[]> {
  const withGroups = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load(`items/id,items/left,items/top,items/width,items/height${withGroups ? ",items/type" : ""}`);
    await context.sync();

    const found: { shape: PowerPoint.Shape; tags: ShapeTags; groupId?: string }[] = [];
    const groups: { groupId: string; members: PowerPoint.ShapeScopedCollection }[] = [];
    for (const shape of shapes.items) {
      found.push({ shape, tags: loadTags(shape) });
      if (withGroups && shape.type === "Group") {
        const members = shape.group.shapes;
        members.load("items/id,items/left,items/top,items/width,items/height");
        groups.push({ groupId: shape.id, members });
      }
    }
    await context.sync();
    for (const { groupId, members } of groups) {
      for (const shape of members.items) found.push({ shape, tags: loadTags(shape), groupId });
    }
    await context.sync();

    return found
      .filter(({ tags }) => !tags.id.isNullObject)
      .map(({ shape, tags, groupId }) => {
        const current = { width: shape.width, height: shape.height };
        return {
          shapeId: shape.id,
          blockId: tags.id.value,
          hash: valueOf(tags.hash),
          theme: valueOf(tags.theme),
          resized: isResized(parseBox(valueOf(tags.box)), current),
          fontPt: Number(valueOf(tags.fontPt)) || null,
          groupId,
          left: shape.left,
          top: shape.top,
          ...current,
        };
      });
  });
}

/**
 * プレゼン内で最初に見つかった図のタグ M4P_THEME（図を作ったときのテーマ）を返す。図が無ければ null。
 * 表示中のスライドを先に調べ、無ければ先頭のスライドから順に調べる。
 */
export async function findDiagramThemeKey(currentSlideId?: string): Promise<string | null> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const ordered = [...slides.items].sort((a, b) => Number(b.id === currentSlideId) - Number(a.id === currentSlideId));

    const perSlide = ordered.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/id");
      return shapes;
    });
    await context.sync();

    const tags = perSlide.flatMap((shapes) =>
      shapes.items.map((shape) => {
        const id = shape.tags.getItemOrNullObject(TAG.ID);
        const theme = shape.tags.getItemOrNullObject(TAG.THEME);
        id.load("value");
        theme.load("value");
        return { id, theme };
      }),
    );
    await context.sync();

    const found = tags.find((t) => !t.id.isNullObject);
    if (!found) return null;
    return found.theme.isNullObject ? "" : found.theme.value;
  });
}
