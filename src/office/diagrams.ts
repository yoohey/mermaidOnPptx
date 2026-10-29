/**
 * スライド上の、このアドインが管理する図（タグ M4P_ID を持つ図形）を読む。
 */
import { TAG } from "../core/tags";
import type { ExistingDiagram } from "../core/sync";

export async function readDiagrams(slideId: string): Promise<ExistingDiagram[]> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(slideId).shapes;
    shapes.load("items/id,items/left,items/top,items/width,items/height");
    await context.sync();

    const withTags = shapes.items.map((shape) => {
      const tags = {
        id: shape.tags.getItemOrNullObject(TAG.ID),
        hash: shape.tags.getItemOrNullObject(TAG.HASH),
        theme: shape.tags.getItemOrNullObject(TAG.THEME),
      };
      Object.values(tags).forEach((t) => t.load("value"));
      return { shape, tags };
    });
    await context.sync();

    return withTags
      .filter(({ tags }) => !tags.id.isNullObject)
      .map(({ shape, tags }) => ({
        shapeId: shape.id,
        blockId: tags.id.value,
        hash: tags.hash.isNullObject ? "" : tags.hash.value,
        theme: tags.theme.isNullObject ? "" : tags.theme.value,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
      }));
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
