/** 正規化済みソースの SHA-256 を16進で返し、先頭16文字を使う（タグ M4P_HASH 用） */
export async function sourceHash(source: string): Promise<string> {
  const bytes = new TextEncoder().encode(source);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("")
    .slice(0, 16);
}
