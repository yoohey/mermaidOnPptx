/**
 * ノートに貼るための、主な図の最小構成サンプル。パネルからコピーして使う。
 */

export interface Sample {
  id: string;
  label: string;
  code: string;
}

export const SAMPLES: Sample[] = [
  {
    id: "sequence",
    label: "シーケンス図",
    code: `sequenceDiagram
  participant U as ユーザー
  participant A as アプリ
  U->>A: ログイン
  A-->>U: OK`,
  },
  {
    id: "flowchart",
    label: "フローチャート",
    code: `flowchart LR
  A[開始] --> B{条件}
  B -- はい --> C[処理]
  B -- いいえ --> D[終了]
  C --> D`,
  },
  {
    id: "state",
    label: "状態遷移図",
    code: `stateDiagram-v2
  [*] --> 待機
  待機 --> 実行中: 開始
  実行中 --> 待機: 完了
  実行中 --> [*]: 中止`,
  },
  {
    id: "class",
    label: "クラス図",
    code: `classDiagram
  class 注文 {
    +番号
    +合計()
  }
  class 明細 {
    +数量
  }
  注文 "1" --> "*" 明細`,
  },
  {
    id: "er",
    label: "ER図",
    code: `erDiagram
  顧客 ||--o{ 注文 : する
  注文 ||--|{ 明細 : 含む`,
  },
  {
    id: "gantt",
    label: "ガントチャート",
    code: `gantt
  dateFormat YYYY-MM-DD
  axisFormat %m/%d
  section 準備
  設計 :a1, 2026-11-01, 5d
  section 開発
  実装 :after a1, 7d`,
  },
  {
    id: "pie",
    label: "円グラフ",
    code: `pie
  "A" : 50
  "B" : 30
  "C" : 20`,
  },
  {
    id: "mindmap",
    label: "マインドマップ",
    code: `mindmap
  root((テーマ))
    案1
      詳細
    案2`,
  },
];

/** ノートにそのまま貼れる形（```mermaid の囲み付き） */
export function fencedSample(sample: Sample): string {
  return "```mermaid\n" + sample.code + "\n```";
}
