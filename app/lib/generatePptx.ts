import PptxGenJS from "pptxgenjs";

export type EngineerData = Record<string, string>;

function first3(data: EngineerData, prefix: string): string[] {
  return [1, 2, 3]
    .map((i) => data[`${prefix}${i}`] ?? "")
    .filter(Boolean);
}

// カラーパレット
const C = {
  navy:       "0F2744",
  navyMid:    "1C3D6B",
  blue:       "0057B8",
  bluePale:   "E8F1FB",
  blueMid:    "C0D8F5",
  orange:     "D97706",
  orangePale: "FEF3C7",
  purple:     "6D28D9",
  purplePale: "EDE9FE",
  green:      "065F46",
  greenPale:  "D1FAE5",
  slate:      "4A5F78",
  muted:      "7A90A8",
  border:     "D4DDE8",
  bg:         "F5F8FC",
  white:      "FFFFFF",
};

// 幅 13.33 in, 高さ 7.5 in (LAYOUT_WIDE)
// 左パネル幅 3.6in, 右エリア 9.4in (from x=3.8)
const LW = 3.6;
const RX = 3.85;
const RW = 9.1;
const SW = 13.33; // slide width

export function generatePptx(data: EngineerData): void {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  const slide = pptx.addSlide();

  /* ──────────────────────────────────────────
   * 背景
   * ────────────────────────────────────────── */
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SW, h: 7.5,
    fill: { color: C.bg },
    line: { color: C.bg },
  });

  /* ──────────────────────────────────────────
   * 左パネル（navy グラデーション風）
   * ────────────────────────────────────────── */
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: LW, h: 7.5,
    fill: { color: C.navy },
    line: { color: C.navy },
  });
  // アクセントライン
  slide.addShape(pptx.ShapeType.rect, {
    x: LW - 0.05, y: 0, w: 0.05, h: 7.5,
    fill: { color: C.blue },
    line: { color: C.blue },
  });

  // ラベル
  slide.addText("ATWARE ENGINEER PROFILE", {
    x: 0.3, y: 0.28, w: LW - 0.4, h: 0.2,
    fontSize: 7, bold: true, color: "6A9FCC",
    charSpacing: 2,
  });

  // 氏名
  slide.addText(data["氏名"] ?? "", {
    x: 0.3, y: 0.52, w: LW - 0.4, h: 0.7,
    fontSize: 26, bold: true, color: C.white,
    fontFace: "Meiryo UI",
    charSpacing: -0.5,
  });

  // キャッチフレーズ
  slide.addText(`"${data["キャッチフレーズ"] ?? ""}"`, {
    x: 0.3, y: 1.27, w: LW - 0.4, h: 0.65,
    fontSize: 10, italic: true, color: "9EC8F0",
    fontFace: "Meiryo UI",
    wrap: true,
  });

  // 区切り線
  slide.addShape(pptx.ShapeType.line, {
    x: 0.3, y: 2.0, w: LW - 0.6, h: 0,
    line: { color: "2A4F7A", width: 0.75 },
  });

  // エンジニアタイプ
  slide.addText("エンジニアタイプ", {
    x: 0.3, y: 2.1, w: LW - 0.4, h: 0.22,
    fontSize: 7.5, bold: true, color: "7AAED4", charSpacing: 1.5,
  });
  slide.addText(data["エンジニアタイプ"] ?? "—", {
    x: 0.3, y: 2.35, w: LW - 0.4, h: 0.5,
    fontSize: 10, color: "D4E8F8",
    fontFace: "Meiryo UI", wrap: true,
  });

  // 主戦場スキル
  slide.addText("主戦場スキル", {
    x: 0.3, y: 2.93, w: LW - 0.4, h: 0.22,
    fontSize: 7.5, bold: true, color: "7AAED4", charSpacing: 1.5,
  });
  slide.addText(data["主戦場スキル"] ?? "—", {
    x: 0.3, y: 3.17, w: LW - 0.4, h: 0.7,
    fontSize: 9.5, color: "D4E8F8",
    fontFace: "Meiryo UI", wrap: true,
  });

  // 担ってきた役割
  const roles = data["担ってきた役割"] ?? "";
  if (roles) {
    slide.addText("担ってきた役割", {
      x: 0.3, y: 3.95, w: LW - 0.4, h: 0.22,
      fontSize: 7.5, bold: true, color: "7AAED4", charSpacing: 1.5,
    });
    slide.addText(roles, {
      x: 0.3, y: 4.18, w: LW - 0.4, h: 0.45,
      fontSize: 9.5, color: "D4E8F8",
      fontFace: "Meiryo UI", wrap: true,
    });
  }

  // フッター
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 7.18, w: LW, h: 0.32,
    fill: { color: "0A1E38" },
    line: { color: "0A1E38" },
  });
  slide.addText("atWare", {
    x: 0.3, y: 7.2, w: LW - 0.3, h: 0.24,
    fontSize: 9, bold: true, color: "4A7AAA",
  });

  /* ──────────────────────────────────────────
   * 右エリア — 3段構成
   *   A: 価値の出し方 (top)
   *   B: こだわり (mid)
   *   C: 代表案件 (bottom)
   * ────────────────────────────────────────── */

  // ── A: 価値の出し方 ──
  const CW = (RW - 0.24) / 3; // カラム幅 (3列、gap 0.12)

  slide.addText("価値の出し方", {
    x: RX, y: 0.28, w: RW, h: 0.22,
    fontSize: 8, bold: true, color: C.blue, charSpacing: 2,
  });

  const values = first3(data, "価値の出し方");
  for (let i = 0; i < 3; i++) {
    const cx = RX + i * (CW + 0.12);
    // カード背景
    slide.addShape(pptx.ShapeType.rect, {
      x: cx, y: 0.56, w: CW, h: 1.1,
      fill: { color: C.bluePale },
      line: { color: C.blueMid, width: 0.75 },
      rectRadius: 0.07,
    });
    // 番号バッジ
    slide.addShape(pptx.ShapeType.ellipse, {
      x: cx + 0.1, y: 0.63, w: 0.22, h: 0.22,
      fill: { color: C.blue },
      line: { color: C.blue },
    });
    slide.addText(String(i + 1), {
      x: cx + 0.1, y: 0.63, w: 0.22, h: 0.22,
      fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle",
    });
    slide.addText(values[i] ?? "—", {
      x: cx + 0.1, y: 0.9, w: CW - 0.2, h: 0.7,
      fontSize: 9.5, color: C.navy,
      fontFace: "Meiryo UI", wrap: true, valign: "top",
    });
  }

  // ── B: こだわり ──
  slide.addText("こだわり", {
    x: RX, y: 1.82, w: RW, h: 0.22,
    fontSize: 8, bold: true, color: C.orange, charSpacing: 2,
  });

  const kodawari = first3(data, "こだわり");
  for (let i = 0; i < 3; i++) {
    const cx = RX + i * (CW + 0.12);
    slide.addShape(pptx.ShapeType.rect, {
      x: cx, y: 2.1, w: CW, h: 1.1,
      fill: { color: C.orangePale },
      line: { color: "F5D78A", width: 0.75 },
      rectRadius: 0.07,
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: cx + 0.1, y: 2.17, w: 0.22, h: 0.22,
      fill: { color: C.orange },
      line: { color: C.orange },
    });
    slide.addText(String(i + 1), {
      x: cx + 0.1, y: 2.17, w: 0.22, h: 0.22,
      fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle",
    });
    slide.addText(kodawari[i] ?? "—", {
      x: cx + 0.1, y: 2.44, w: CW - 0.2, h: 0.7,
      fontSize: 9.5, color: C.navy,
      fontFace: "Meiryo UI", wrap: true, valign: "top",
    });
  }

  // 区切り
  slide.addShape(pptx.ShapeType.line, {
    x: RX, y: 3.35, w: RW, h: 0,
    line: { color: C.border, width: 0.75 },
  });

  // ── C: 代表案件 ──
  slide.addText("代表案件", {
    x: RX, y: 3.47, w: RW, h: 0.22,
    fontSize: 8, bold: true, color: C.slate, charSpacing: 2,
  });

  // 案件概要（左2/3）
  const projW = RW * 0.62;
  slide.addShape(pptx.ShapeType.rect, {
    x: RX, y: 3.75, w: projW, h: 1.55,
    fill: { color: C.white },
    line: { color: C.border, width: 0.75 },
    rectRadius: 0.07,
  });
  slide.addText("案件概要", {
    x: RX + 0.15, y: 3.83, w: projW - 0.3, h: 0.2,
    fontSize: 7.5, bold: true, color: C.muted, charSpacing: 1,
  });
  slide.addText(data["案件1：概要"] ?? "—", {
    x: RX + 0.15, y: 4.06, w: projW - 0.3, h: 1.15,
    fontSize: 9.5, color: C.slate,
    fontFace: "Meiryo UI", wrap: true, valign: "top",
  });

  // 技術スタック（右1/3）
  const stackX = RX + projW + 0.15;
  const stackW = RW - projW - 0.15;
  slide.addShape(pptx.ShapeType.rect, {
    x: stackX, y: 3.75, w: stackW, h: 1.55,
    fill: { color: C.white },
    line: { color: C.border, width: 0.75 },
    rectRadius: 0.07,
  });
  slide.addText("技術スタック", {
    x: stackX + 0.15, y: 3.83, w: stackW - 0.3, h: 0.2,
    fontSize: 7.5, bold: true, color: C.muted, charSpacing: 1,
  });
  slide.addText(data["案件1：技術スタック"] ?? "—", {
    x: stackX + 0.15, y: 4.06, w: stackW - 0.3, h: 1.15,
    fontSize: 9.5, color: C.slate,
    fontFace: "Meiryo UI", wrap: true, valign: "top",
  });

  // ── ATW 提供価値（任意） ──
  const atwValue = data["ATWとして提供できる価値 → 価値"] ?? "";
  if (atwValue) {
    slide.addShape(pptx.ShapeType.rect, {
      x: RX, y: 5.45, w: RW, h: 0.62,
      fill: { color: C.purplePale },
      line: { color: "C4B5FD", width: 0.75 },
      rectRadius: 0.07,
    });
    slide.addText("ATWとして提供できる価値", {
      x: RX + 0.15, y: 5.5, w: 2.2, h: 0.22,
      fontSize: 7.5, bold: true, color: C.purple,
    });
    slide.addText(atwValue, {
      x: RX + 2.4, y: 5.48, w: RW - 2.55, h: 0.52,
      fontSize: 10, color: C.purple,
      fontFace: "Meiryo UI", wrap: true, valign: "middle",
    });
  }

  // ── 右フッター ──
  slide.addShape(pptx.ShapeType.rect, {
    x: LW, y: 7.18, w: SW - LW, h: 0.32,
    fill: { color: C.border },
    line: { color: C.border },
  });
  slide.addText("Engineer Profile  |  Confidential", {
    x: RX, y: 7.2, w: 5, h: 0.24,
    fontSize: 7.5, color: C.muted,
  });
  slide.addText(new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" }), {
    x: SW - 2, y: 7.2, w: 1.8, h: 0.24,
    fontSize: 7.5, color: C.muted, align: "right",
  });

  pptx.writeFile({ fileName: `${data["氏名"] ?? "profile"}_profile.pptx` });
}
