import PptxGenJS from "pptxgenjs";

export type EngineerData = Record<string, string>;

function first3(data: EngineerData, prefix: string): string[] {
  return [1, 2, 3].map((i) => data[`${prefix}${i}`] ?? "").filter(Boolean);
}

function getProjects(data: EngineerData) {
  return [1, 2, 3].map((i) => ({
    overview: data[`案件${i}：概要`] ?? "",
    stack:    data[`案件${i}：技術スタック`] ?? "",
    role:     data[`案件${i}：役割`] ?? "",
    period:   data[`案件${i}：期間`] ?? "",
    scale:    data[`案件${i}：規模`] ?? "",
    result:   data[`案件${i}：成果`] ?? "",
    extra:    Object.entries(data)
      .filter(([k, v]) => k.startsWith(`案件${i}：`) && v)
      .map(([k, v]) => ({ label: k.slice(`案件${i}：`.length), value: v }))
      .filter(({ label }) => !["概要","技術スタック","役割","期間","規模","成果"].includes(label)),
  })).filter((p) => p.overview);
}

const C = {
  navy:         "0F2744",
  blue:         "0057B8",
  bluePale:     "E8F1FB",
  blueBorder:   "C0D8F5",
  orange:       "D97706",
  orangePale:   "FEF3C7",
  orangeBorder: "F5D78A",
  green:        "065F46",
  greenPale:    "D1FAE5",
  greenBorder:  "99D8C0",
  slate:        "4A5F78",
  muted:        "7A90A8",
  border:       "D4DDE8",
  bg:           "F5F8FC",
  white:        "FFFFFF",
};

// スライドサイズ: 13.33 × 7.5 in (LAYOUT_WIDE)
const LW  = 2.6;  // 左パネル幅
const RX  = 2.75; // 右エリア開始X
const RW  = 10.4; // 右エリア幅
const SW  = 13.33;
const SH  = 7.5;

/**
 * テキストを高さ・幅・フォントサイズに基づいて切り詰める。
 * pptxgenjsは<a:normAutofit/>の自動縮小を確実に適用しないため、
 * JS側で切り詰めることで重なりを防ぐ。
 * @param text 元テキスト
 * @param wIn  テキストボックス幅（inch）
 * @param hIn  テキストボックス高さ（inch）
 * @param fontSize フォントサイズ（pt）
 * @param lineSpacing 行間倍率（デフォルト1.3）
 * @param isBold 太字かどうか（太字は文字幅が若干広い）
 */
function fitText(text: string, wIn: number, hIn: number, fontSize: number, lineSpacing = 1.3, isBold = false): string {
  // 1文字あたりの幅（inch）: 日本語/太字は広め
  const charW = (fontSize * (isBold ? 0.65 : 0.58)) / 72;
  // 1行あたりの高さ（inch）
  const lineH = (fontSize / 72) * lineSpacing * 1.15; // 1.15はpptxgenjs内部余白補正
  const charsPerLine = Math.max(1, Math.floor(wIn / charW));
  const maxLines     = Math.max(1, Math.floor(hIn / lineH));
  const maxChars     = charsPerLine * maxLines;
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars - 1) + "…";
}

export function generatePptx(data: EngineerData): void {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  const slide = pptx.addSlide();

  // 背景
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SW, h:SH, fill:{color:C.bg}, line:{color:C.bg} });

  /* ════ 左パネル ════ */
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:LW, h:SH, fill:{color:C.navy}, line:{color:C.navy} });

  slide.addText("ATWARE ENGINEER PROFILE", {
    x:0.22, y:0.22, w:LW-0.3, h:0.2,
    fontSize:8, bold:true, color:"5A8AAE", charSpacing:2,
  });
  slide.addText(data["氏名"] ?? "", {
    x:0.22, y:0.46, w:LW-0.3, h:0.68,
    fontSize:26, bold:true, color:C.white, fontFace:"Meiryo UI",
  });
  slide.addText(`"${data["キャッチフレーズ"] ?? ""}"`, {
    x:0.22, y:1.2, w:LW-0.3, h:0.6,
    fontSize:11, italic:true, color:"9EC8F0", fontFace:"Meiryo UI", wrap:true, lineSpacingMultiple:1.3,
  });

  // 要約
  const summary = data["__summary"] ?? "";
  if (summary) {
    const sumH = 0.9;
    slide.addText(fitText(summary, LW-0.34, sumH, 11, 1.4), {
      x:0.22, y:1.88, w:LW-0.3, h:sumH,
      fontSize:11, color:"CCDDE8", fontFace:"Meiryo UI", wrap:true, lineSpacingMultiple:1.4,
    });
  }

  const leftFields: [string, string][] = [
    ["エンジニアタイプ", data["エンジニアタイプ"] ?? ""],
    ["主戦場スキル",     data["主戦場スキル"] ?? ""],
    ["担ってきた役割",   data["担ってきた役割"] ?? ""],
  ].filter(([, v]) => v) as [string, string][];

  // 左パネルフィールド: コンパクト積み上げ（均等分割しない）
  const fieldStart = summary ? 2.9 : 1.92;
  const fieldEnd   = SH - 0.44;
  const LABEL_H    = 0.22;
  const LBL_GAP    = 0.05;   // ラベル→値の隙間
  const FIELD_GAP  = 0.2;    // フィールド間の余白
  const LINE_H10   = (10 / 72) * 1.3 * 1.18; // 10pt での1行高さ(inch)
  const FW         = LW - 0.34;

  let curY = fieldStart;
  leftFields.forEach(([label, val]) => {
    if (curY + LABEL_H + 0.2 > fieldEnd) return; // スペース不足はスキップ
    // 実際の行数を推定してvalHを決定
    const charW = (10 * 0.55) / 72;
    const cpl   = Math.max(1, Math.floor(FW / charW));
    const lines = Math.min(Math.ceil(val.length / cpl), 4);
    const valH  = Math.min(lines * LINE_H10 + 0.04, fieldEnd - curY - LABEL_H - LBL_GAP - FIELD_GAP);
    if (valH < 0.15) return;

    slide.addText(label, { x:0.22, y:curY, w:LW-0.3, h:LABEL_H, fontSize:9, bold:true, color:"7AAED4", charSpacing:1.5 });
    slide.addText(fitText(val, FW, valH, 10, 1.3), {
      x:0.22, y:curY + LABEL_H + LBL_GAP, w:LW-0.3, h:valH,
      fontSize:10, color:"D4E8F8", fontFace:"Meiryo UI", wrap:true, lineSpacingMultiple:1.3,
    });
    curY += LABEL_H + LBL_GAP + valH + FIELD_GAP;
  });

  // フッター
  slide.addShape(pptx.ShapeType.rect, { x:0, y:SH-0.36, w:LW, h:0.36, fill:{color:"0A1E38"}, line:{color:"0A1E38"} });
  slide.addText("atWare", { x:0.22, y:SH-0.30, w:LW-0.3, h:0.24, fontSize:10, bold:true, color:"4A7AAA" });

  /* ════ 右エリア ════ */
  const values   = first3(data, "価値の出し方");
  const kodawari = first3(data, "こだわり");
  const projs    = getProjects(data);
  const month    = new Date().toLocaleDateString("ja-JP", { year:"numeric", month:"long" });

  // ① 価値の出し方（横3枚）
  slide.addText("価値の出し方", { x:RX, y:0.18, w:RW, h:0.24, fontSize:10, bold:true, color:C.blue, charSpacing:2 });
  const vCW = (RW - 0.16) / 3;
  for (let i = 0; i < 3; i++) {
    const cx = RX + i * (vCW + 0.08);
    slide.addShape(pptx.ShapeType.rect, { x:cx, y:0.46, w:vCW, h:0.9, fill:{color:C.bluePale}, line:{color:C.blueBorder, width:0.75}, rectRadius:0.05 });
    slide.addShape(pptx.ShapeType.ellipse, { x:cx+0.1, y:0.55, w:0.22, h:0.22, fill:{color:C.blue}, line:{color:C.blue} });
    slide.addText(String(i+1), { x:cx+0.1, y:0.55, w:0.22, h:0.22, fontSize:10, bold:true, color:C.white, align:"center", valign:"middle" });
    slide.addText(values[i] ?? "—", { x:cx+0.36, y:0.54, w:vCW-0.46, h:0.74, fontSize:12, color:C.navy, fontFace:"Meiryo UI", wrap:true, valign:"top", lineSpacingMultiple:1.3 });
  }

  // ② 代表案件（縦並び）
  // こだわりを固定位置に置くため、代表案件エリアの高さからprojHを逆算
  const KOD_LABEL_Y = 5.94;   // こだわりラベルのY位置（固定）
  const KOD_H       = 1.0;    // こだわりカードの高さ（固定）
  const PROJ_TOP    = 1.66;   // 代表案件カード開始Y
  const PROJ_AREA_H = KOD_LABEL_Y - 0.1 - PROJ_TOP; // = 4.24 in
  const numProjs    = Math.min(projs.length, 3) || 1;
  const projH       = (PROJ_AREA_H - Math.max(0, numProjs - 1) * 0.08) / numProjs;

  slide.addText("代表案件", { x:RX, y:1.4, w:RW, h:0.24, fontSize:10, bold:true, color:C.slate, charSpacing:2 });

  const projW   = RW - 2.0;
  const stackX  = RX + projW + 0.12;
  const stackW  = 1.88;

  projs.slice(0, 3).forEach((p, i) => {
    const py = PROJ_TOP + i * (projH + 0.1);

    slide.addShape(pptx.ShapeType.rect, { x:RX, y:py, w:RW, h:projH, fill:{color:C.white}, line:{color:C.border, width:0.75}, rectRadius:0.05 });

    // 案件バッジ — fontFace指定で漢字と数字を同一フォントに統一
    slide.addShape(pptx.ShapeType.rect, { x:RX+0.1, y:py+0.1, w:0.68, h:0.28, fill:{color:C.bluePale}, line:{color:C.blueBorder, width:0.75}, rectRadius:0.03 });
    slide.addText(`案件 ${i+1}`, { x:RX+0.1, y:py+0.1, w:0.68, h:0.28, fontSize:9, bold:true, color:C.blue, fontFace:"Meiryo UI", align:"center", valign:"middle" });

    // カード内の配置 — margin:0 でOOXMLデフォルト内部余白(0.064in×2)を除去
    // これにより y座標がそのままテキスト先頭になる
    const MARGIN = 0;
    const titleCharsPerLine = Math.max(1, Math.floor((projW - 0.84) / (13 / 72)));
    const titleLines = Math.min(Math.ceil(p.overview.length / titleCharsPerLine), 3);
    const TITLE_H = titleLines * (13 / 72) * 1.2 + 0.02;
    const RES_H   = p.result ? 0.30 : 0;
    const meta    = [p.role && `役割: ${p.role}`, (p.period||p.scale) && [p.period,p.scale].filter(Boolean).join(" | ")].filter(Boolean).join("　");
    const extras  = p.extra.slice(0, 2);

    const titleY   = py + 0.08;
    const detailY  = titleY + TITLE_H + 0.04;
    const detailH  = projH - (detailY - py) - RES_H - 0.04;

    // 概要（タイトル）
    slide.addText(fitText(p.overview, projW-0.84, TITLE_H, 13, 1.2, true), {
      x:RX+0.84, y:titleY, w:projW-0.84, h:TITLE_H,
      fontSize:13, bold:true, color:C.navy, fontFace:"Meiryo UI", wrap:true, lineSpacingMultiple:1.2,
      valign:"top", margin:MARGIN,
    });

    // 詳細テキスト（meta + extra を1ボックスに統合 — 分割すると anchor:ctr で余白が生まれる）
    const detailLines = [
      ...(meta ? [meta] : []),
      ...extras.map(({ label, value }) => `${label}: ${value}`),
    ];
    if (detailLines.length > 0) {
      const detailText = detailLines.join("\n");
      slide.addText(fitText(detailText, projW-0.84, detailH, 10, 1.4), {
        x:RX+0.84, y:detailY, w:projW-0.84, h:detailH,
        fontSize:10, color:C.slate, fontFace:"Meiryo UI", wrap:true, lineSpacingMultiple:1.4,
        valign:"top", margin:MARGIN,
      });
    }

    // 成果（カード下部に固定）
    if (p.result) {
      const ry = py + projH - RES_H;
      slide.addShape(pptx.ShapeType.rect, { x:RX+0.68, y:ry, w:projW-0.74, h:0.26, fill:{color:C.greenPale}, line:{color:C.greenBorder, width:0.75}, rectRadius:0.03 });
      slide.addText(fitText(`成果: ${p.result}`, projW-0.96, 0.24, 10), {
        x:RX+0.78, y:ry+0.03, w:projW-0.96, h:0.24, fontSize:10, color:C.green, fontFace:"Meiryo UI",
      });
    }

    // 技術スタック（右側）
    if (p.stack) {
      slide.addText("技術スタック", { x:stackX, y:py+0.1, w:stackW, h:0.2, fontSize:9, bold:true, color:C.muted });
      const tags = p.stack.split(/[,、]/).map((s: string) => s.trim()).filter(Boolean);
      let tx = stackX, ty = py+0.32;
      for (const tag of tags) {
        const tw = Math.min(tag.length * 0.13 + 0.22, stackW);
        if (tx + tw > stackX + stackW) { tx = stackX; ty += 0.28; }
        if (ty + 0.24 > py + projH - 0.06) break;
        slide.addShape(pptx.ShapeType.rect, { x:tx, y:ty, w:tw, h:0.24, fill:{color:C.bluePale}, line:{color:C.blueBorder, width:0.5}, rectRadius:0.03 });
        slide.addText(tag, { x:tx, y:ty, w:tw, h:0.24, fontSize:10, bold:true, color:C.blue, align:"center", valign:"middle" });
        tx += tw + 0.06;
      }
    }
  });

  // ③ こだわり（横3枚、固定位置・固定高さ）
  slide.addText("こだわり", { x:RX, y:KOD_LABEL_Y, w:RW, h:0.24, fontSize:10, bold:true, color:C.orange, charSpacing:2 });
  const kCW = (RW - 0.16) / 3;
  for (let i = 0; i < 3; i++) {
    const cx = RX + i * (kCW + 0.08);
    slide.addShape(pptx.ShapeType.rect, { x:cx, y:KOD_LABEL_Y+0.28, w:kCW, h:KOD_H, fill:{color:C.orangePale}, line:{color:C.orangeBorder, width:0.75}, rectRadius:0.05 });
    slide.addShape(pptx.ShapeType.ellipse, { x:cx+0.1, y:KOD_LABEL_Y+0.36, w:0.22, h:0.22, fill:{color:C.orange}, line:{color:C.orange} });
    slide.addText(String(i+1), { x:cx+0.1, y:KOD_LABEL_Y+0.36, w:0.22, h:0.22, fontSize:10, bold:true, color:C.white, align:"center", valign:"middle" });
    slide.addText(kodawari[i] ?? "—", { x:cx+0.36, y:KOD_LABEL_Y+0.36, w:kCW-0.46, h:KOD_H-0.1, fontSize:12, color:C.navy, fontFace:"Meiryo UI", wrap:true, valign:"top", lineSpacingMultiple:1.3 });
  }

  // フッター
  slide.addShape(pptx.ShapeType.line, { x:RX, y:SH-0.3, w:RW, h:0, line:{color:C.border, width:0.75} });
  slide.addText("Engineer Profile  |  Confidential", { x:RX, y:SH-0.26, w:5, h:0.22, fontSize:9, color:C.muted });
  slide.addText(month, { x:SW-2.2, y:SH-0.26, w:2.0, h:0.22, fontSize:9, color:C.muted, align:"right" });

  const safeName = (data["氏名"] ?? "profile").replace(/[^\w　-鿿゠-ヿ぀-ゟ\-]/g, "_");
  pptx.writeFile({ fileName: `${safeName}_profile.pptx` });
}

// カードテーマ（EngineerCard.tsx と同値）
const CARD_THEMES_PPTX = [
  { id: "white",  avatarBg: "F2F0EB", avatarText: "A89F90", cardBg: "FFFFFF",  border: "E4E0D8" },
  { id: "slate",  avatarBg: "DDE4ED", avatarText: "8A9DB5", cardBg: "F8FAFD",  border: "CDD7E4" },
  { id: "sage",   avatarBg: "DDE8E1", avatarText: "8AAD95", cardBg: "F7FBF8",  border: "C8DDD0" },
  { id: "rose",   avatarBg: "EDE0E0", avatarText: "B59090", cardBg: "FDF8F8",  border: "E0CCCC" },
  { id: "indigo", avatarBg: "E0E0ED", avatarText: "9090B5", cardBg: "F8F8FD",  border: "CCCCE0" },
] as const;

export function generateCardPptx(
  fields: { name: string; engineerType: string; catchphrase: string; summary: string; skills: string },
  themeId: string = "white"
) {
  const theme = CARD_THEMES_PPTX.find((t) => t.id === themeId) ?? CARD_THEMES_PPTX[0];

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.33 × 7.5 in

  // カードサイズ・位置
  const CW = 3.2;   // カード幅
  const CH = 5.0;   // カード高さ
  const CX = (SW - CW) / 2;
  const CY = (SH - CH) / 2;
  const AVH = 1.1;  // アバターエリア高さ
  const R = 0.08;   // 角丸

  const slide = pptx.addSlide();
  slide.background = { color: C.bg };

  // カード外枠
  slide.addShape(pptx.ShapeType.rect, {
    x: CX, y: CY, w: CW, h: CH,
    fill: { color: theme.cardBg },
    line: { color: theme.border, width: 1 },
    rectRadius: R,
  });

  // アバターエリア背景
  slide.addShape(pptx.ShapeType.rect, {
    x: CX, y: CY, w: CW, h: AVH,
    fill: { color: theme.avatarBg },
    line: { color: theme.border, width: 0 },
    rectRadius: R,
  });
  // アバターエリア下部を角丸なしで上書き（下半分を覆う）
  slide.addShape(pptx.ShapeType.rect, {
    x: CX, y: CY + AVH * 0.5, w: CW, h: AVH * 0.5,
    fill: { color: theme.avatarBg },
    line: { color: "FFFFFF", width: 0 },
  });
  // アバターエリア下境界線
  slide.addShape(pptx.ShapeType.line, {
    x: CX, y: CY + AVH, w: CW, h: 0,
    line: { color: theme.border, width: 0.75 },
  });

  // イニシャル円
  const initials = fields.name.split(/\s+/).map((s) => s[0]).join("").slice(0, 2) || "　";
  const circleSize = 0.64;
  const circleX = CX + (CW - circleSize) / 2;
  const circleY = CY + (AVH - circleSize) / 2;
  slide.addShape(pptx.ShapeType.ellipse, {
    x: circleX, y: circleY, w: circleSize, h: circleSize,
    fill: { color: "FFFFFF", transparency: 40 },
    line: { color: theme.border, width: 0 },
  });
  slide.addText(initials, {
    x: circleX, y: circleY, w: circleSize, h: circleSize,
    fontSize: 18, bold: true, color: theme.avatarText,
    align: "center", valign: "middle", fontFace: "Meiryo UI",
  });

  // 本文エリア
  const BX = CX + 0.18;
  const BW = CW - 0.36;
  let BY = CY + AVH + 0.16;

  // エンジニアタイプ
  if (fields.engineerType) {
    const typeText = fields.engineerType.split(/[,、]/).map((s) => s.trim()).filter(Boolean).join(" / ");
    const typeCharW = (10 * 0.58) / 72;
    const typeCPL = Math.max(1, Math.floor(BW / typeCharW));
    const typeLines = Math.min(Math.ceil(typeText.length / typeCPL), 3);
    const typeLineH = (10 / 72) * 1.3 * 1.15;
    const typeH = typeLines * typeLineH + 0.04;
    slide.addText(fitText(typeText, BW, typeH, 10, 1.3), {
      x: BX, y: BY, w: BW, h: typeH,
      fontSize: 10, color: C.muted, fontFace: "Meiryo UI", wrap: true,
      valign: "top", margin: 0,
    });
    BY += typeH + 0.06;
  }

  // キャッチフレーズ
  if (fields.catchphrase) {
    const cpLines = Math.min(Math.ceil(fields.catchphrase.length / Math.floor(BW / (10.5 / 72 * 0.6))), 3);
    const cpH = cpLines * (10.5 / 72) * 1.5 + 0.04;
    slide.addText(fitText(fields.catchphrase, BW, cpH, 10.5, 1.5), {
      x: BX, y: BY, w: BW, h: cpH,
      fontSize: 10.5, color: "666666", fontFace: "Meiryo UI", wrap: true, lineSpacingMultiple: 1.5,
      valign: "top", margin: 0,
    });
    BY += cpH + 0.1;
  }

  // 区切り線
  slide.addShape(pptx.ShapeType.line, {
    x: BX, y: BY, w: BW, h: 0,
    line: { color: theme.border, width: 0.75 },
  });
  BY += 0.12;

  // 紹介文（実際の行数で高さを計算し、空白を残さない）
  if (fields.summary) {
    const sumCharW = (10.5 * 0.58) / 72;
    const sumCPL = Math.max(1, Math.floor(BW / sumCharW));
    const sumLineH = (10.5 / 72) * 1.75 * 1.15;
    const sumMaxBottom = CY + CH - (fields.skills ? 0.9 : 0.18);
    const sumMaxH = sumMaxBottom - BY;
    const sumMaxLines = Math.max(1, Math.floor(sumMaxH / sumLineH));
    const sumLines = Math.min(Math.ceil(fields.summary.length / sumCPL), sumMaxLines);
    const sumH = Math.min(sumLines * sumLineH + 0.06, sumMaxH);
    slide.addText(fitText(fields.summary, BW, sumH, 10.5, 1.75), {
      x: BX, y: BY, w: BW, h: sumH,
      fontSize: 10.5, color: "444444", fontFace: "Meiryo UI", wrap: true, lineSpacingMultiple: 1.75,
      valign: "top", margin: 0,
    });
    BY += sumH + 0.1;
  }

  // スキル
  if (fields.skills) {
    slide.addShape(pptx.ShapeType.line, {
      x: BX, y: BY, w: BW, h: 0,
      line: { color: theme.border, width: 0.75 },
    });
    BY += 0.1;
    slide.addText("スキル", {
      x: BX, y: BY, w: BW, h: 0.16,
      fontSize: 9, bold: true, color: C.muted, fontFace: "Meiryo UI", charSpacing: 1,
    });
    BY += 0.2;
    const skillH = Math.max(0.2, CY + CH - BY - 0.1);
    slide.addText(fitText(fields.skills, BW, skillH, 10, 1.6), {
      x: BX, y: BY, w: BW, h: skillH,
      fontSize: 10, color: "555555", fontFace: "Meiryo UI", wrap: true, lineSpacingMultiple: 1.6,
      valign: "top", margin: 0,
    });
  }

  const safeName = (fields.name || "card").replace(/[^\w　-鿿゠-ヿ぀-ゟ\-]/g, "_");
  pptx.writeFile({ fileName: `${safeName}_card.pptx` });
}
