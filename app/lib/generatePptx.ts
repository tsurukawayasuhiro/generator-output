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

export function generatePptx(data: EngineerData): void {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  const slide = pptx.addSlide();

  // 背景
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SW, h:SH, fill:{color:C.bg}, line:{color:C.bg} });

  /* ════ 左パネル ════ */
  slide.addShape(pptx.ShapeType.rect, { x:0, y:0, w:LW, h:SH, fill:{color:C.navy}, line:{color:C.navy} });

  slide.addText("ATWARE ENGINEER PROFILE", {
    x:0.22, y:0.22, w:LW-0.3, h:0.18,
    fontSize:6.5, bold:true, color:"5A8AAE", charSpacing:2,
  });
  slide.addText(data["氏名"] ?? "", {
    x:0.22, y:0.42, w:LW-0.3, h:0.62,
    fontSize:24, bold:true, color:C.white, fontFace:"Meiryo UI",
  });
  slide.addText(`"${data["キャッチフレーズ"] ?? ""}"`, {
    x:0.22, y:1.1, w:LW-0.3, h:0.55,
    fontSize:9, italic:true, color:"9EC8F0", fontFace:"Meiryo UI", wrap:true,
  });

  // 要約
  const summary = data["__summary"] ?? "";
  if (summary) {
    slide.addText(summary, {
      x:0.22, y:1.72, w:LW-0.3, h:0.85,
      fontSize:8.5, color:"CCDDE8", fontFace:"Meiryo UI", wrap:true, lineSpacingMultiple:1.3,
    });
  }

  const leftFields: [string, string][] = [
    ["エンジニアタイプ", data["エンジニアタイプ"] ?? ""],
    ["主戦場スキル",     data["主戦場スキル"] ?? ""],
    ["担ってきた役割",   data["担ってきた役割"] ?? ""],
  ].filter(([, v]) => v) as [string, string][];

  let ly = summary ? 2.65 : 1.72;
  for (const [label, val] of leftFields) {
    slide.addText(label, { x:0.22, y:ly, w:LW-0.3, h:0.17, fontSize:6.5, bold:true, color:"7AAED4", charSpacing:1.5 });
    ly += 0.18;
    slide.addText(val,   { x:0.22, y:ly, w:LW-0.3, h:0.36, fontSize:9, color:"D4E8F8", fontFace:"Meiryo UI", wrap:true });
    ly += 0.42;
  }

  // フッター
  slide.addShape(pptx.ShapeType.rect, { x:0, y:SH-0.32, w:LW, h:0.32, fill:{color:"0A1E38"}, line:{color:"0A1E38"} });
  slide.addText("atWare", { x:0.22, y:SH-0.28, w:LW-0.3, h:0.22, fontSize:8.5, bold:true, color:"4A7AAA" });

  /* ════ 右エリア ════ */
  const values   = first3(data, "価値の出し方");
  const kodawari = first3(data, "こだわり");
  const projs    = getProjects(data);
  const month    = new Date().toLocaleDateString("ja-JP", { year:"numeric", month:"long" });

  // ① 価値の出し方（横3枚）
  slide.addText("価値の出し方", { x:RX, y:0.18, w:RW, h:0.2, fontSize:8, bold:true, color:C.blue, charSpacing:2 });
  const vCW = (RW - 0.16) / 3;
  for (let i = 0; i < 3; i++) {
    const cx = RX + i * (vCW + 0.08);
    slide.addShape(pptx.ShapeType.rect, { x:cx, y:0.42, w:vCW, h:0.82, fill:{color:C.bluePale}, line:{color:C.blueBorder, width:0.75}, rectRadius:0.05 });
    slide.addShape(pptx.ShapeType.ellipse, { x:cx+0.1, y:0.5, w:0.19, h:0.19, fill:{color:C.blue}, line:{color:C.blue} });
    slide.addText(String(i+1), { x:cx+0.1, y:0.5, w:0.19, h:0.19, fontSize:8, bold:true, color:C.white, align:"center", valign:"middle" });
    slide.addText(values[i] ?? "—", { x:cx+0.32, y:0.5, w:vCW-0.42, h:0.68, fontSize:9.5, color:C.navy, fontFace:"Meiryo UI", wrap:true, valign:"top" });
  }

  // ② 代表案件（縦並び）
  // こだわりを固定位置に置くため、代表案件エリアの高さからprojHを逆算
  const KOD_LABEL_Y = 5.94;   // こだわりラベルのY位置（固定）
  const KOD_H       = 0.9;    // こだわりカードの高さ（固定）
  const PROJ_TOP    = 1.6;    // 代表案件カード開始Y
  const PROJ_AREA_H = KOD_LABEL_Y - 0.1 - PROJ_TOP; // = 4.24 in
  const numProjs    = Math.min(projs.length, 3) || 1;
  const projH       = (PROJ_AREA_H - Math.max(0, numProjs - 1) * 0.08) / numProjs;

  slide.addText("代表案件", { x:RX, y:1.34, w:RW, h:0.2, fontSize:8, bold:true, color:C.slate, charSpacing:2 });

  const projW   = RW - 1.7;
  const stackX  = RX + projW + 0.1;
  const stackW  = 1.6;

  projs.slice(0, 3).forEach((p, i) => {
    const py = PROJ_TOP + i * (projH + 0.08);

    // カード背景
    slide.addShape(pptx.ShapeType.rect, { x:RX, y:py, w:RW, h:projH, fill:{color:C.white}, line:{color:C.border, width:0.75}, rectRadius:0.05 });

    // 案件バッジ
    slide.addShape(pptx.ShapeType.rect, { x:RX+0.1, y:py+0.1, w:0.38, h:0.18, fill:{color:C.bluePale}, line:{color:C.blueBorder, width:0.75}, rectRadius:0.03 });
    slide.addText(`案件 ${i+1}`, { x:RX+0.1, y:py+0.1, w:0.38, h:0.18, fontSize:7, bold:true, color:C.blue, align:"center", valign:"middle" });

    // 概要
    slide.addText(p.overview, { x:RX+0.55, y:py+0.08, w:projW-0.55, h:0.36, fontSize:10, bold:true, color:C.navy, fontFace:"Meiryo UI", wrap:true });

    // 役割・期間・規模
    const meta = [p.role && `役割: ${p.role}`, (p.period||p.scale) && [p.period,p.scale].filter(Boolean).join(" | ")].filter(Boolean).join("　");
    if (meta) slide.addText(meta, { x:RX+0.55, y:py+0.46, w:projW-0.55, h:0.18, fontSize:8, color:C.slate, fontFace:"Meiryo UI" });

    // extra フィールド
    p.extra.slice(0, 2).forEach(({ label, value }, ei) => {
      slide.addText(`${label}: ${value}`, { x:RX+0.55, y:py+0.46+(ei+1)*0.18, w:projW-0.55, h:0.18, fontSize:8, color:C.slate, fontFace:"Meiryo UI", wrap:true });
    });

    // 成果
    if (p.result) {
      slide.addShape(pptx.ShapeType.rect, { x:RX+0.55, y:py+0.82, w:projW-0.6, h:0.2, fill:{color:C.greenPale}, line:{color:C.greenBorder, width:0.75}, rectRadius:0.03 });
      slide.addText(`成果: ${p.result}`, { x:RX+0.65, y:py+0.83, w:projW-0.8, h:0.18, fontSize:8, color:C.green, fontFace:"Meiryo UI" });
    }

    // 技術スタック（右側）
    if (p.stack) {
      slide.addText("技術スタック", { x:stackX, y:py+0.08, w:stackW, h:0.16, fontSize:7, bold:true, color:C.muted });
      const tags = p.stack.split(/[,、]/).map((s: string) => s.trim()).filter(Boolean);
      let tx = stackX, ty = py+0.26;
      for (const tag of tags) {
        const tw = Math.min(tag.length * 0.09 + 0.18, stackW);
        if (tx + tw > stackX + stackW) { tx = stackX; ty += 0.22; }
        if (ty + 0.2 > py + projH - 0.05) break;
        slide.addShape(pptx.ShapeType.rect, { x:tx, y:ty, w:tw, h:0.18, fill:{color:C.bluePale}, line:{color:C.blueBorder, width:0.5}, rectRadius:0.03 });
        slide.addText(tag, { x:tx, y:ty, w:tw, h:0.18, fontSize:7.5, bold:true, color:C.blue, align:"center", valign:"middle" });
        tx += tw + 0.05;
      }
    }
  });

  // ③ こだわり（横3枚、固定位置・固定高さ）
  slide.addText("こだわり", { x:RX, y:KOD_LABEL_Y, w:RW, h:0.2, fontSize:8, bold:true, color:C.orange, charSpacing:2 });
  const kCW = (RW - 0.16) / 3;
  for (let i = 0; i < 3; i++) {
    const cx = RX + i * (kCW + 0.08);
    slide.addShape(pptx.ShapeType.rect, { x:cx, y:KOD_LABEL_Y+0.24, w:kCW, h:KOD_H, fill:{color:C.orangePale}, line:{color:C.orangeBorder, width:0.75}, rectRadius:0.05 });
    slide.addShape(pptx.ShapeType.ellipse, { x:cx+0.1, y:KOD_LABEL_Y+0.32, w:0.19, h:0.19, fill:{color:C.orange}, line:{color:C.orange} });
    slide.addText(String(i+1), { x:cx+0.1, y:KOD_LABEL_Y+0.32, w:0.19, h:0.19, fontSize:8, bold:true, color:C.white, align:"center", valign:"middle" });
    slide.addText(kodawari[i] ?? "—", { x:cx+0.32, y:KOD_LABEL_Y+0.32, w:kCW-0.42, h:KOD_H-0.1, fontSize:9.5, color:C.navy, fontFace:"Meiryo UI", wrap:true, valign:"top" });
  }

  // フッター
  slide.addShape(pptx.ShapeType.line, { x:RX, y:SH-0.26, w:RW, h:0, line:{color:C.border, width:0.75} });
  slide.addText("Engineer Profile  |  Confidential", { x:RX, y:SH-0.22, w:5, h:0.18, fontSize:7, color:C.muted });
  slide.addText(month, { x:SW-2, y:SH-0.22, w:1.8, h:0.18, fontSize:7, color:C.muted, align:"right" });

  pptx.writeFile({ fileName: `${data["氏名"] ?? "profile"}_profile.pptx` });
}
