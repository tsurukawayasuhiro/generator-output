"use client";

import { forwardRef } from "react";
import type { CardFields } from "./EngineerCard";

interface Props {
  fields: CardFields;
  rawData: Record<string, string>;
  editable?: boolean;
  onChange?: (fields: CardFields) => void;
}

function first3(data: Record<string, string>, prefix: string): string[] {
  return [1, 2, 3].map((i) => data[`${prefix}${i}`] ?? "").filter(Boolean);
}

function getProjects(data: Record<string, string>) {
  return [1, 2, 3].map((i) => {
    // 「案件N：」で始まる全フィールドを収集
    const prefix = `案件${i}：`;
    const extra: { label: string; value: string }[] = [];
    const KNOWN = new Set(["概要", "技術スタック", "役割", "期間", "規模", "成果"]);

    for (const [key, val] of Object.entries(data)) {
      if (key.startsWith(prefix) && val) {
        const label = key.slice(prefix.length);
        if (!KNOWN.has(label)) extra.push({ label, value: val });
      }
    }

    return {
      overview: data[`${prefix}概要`] ?? "",
      stack:    data[`${prefix}技術スタック`] ?? "",
      role:     data[`${prefix}役割`] ?? "",
      period:   data[`${prefix}期間`] ?? "",
      scale:    data[`${prefix}規模`] ?? "",
      result:   data[`${prefix}成果`] ?? "",
      extra,    // その他のフィールド
    };
  }).filter((p) => p.overview);
}

const C = {
  navy: "#0F2744",
  blue: "#0057B8", bluePale: "#E8F1FB", blueBorder: "#C0D8F5",
  orange: "#D97706", orangePale: "#FEF3C7", orangeBorder: "#F5D78A",
  green: "#065F46", greenPale: "#D1FAE5", greenBorder: "#99D8C0",
  slate: "#4A5F78", muted: "#7A90A8",
  border: "#D4DDE8", bg: "#F5F8FC", white: "#FFFFFF",
};

const W = 960;
const H = 540;
const LW = 240;

function Label({ text, color }: { text: string; color: string }) {
  return (
    <div style={{ fontSize: 8, fontWeight: 700, color, letterSpacing: "0.12em",
      textTransform: "uppercase" as const, marginBottom: 5 }}>
      {text}
    </div>
  );
}

function CircleBadge({ n, color }: { n: number; color: string }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", justifyContent: "center",
      width: 17, height: 17, borderRadius: "50%",
      background: color, color: C.white, fontSize: 8.5, fontWeight: 700, flexShrink: 0,
    }}>{n}</span>
  );
}

function StackTags({ stack }: { stack: string }) {
  const tags = stack.split(/[,、]/).map(s => s.trim()).filter(Boolean);
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
      {tags.map((t, i) => (
        <span key={i} style={{
          fontSize: 7.5, fontWeight: 600, padding: "2px 6px", borderRadius: 3,
          background: C.bluePale, color: C.blue, border: `1px solid ${C.blueBorder}`,
        }}>{t}</span>
      ))}
    </div>
  );
}

const EngineerProfile = forwardRef<HTMLDivElement, Props>(function EngineerProfile(
  { fields, rawData, editable = false, onChange },
  ref
) {
  function ep(key: keyof CardFields) {
    if (!editable) return {};
    return {
      contentEditable: true as const,
      suppressContentEditableWarning: true,
      className: "profile-editable",
      onBlur: (e: React.FocusEvent<HTMLElement>) =>
        onChange?.({ ...fields, [key]: e.currentTarget.innerText }),
    };
  }
  const values   = first3(rawData, "価値の出し方");
  const kodawari = first3(rawData, "こだわり");
  const projs    = getProjects(rawData);
  const atwValue = rawData["ATWとして提供できる価値 → 価値"] ?? "";
  const month    = new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" });

  return (
    <div style={{ width: "100%", aspectRatio: "16/9", position: "relative" }}>
      <div
        ref={ref}
        style={{
          position: "absolute", inset: 0, width: W, height: H,
          transformOrigin: "top left",
          transform: "scale(var(--profile-scale, 1))",
          fontFamily: "'Hiragino Kaku Gothic ProN', 'Meiryo', 'Yu Gothic', sans-serif",
          overflow: "hidden", background: C.bg, display: "flex",
        }}
      >
        {/* ════ 左パネル ════ */}
        <div style={{ width: LW, background: C.navy, display: "flex", flexDirection: "column", flexShrink: 0 }}>
          <div style={{ flex: 1, padding: "18px 16px 0", overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <div style={{ fontSize: 6.5, fontWeight: 700, letterSpacing: "0.15em",
              color: "rgba(255,255,255,0.3)", textTransform: "uppercase" as const, marginBottom: 7 }}>
              ATWARE ENGINEER PROFILE
            </div>
            <div style={{ fontSize: 23, fontWeight: 800, color: C.white, letterSpacing: "-0.02em", lineHeight: 1.15, marginBottom: 7 }}
              {...ep("name")}>
              {fields.name}
            </div>
            <div style={{ fontSize: 9, color: "rgba(255,255,255,0.55)", fontStyle: "italic", lineHeight: 1.5, marginBottom: 10 }}>
              &ldquo;{fields.catchphrase}&rdquo;
            </div>
            {fields.summary && (
              <div style={{ fontSize: 9, color: "rgba(255,255,255,0.78)", lineHeight: 1.7, marginBottom: 10 }}>
                {fields.summary}
              </div>
            )}
            <div style={{ height: 1, background: "rgba(255,255,255,0.1)", marginBottom: 10 }} />
            {[
              ["エンジニアタイプ", fields.engineerType],
              ["主戦場スキル",     fields.skills],
              ["担ってきた役割",   rawData["担ってきた役割"]],
            ].filter(([, v]) => v).map(([label, val]) => (
              <div key={label} style={{ marginBottom: 8 }}>
                <div style={{ fontSize: 6.5, fontWeight: 700, letterSpacing: "0.1em",
                  color: "#7AAED4", textTransform: "uppercase" as const, marginBottom: 2 }}>{label}</div>
                <div style={{ fontSize: 9, color: "#D4E8F8", lineHeight: 1.55 }}>{val}</div>
              </div>
            ))}
          </div>
          <div style={{ padding: "7px 16px", background: "rgba(0,0,0,0.25)" }}>
            <div style={{ fontSize: 8, fontWeight: 700, color: "#4A7AAA" }}>atWare</div>
          </div>
        </div>

        {/* ════ 右エリア（absolute配置で高さを確定） ════ */}
        <div style={{ flex: 1, position: "relative", overflow: "hidden" }}>
          {/* 座標系: 右エリア幅=720, 高さ=540 */}
          {/* ① 価値の出し方: top=14, height=76 */}
          <div style={{ position: "absolute", top: 14, left: 14, right: 14, height: 76, overflow: "hidden" }}>
            <Label text="価値の出し方" color={C.blue} />
            <div style={{ display: "flex", gap: 6, height: 57 }}>
              {[0,1,2].map((i) => (
                <div key={i} style={{
                  flex: 1, background: C.bluePale, border: `1px solid ${C.blueBorder}`,
                  borderRadius: 5, padding: "5px 8px", display: "flex", alignItems: "flex-start", gap: 5,
                  overflow: "hidden",
                }}>
                  <CircleBadge n={i+1} color={C.blue} />
                  <div style={{ fontSize: 9.5, color: C.navy, lineHeight: 1.4, overflow: "hidden" }}>{values[i] ?? "—"}</div>
                </div>
              ))}
            </div>
          </div>

          {/* ② 代表案件: top=98, bottom=430 */}
          <div style={{ position: "absolute", top: 98, left: 14, right: 14, bottom: 110, overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <Label text="代表案件" color={C.slate} />
            <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 6, overflow: "hidden" }}>
              {projs.length === 0 && (
                <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 5, padding: "8px 10px", fontSize: 9, color: C.muted }}>—</div>
              )}
              {projs.slice(0, 3).map((p, i) => (
                <div key={i} style={{
                  flex: 1, background: C.white, border: `1px solid ${C.border}`,
                  borderRadius: 5, padding: "8px 12px",
                  display: "flex", gap: 12, overflow: "hidden",
                }}>
                  {/* 左：番号バッジ */}
                  <div style={{ flexShrink: 0, paddingTop: 1 }}>
                    <span style={{
                      fontSize: 7.5, fontWeight: 700, padding: "2px 7px", borderRadius: 3,
                      background: C.bluePale, color: C.blue, border: `1px solid ${C.blueBorder}`,
                    }}>案件 {i+1}</span>
                  </div>

                  {/* 中央：概要・役割・成果・その他 */}
                  <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 3, minWidth: 0, overflow: "hidden" }}>
                    <div style={{ fontSize: 10.5, color: C.navy, lineHeight: 1.6, fontWeight: 500 }}>
                      {p.overview}
                    </div>
                    {p.role && (
                      <div style={{ fontSize: 8.5, color: C.slate }}>
                        <span style={{ fontWeight: 700, color: C.muted, marginRight: 4 }}>役割</span>{p.role}
                      </div>
                    )}
                    {(p.period || p.scale) && (
                      <div style={{ fontSize: 8.5, color: C.muted }}>
                        {[p.period, p.scale].filter(Boolean).join("　")}
                      </div>
                    )}
                    {p.extra.map(({ label, value }) => (
                      <div key={label} style={{ fontSize: 8.5, color: C.slate }}>
                        <span style={{ fontWeight: 700, color: C.muted, marginRight: 4 }}>{label}</span>{value}
                      </div>
                    ))}
                    {p.result && (
                      <div style={{
                        display: "inline-flex", alignItems: "center", gap: 4,
                        background: C.greenPale, border: `1px solid ${C.greenBorder}`,
                        borderRadius: 4, padding: "2px 8px",
                        fontSize: 8.5, color: C.green, flexShrink: 0,
                      }}>
                        <span style={{ fontWeight: 700 }}>成果</span>{p.result}
                      </div>
                    )}
                  </div>

                  {/* 右：技術スタックタグ */}
                  {p.stack && (
                    <div style={{ flexShrink: 0, width: 200, overflow: "hidden" }}>
                      <div style={{ fontSize: 7.5, fontWeight: 700, color: C.muted, marginBottom: 4 }}>技術スタック</div>
                      <StackTags stack={p.stack} />
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* ③ こだわり: bottom=22, height=84 */}
          <div style={{ position: "absolute", left: 14, right: 14, bottom: 22, height: 84, overflow: "hidden" }}>
            <Label text="こだわり" color={C.orange} />
            <div style={{ display: "flex", gap: 6, height: 65 }}>
              {[0,1,2].map((i) => (
                <div key={i} style={{
                  flex: 1, background: C.orangePale, border: `1px solid ${C.orangeBorder}`,
                  borderRadius: 5, padding: "5px 8px", display: "flex", alignItems: "flex-start", gap: 5,
                  overflow: "hidden",
                }}>
                  <CircleBadge n={i+1} color={C.orange} />
                  <div style={{ fontSize: 9.5, color: C.navy, lineHeight: 1.4, overflow: "hidden" }}>{kodawari[i] ?? "—"}</div>
                </div>
              ))}
              {atwValue && (
                <div style={{
                  flex: 1, background: "#EDE9FE", border: "1px solid #C4B5FD",
                  borderRadius: 5, padding: "5px 8px", overflow: "hidden",
                }}>
                  <div style={{ fontSize: 6.5, fontWeight: 700, color: "#6D28D9",
                    textTransform: "uppercase" as const, marginBottom: 3 }}>ATW提供価値</div>
                  <div style={{ fontSize: 9, color: "#5B21B6", lineHeight: 1.4 }}>{atwValue}</div>
                </div>
              )}
            </div>
          </div>

          {/* フッター: bottom=4, height=16 */}
          <div style={{
            position: "absolute", left: 14, right: 14, bottom: 4, height: 16,
            borderTop: `1px solid ${C.border}`, paddingTop: 3,
            display: "flex", justifyContent: "space-between",
          }}>
            <span style={{ fontSize: 7, color: C.muted }}>Engineer Profile  |  Confidential</span>
            <span style={{ fontSize: 7, color: C.muted }}>{month}</span>
          </div>
        </div>
      </div>
    </div>
  );
});

export default EngineerProfile;
