"use client";

import { forwardRef } from "react";

export const CARD_THEMES = [
  { id: "white",  avatarBg: "#F2F0EB", avatarText: "#A89F90", cardBg: "#FFFFFF",  border: "#E4E0D8" },
  { id: "slate",  avatarBg: "#DDE4ED", avatarText: "#8A9DB5", cardBg: "#F8FAFD",  border: "#CDD7E4" },
  { id: "sage",   avatarBg: "#DDE8E1", avatarText: "#8AAD95", cardBg: "#F7FBF8",  border: "#C8DDD0" },
  { id: "rose",   avatarBg: "#EDE0E0", avatarText: "#B59090", cardBg: "#FDF8F8",  border: "#E0CCCC" },
  { id: "indigo", avatarBg: "#E0E0ED", avatarText: "#9090B5", cardBg: "#F8F8FD",  border: "#CCCCE0" },
] as const;

export type ThemeId = typeof CARD_THEMES[number]["id"];

export interface CardFields {
  name: string;
  nameRoman: string;
  avatarText: string;  // アバターエリアに表示するテキスト（空の場合は氏名のイニシャル）
  engineerType: string;
  catchphrase: string;
  summary: string;
  skills: string;
}

interface Props {
  fields: CardFields;
  theme: ThemeId;
  editable?: boolean;
  onChange?: (fields: CardFields) => void;
}

const EngineerCard = forwardRef<HTMLDivElement, Props>(function EngineerCard(
  { fields, theme, editable = false, onChange },
  ref
) {
  const t = CARD_THEMES.find((c) => c.id === theme) ?? CARD_THEMES[0];

  const initials = fields.avatarText || fields.name
    .split(/\s+/)
    .map((s) => s[0])
    .join("")
    .slice(0, 2);

  function ep(key: keyof CardFields) {
    if (!editable) return {};
    return {
      contentEditable: true as const,
      suppressContentEditableWarning: true,
      onBlur: (e: React.FocusEvent<HTMLElement>) =>
        onChange?.({ ...fields, [key]: e.currentTarget.innerText }),
      style: { outline: "none", cursor: "text", borderBottom: "1px dashed #ccc" } as React.CSSProperties,
    };
  }

  return (
    <div
      ref={ref}
      style={{
        width: 280,
        fontFamily: "'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif",
        background: t.cardBg,
        borderRadius: 12,
        border: `1px solid ${t.border}`,
        boxShadow: "0 2px 16px rgba(0,0,0,0.07)",
        display: "flex",
        flexDirection: "column",
        overflow: "hidden",
      }}
    >
      {/* Avatar */}
      <div style={{ background: t.avatarBg, height: 90, display: "flex", alignItems: "center", justifyContent: "center", borderBottom: `1px solid ${t.border}` }}>
        <div
          style={{ width: 54, height: 54, borderRadius: "50%", background: "rgba(255,255,255,0.6)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, fontWeight: 700, color: t.avatarText, ...(editable ? { outline: "none", cursor: "text" } : {}) }}
          {...(editable ? {
            contentEditable: true as const,
            suppressContentEditableWarning: true,
            onBlur: (e: React.FocusEvent<HTMLElement>) =>
              onChange?.({ ...fields, avatarText: e.currentTarget.innerText.trim() }),
          } : {})}
        >
          {initials || "　"}
        </div>
      </div>

      <div style={{ padding: "14px 16px 18px", display: "flex", flexDirection: "column", gap: 7 }}>
        {/* Role */}
        <div style={{ fontSize: 10, color: "#999" }} {...ep("engineerType")}>
          {fields.engineerType.split(/[,、]/).map(s => s.trim()).filter(Boolean).join(" / ") || (editable ? "エンジニアタイプ" : "")}
        </div>

        {/* Catchphrase */}
        <div style={{ fontSize: 10.5, color: "#666", lineHeight: 1.5 }} {...ep("catchphrase")}>
          {fields.catchphrase || (editable ? "キャッチフレーズ" : "")}
        </div>

        <div style={{ height: 1, background: t.border }} />

        {/* Summary */}
        <div style={{ fontSize: 10.5, color: "#444", lineHeight: 1.75 }} {...ep("summary")}>
          {fields.summary || (editable ? "紹介文" : "")}
        </div>

        {/* Skills */}
        {(fields.skills || editable) && (
          <div style={{ borderTop: `1px solid ${t.border}`, paddingTop: 8 }}>
            <div style={{ fontSize: 9, fontWeight: 700, color: "#999", letterSpacing: "0.08em", textTransform: "uppercase" as const, marginBottom: 4 }}>
              スキル
            </div>
            <div style={{ fontSize: 10, color: "#555", lineHeight: 1.6 }} {...ep("skills")}>
              {fields.skills || (editable ? "主戦場スキル" : "")}
            </div>
          </div>
        )}
      </div>
    </div>
  );
});

export default EngineerCard;
