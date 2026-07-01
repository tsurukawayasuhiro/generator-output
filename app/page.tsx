"use client";

import { useEffect, useRef, useState } from "react";
import EngineerCard, { CARD_THEMES, type CardFields, type ThemeId } from "./components/EngineerCard";
import EngineerProfile from "./components/EngineerProfile";

type Phase = "idle" | "fetching-sheets" | "loading" | "summarizing" | "done";
type Tab = "card" | "profile";

async function apiFetch(path: string, body: unknown) {
  return fetch(path, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-internal-token": process.env.NEXT_PUBLIC_INTERNAL_API_TOKEN ?? "",
    },
    body: JSON.stringify(body),
  });
}

export default function Home() {
  const [spreadsheetId, setSpreadsheetId] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [phase, setPhase] = useState<Phase>("idle");
  const [error, setError] = useState("");
  const [tab, setTab] = useState<Tab>("card");

  const [fields, setFields] = useState<CardFields>({
    name: "", nameRoman: "", avatarText: "", engineerType: "", catchphrase: "", summary: "", skills: "",
  });
  const [theme, setTheme] = useState<ThemeId>("white");
  const [rawData, setRawData] = useState<Record<string, string>>({});

  const profileRef = useRef<HTMLDivElement>(null);
  const profileWrapRef = useRef<HTMLDivElement>(null);
  const isLoading = phase === "fetching-sheets" || phase === "loading" || phase === "summarizing";

  // 16:9プレビューのスケール計算
  useEffect(() => {
    function updateScale() {
      const wrap = profileWrapRef.current;
      const inner = profileRef.current;
      if (!wrap || !inner) return;
      const scale = wrap.clientWidth / 960;
      // CSS変数をstyle属性経由でセット
      (inner as HTMLElement).style.setProperty("--profile-scale", String(scale));
      // outer wrapper の高さを実際の描画高さに合わせる
      const outer = inner.parentElement as HTMLElement | null;
      if (outer) outer.style.height = `${540 * scale}px`;
    }
    updateScale();
    window.addEventListener("resize", updateScale);
    return () => window.removeEventListener("resize", updateScale);
  }, [tab, phase]);

  async function handleFetchSheets() {
    if (!spreadsheetId.trim()) { setError("スプレッドシート ID を入力してください"); return; }
    setError(""); setPhase("fetching-sheets");
    try {
      const res = await apiFetch("/api/read-sheet", { spreadsheetId: spreadsheetId.trim() });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error ?? "読み込みに失敗しました");
      const names: string[] = json.sheetNames ?? [];
      setSheetNames(names);
      setSheetName(names[0] ?? "");
      setPhase("idle");
    } catch (e) {
      setError(e instanceof Error ? e.message : "エラーが発生しました");
      setPhase("idle");
    }
  }

  async function handleLoad() {
    if (!spreadsheetId.trim()) { setError("スプレッドシート ID を入力してください"); return; }
    if (!sheetName.trim()) { setError("シートを選択してください"); return; }
    setError(""); setPhase("loading");
    try {
      const res = await apiFetch("/api/read-sheet", { spreadsheetId: spreadsheetId.trim(), sheetName: sheetName.trim() });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error ?? "読み込みに失敗しました");
      const data: Record<string, string> = json.data ?? {};
      setRawData(data);
      setFields({
        name: data["氏名"] ?? "",
        nameRoman: "",
        avatarText: "",
        engineerType: data["エンジニアタイプ"] ?? "",
        catchphrase: data["キャッチフレーズ"] ?? "",
        summary: "",
        skills: data["主戦場スキル"] ?? "",
      });
      setPhase("summarizing");
      const sres = await apiFetch("/api/summarize-card", { data });
      const sjson = await sres.json();
      if (!sres.ok) throw new Error(sjson.error ?? "要約生成に失敗しました");
      setFields((f) => ({ ...f, summary: sjson.summary ?? "" }));
      setPhase("done");
    } catch (e) {
      setError(e instanceof Error ? e.message : "エラーが発生しました");
      setPhase("idle");
    }
  }

  async function handleReSummarize() {
    if (!Object.keys(rawData).length) return;
    setPhase("summarizing");
    try {
      const res = await apiFetch("/api/summarize-card", { data: rawData });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error ?? "失敗しました");
      setFields((f) => ({ ...f, summary: json.summary ?? "" }));
    } finally {
      setPhase("done");
    }
  }

  async function handleDownloadCard() {
    const { generateCardPptx } = await import("./lib/generatePptx");
    generateCardPptx(fields, theme);
  }


  async function handleDownloadPptx() {
    if (!Object.keys(rawData).length) return;
    const { generatePptx } = await import("./lib/generatePptx");
    generatePptx(rawData);
  }

  const hasDone = phase === "done";

  return (
    <div className="min-h-screen bg-[#F5F8FC]">
      <header className="bg-[#0F2744] text-white px-6 py-5">
        <div className="max-w-5xl mx-auto">
          <div className="text-xs font-semibold tracking-widest text-blue-200 opacity-70 mb-1 uppercase">Atware Internal Tools</div>
          <h1 className="text-xl font-bold tracking-tight">generator-output</h1>
          <p className="text-sm text-white/60 mt-0.5">エンジニア名鑑 アウトプット生成ツール</p>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8 flex flex-col gap-6">

        {/* ── 入力 ── */}
        <section className="bg-white border border-[#D4DDE8] rounded-xl overflow-hidden">
          <div className="bg-[#F5F8FC] border-b border-[#D4DDE8] px-5 py-3 flex items-center gap-2">
            <span className="w-6 h-6 rounded-full bg-[#0057B8] text-white text-xs font-bold flex items-center justify-center">1</span>
            <span className="text-xs font-bold tracking-widest uppercase text-[#4A5F78]">データ読み込み</span>
          </div>
          <div className="p-5 flex flex-col gap-4">
            <div className="flex gap-3 items-end">
              <div className="flex flex-col gap-1.5 flex-1">
                <label className="text-xs font-semibold text-[#4A5F78] uppercase tracking-wide">スプレッドシート ID</label>
                <input
                  type="text"
                  placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
                  value={spreadsheetId}
                  onChange={(e) => { setSpreadsheetId(e.target.value); setSheetNames([]); setSheetName(""); }}
                  onKeyDown={(e) => e.key === "Enter" && handleFetchSheets()}
                  className="border border-[#D4DDE8] rounded-lg px-3 py-2 text-sm text-[#0F2744] placeholder:text-[#B0BEC8] focus:outline-none focus:ring-2 focus:ring-[#0057B8]/30 focus:border-[#0057B8]"
                />
              </div>
              <button onClick={handleFetchSheets} disabled={isLoading}
                className="bg-[#4A5F78] text-white text-sm font-semibold px-4 py-2 rounded-lg hover:bg-[#374a5e] disabled:opacity-50 disabled:cursor-not-allowed transition-colors whitespace-nowrap">
                {phase === "fetching-sheets" ? "取得中…" : "シート一覧を取得"}
              </button>
            </div>
            {sheetNames.length > 0 && (
              <div className="flex gap-3 items-end">
                <div className="flex flex-col gap-1.5 flex-1">
                  <label className="text-xs font-semibold text-[#4A5F78] uppercase tracking-wide">対象シート（氏名）</label>
                  <select
                    value={sheetName}
                    onChange={(e) => setSheetName(e.target.value)}
                    className="border border-[#D4DDE8] rounded-lg px-3 py-2 text-sm text-[#0F2744] focus:outline-none focus:ring-2 focus:ring-[#0057B8]/30 focus:border-[#0057B8] bg-white"
                  >
                    {sheetNames.map((n) => (
                      <option key={n} value={n}>{n}</option>
                    ))}
                  </select>
                </div>
                <button onClick={handleLoad} disabled={isLoading}
                  className="bg-[#0057B8] text-white text-sm font-semibold px-5 py-2 rounded-lg hover:bg-[#0046A0] disabled:opacity-50 disabled:cursor-not-allowed transition-colors whitespace-nowrap">
                  {phase === "loading" ? "読み込み中…" : phase === "summarizing" ? "AI 要約生成中…" : "読み込む"}
                </button>
              </div>
            )}
            {error && <div className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-lg px-3 py-2">{error}</div>}
          </div>
        </section>

        {/* ── プレビュー（タブ） ── */}
        {(hasDone || phase === "summarizing") && (
          <section className="bg-white border border-[#D4DDE8] rounded-xl overflow-hidden">
            {/* タブヘッダー */}
            <div className="flex border-b border-[#D4DDE8]">
              {([["card","🃏 トレーディングカード"], ["profile","📊 詳細プロフィール（16:9）"]] as [Tab, string][]).map(([t, label]) => (
                <button key={t} onClick={() => setTab(t)}
                  className={`px-5 py-3 text-sm font-semibold transition-colors border-b-2 ${
                    tab === t ? "border-[#0057B8] text-[#0057B8] bg-white" : "border-transparent text-[#7A90A8] bg-[#F5F8FC] hover:text-[#4A5F78]"
                  }`}>
                  {label}
                </button>
              ))}
            </div>

            {/* ── カードタブ ── */}
            {tab === "card" && (
              <div className="p-5 flex flex-col gap-5">
                <div className="flex items-center gap-4 flex-wrap">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-semibold text-[#7A90A8]">カラー</span>
                    <div className="flex gap-1.5">
                      {CARD_THEMES.map((t) => (
                        <button key={t.id} onClick={() => setTheme(t.id)} title={t.id}
                          style={{ background: t.avatarBg, borderColor: theme === t.id ? "#0057B8" : t.border }}
                          className="w-6 h-6 rounded-full border-2 transition-all" />
                      ))}
                    </div>
                  </div>
                  <div className="flex gap-2 ml-auto flex-wrap">
                    <button onClick={handleReSummarize} disabled={phase === "summarizing"}
                      className="border border-[#D4DDE8] text-[#4A5F78] text-sm font-medium px-3 py-1.5 rounded-lg hover:bg-[#F5F8FC] disabled:opacity-50 transition-colors">
                      {phase === "summarizing" ? "生成中…" : "要約を再生成"}
                    </button>
                    {hasDone && (
                      <button onClick={handleDownloadCard}
                        className="bg-[#D97706] text-white text-sm font-semibold px-4 py-1.5 rounded-lg hover:bg-[#B45309] transition-colors">
                        .pptx ダウンロード
                      </button>
                    )}
                  </div>
                </div>
                {hasDone && <p className="text-xs text-[#7A90A8]">テキストは直接クリックして編集できます</p>}
                <div className="flex justify-center py-4">
                  {phase === "summarizing"
                    ? <div className="text-sm text-[#7A90A8] animate-pulse py-16">AI が紹介文を生成しています…</div>
                    : <EngineerCard fields={fields} theme={theme} editable onChange={setFields} />
                  }
                </div>
              </div>
            )}

            {/* ── 詳細プロフィールタブ ── */}
            {tab === "profile" && (
              <div className="p-5 flex flex-col gap-4">
                <div className="flex justify-end gap-2">
                  <button onClick={handleDownloadPptx}
                    className="bg-[#D97706] text-white text-sm font-semibold px-4 py-1.5 rounded-lg hover:bg-[#B45309] transition-colors">
                    .pptx ダウンロード
                  </button>
                </div>
                {/* 16:9 プレビュー */}
                <p className="text-xs text-[#7A90A8]">氏名は直接クリックして編集できます（名前を伏せる場合など）</p>
                <div ref={profileWrapRef} className="w-full rounded-lg overflow-hidden border border-[#D4DDE8] shadow-sm">
                  <EngineerProfile ref={profileRef} fields={fields} rawData={rawData} editable onChange={setFields} />
                </div>
              </div>
            )}
          </section>
        )}
      </main>
    </div>
  );
}
