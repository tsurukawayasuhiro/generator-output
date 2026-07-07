import OpenAI from "openai";
import { NextRequest } from "next/server";

const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

function cap(s: string, n = 500) { return s.slice(0, n); }

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const data = body.data as Record<string, string>;

    // 各フィールドを上限付きで取得（大量入力によるトークン浪費・インジェクション緩和）
    const name         = cap(data["氏名"] ?? "", 50);
    const engineerType = cap(data["エンジニアタイプ"] ?? "");
    const catchphrase  = cap(data["キャッチフレーズ"] ?? "");
    const reason       = cap(data["エンジニアをやっている理由・原動力"] ?? "");
    const episode      = cap(data["自慢のエピソード"] ?? "");

    // データは user ロールの JSON として渡す（system/user 分離でプロンプトインジェクション緩和）
    const response = await client.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "system",
          content: [
            "あなたは Atware 社内向けコピーライターです。",
            "エンジニア情報をもとに、営業・提案シーンで使える日本語の紹介文を生成してください。",
            "【制約】",
            `- 「${name}は〜」のように氏名を主語にした書き出しは禁止。`,
            "- 体言止めは使わず、自然な文体で。",
            "- 必ず100文字以内（日本語）で完結した文にすること。80〜100文字が理想。",
            "- 文は必ず句点（。）で終わること。",
            "- 紹介文のみを出力（見出し・説明不要）。",
          ].join("\n"),
        },
        {
          role: "user",
          content: JSON.stringify({ name, engineerType, catchphrase, reason, episode }),
        },
      ],
      max_tokens: 160,
    });

    const summary = response.choices[0]?.message?.content?.trim() ?? "";
    return Response.json({ summary });
  } catch (err) {
    console.error("[summarize-card]", err);
    return Response.json({ error: "要約の生成に失敗しました" }, { status: 500 });
  }
}
