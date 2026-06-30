import OpenAI from "openai";
import { NextRequest } from "next/server";

const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const data = body.data as Record<string, string>;

    const name = data["氏名"] ?? "";
    const reason = data["エンジニアをやっている理由・原動力"] ?? "";
    const episode = data["自慢のエピソード"] ?? "";
    const catchphrase = data["キャッチフレーズ"] ?? "";
    const engineerType = data["エンジニアタイプ"] ?? "";

    const prompt = `以下のエンジニア情報をもとに、営業・提案シーンで使える簡潔な紹介文を80〜100文字で生成してください。

【制約】
- 「${name}は〜」のように氏名を主語にした書き出しは禁止。別の書き出しにすること。
- 体言止めは使わず、読みやすい自然な文体で書いてください。
- 80〜100文字に厳密に収めること。それ以上書かない。
- 紹介文のみを出力（見出しや説明は不要）。

氏名: ${name}
エンジニアタイプ: ${engineerType}
キャッチフレーズ: ${catchphrase}
エンジニアをやっている理由・原動力: ${reason}
自慢のエピソード: ${episode}`;

    const response = await client.chat.completions.create({
      model: "gpt-4o",
      messages: [{ role: "user", content: prompt }],
      max_tokens: 400,
    });

    const summary = response.choices[0]?.message?.content?.trim() ?? "";
    return Response.json({ summary });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return Response.json({ error: message }, { status: 500 });
  }
}
