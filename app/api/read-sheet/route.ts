import { google } from "googleapis";
import { NextRequest } from "next/server";

// 認証インスタンスはモジュールスコープで1度だけ生成（毎リクエストで復号しない）
const auth = (() => {
  const b64 = process.env.GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_B64;
  if (!b64) throw new Error("GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_B64 is not set");
  const credentials = JSON.parse(Buffer.from(b64, "base64").toString("utf-8"));
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });
})();

const SPREADSHEET_ID_RE = /^[a-zA-Z0-9_-]{20,60}$/;

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { spreadsheetId, sheetName } = body as {
      spreadsheetId: string;
      sheetName?: string;
    };

    if (!spreadsheetId || !SPREADSHEET_ID_RE.test(spreadsheetId)) {
      return Response.json({ error: "Invalid spreadsheet ID" }, { status: 400 });
    }

    const sheets = google.sheets({ version: "v4", auth });

    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames =
      meta.data.sheets?.map((s) => s.properties?.title ?? "").filter(Boolean) ?? [];

    if (!sheetName) {
      return Response.json({ sheetNames, data: {} });
    }

    // sheetName を取得済み一覧で検証（レンジインジェクション防止）
    if (!sheetNames.includes(sheetName)) {
      return Response.json({ error: "Invalid sheet name" }, { status: 400 });
    }

    const range = `'${sheetName}'!A:B`;
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const rows = res.data.values ?? [];

    const data: Record<string, string> = {};
    for (const row of rows) {
      const key = (row[0] ?? "").trim();
      const value = (row[1] ?? "").trim();
      if (key && !key.startsWith("##")) {
        data[key] = value;
      }
    }

    return Response.json({ sheetNames, data });
  } catch (err) {
    console.error("[read-sheet]", err);
    return Response.json({ error: "シートの読み込みに失敗しました" }, { status: 500 });
  }
}
