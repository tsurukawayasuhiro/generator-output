import { google } from "googleapis";
import { NextRequest } from "next/server";

function getAuth() {
  const b64 = process.env.GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_B64;
  if (!b64) throw new Error("GOOGLE_SERVICE_ACCOUNT_CREDENTIALS_B64 is not set");
  const json = Buffer.from(b64, "base64").toString("utf-8");
  const credentials = JSON.parse(json);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { spreadsheetId, sheetName } = body as {
      spreadsheetId: string;
      sheetName?: string;
    };

    if (!spreadsheetId) {
      return Response.json({ error: "spreadsheetId is required" }, { status: 400 });
    }

    const auth = getAuth();
    const sheets = google.sheets({ version: "v4", auth });

    // Get all sheet names
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames =
      meta.data.sheets?.map((s) => s.properties?.title ?? "").filter(Boolean) ?? [];

    if (!sheetName) {
      return Response.json({ sheetNames, data: {} });
    }

    // Read A:B columns of the specified sheet
    const range = `'${sheetName}'!A:B`;
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const rows = res.data.values ?? [];

    // Build key-value map (skip section rows starting with ##)
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
    const message = err instanceof Error ? err.message : "Unknown error";
    return Response.json({ error: message }, { status: 500 });
  }
}
