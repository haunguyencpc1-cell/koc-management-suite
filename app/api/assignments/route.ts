import { google } from "googleapis";
import { NextResponse } from "next/server";

export async function GET() {
  try {
    const email = (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "").trim();
    const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").trim().replace(/\\n/g, "\n");
    const sheetId = (process.env.GOOGLE_SHEET_ID || "").trim();

    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: email,
        private_key: privateKey,
      },
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });

    const sheets = google.sheets({ version: "v4", auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: "DANH MỤC KOC!B2:C",
    });

    const rows = res.data.values || [];
    const map: Record<string, string> = {};
    rows.forEach((row) => {
      const person = (row[0] || "").toString().trim();
      const creator = (row[1] || "").toString().trim();
      if (creator && person) map[creator] = person;
    });

    return NextResponse.json({ ok: true, map, count: Object.keys(map).length });
  } catch (e: any) {
    return NextResponse.json({
      ok: false,
      error: e?.response?.data?.error?.message || e.message || "Unknown error",
      details: {
        sheetIdUsed: process.env.GOOGLE_SHEET_ID || "MISSING",
        sheetIdLength: (process.env.GOOGLE_SHEET_ID || "").length,
        emailUsed: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "MISSING",
        hasPrivateKey: !!process.env.GOOGLE_PRIVATE_KEY,
        privateKeyLength: (process.env.GOOGLE_PRIVATE_KEY || "").length,
      },
    }, { status: 500 });
  }
}
