import { google } from "googleapis";
import { NextResponse } from "next/server";

export async function GET() {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: {
        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
      },
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });

    const sheets = google.sheets({ version: "v4", auth });
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
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
    return NextResponse.json({ ok: false, error: e.message || "Unknown error" }, { status: 500 });
  }
}
