import { NextRequest, NextResponse } from "next/server";

export function middleware(req: NextRequest) {
  const token = req.headers.get("x-internal-token");
  if (token !== process.env.INTERNAL_API_TOKEN) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  return NextResponse.next();
}

export const config = { matcher: ["/api/:path*"] };
