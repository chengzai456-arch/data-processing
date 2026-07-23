import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import { runPipeline } from "@/lib/processing/pipeline";
import crypto from "crypto";

export const dynamic = "force-dynamic";
export const maxDuration = 300;

export async function POST(req: NextRequest) {
  try {
    const ct = req.headers.get("content-type") || "";

    // JSON mode: 浏览器端已解析 Excel -> 直接拿 JSON 行
    if (ct.includes("application/json")) {
      const body = await req.json();
      const fileData: Record<string, any[]> = body.files || {};

      if (!fileData["file"] || fileData["file"].length === 0) {
        return NextResponse.json({ error: "缺少原始数据" }, { status: 400 });
      }

      // 把 JSON 行转成 xlsx Buffer（以便 pipeline.ts 用 XLSX.read 解析）
      function jsonToBuffer(rows: any[]): Buffer {
        if (!rows || rows.length === 0) return Buffer.alloc(0);
        const h = Object.keys(rows[0]);
        const ws = XLSX.utils.aoa_to_sheet([
          h,
          ...rows.map((r: any) => h.map((k: string) => r[k])),
        ]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        return Buffer.from(
          XLSX.write(wb, { type: "buffer", bookType: "xlsx" }),
        );
      }

      const files: any = { file: jsonToBuffer(fileData["file"]) };
      for (const key of [
        "roster",
        "leave",
        "makeup",
        "sign_this",
        "sign_last",
        "sign_biweek",
        "gus_whitelist",
        "shift_dict",
      ]) {
        if (fileData[key]) files[key] = jsonToBuffer(fileData[key]);
      }

      const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
      const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
      let supabase: any = null;
      if (supabaseUrl && supabaseKey) {
        const { createClient } = await import("@supabase/supabase-js");
        supabase = createClient(supabaseUrl, supabaseKey, {
          auth: { autoRefreshToken: false, persistSession: false },
        });
      }

      const uploadId = crypto.randomUUID();
      const workerType = body["worker_type"] || "GUS";
      const targetDate = body["target_date"] || "";

      const result = await runPipeline(
        files,
        workerType,
        supabase,
        uploadId,
        targetDate,
      );
      return NextResponse.json(result);
    }

    // multipart/form-data mode: 直接传 Excel（兼容，小文件可用）
    const formData = await req.formData();
    const files: Record<string, Buffer> = {};
    const fields: Record<string, string> = {};
    for (const [key, val] of formData.entries()) {
      if (val instanceof File)
        files[key] = Buffer.from(await val.arrayBuffer());
      else fields[key] = String(val);
    }
    if (!files["file"])
      return NextResponse.json({ error: "缺少原始数据" }, { status: 400 });

    const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
    const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
    let supabase: any = null;
    if (supabaseUrl && supabaseKey) {
      const { createClient } = await import("@supabase/supabase-js");
      supabase = createClient(supabaseUrl, supabaseKey, {
        auth: { autoRefreshToken: false, persistSession: false },
      });
    }

    const uploadId = crypto.randomUUID();
    const result = await runPipeline(
      files as any,
      fields["worker_type"] || "GUS",
      supabase,
      uploadId,
      fields["target_date"] || "",
    );
    return NextResponse.json(result);
  } catch (e: any) {
    console.error("[pipeline] 错误:", e);
    return NextResponse.json(
      { error: e.message || "处理失败" },
      { status: 500 },
    );
  }
}
