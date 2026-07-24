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

      // 校验原始数据是否存在
      if (!fileData["file"] || fileData["file"].length === 0) {
        return NextResponse.json(
          { error: "缺少原始数据" },
          { status: 400 },
        );
      }

      // 逐行解析，容错跳过非法行
      function safeParseRow(raw: any): any {
        if (!raw || typeof raw !== "object") return null;
        const code = String(raw["工号"] ?? "").trim();
        if (!code) return null;
        return raw;
      }

      // 过滤有效行
      fileData["file"] = fileData["file"].filter(safeParseRow);

      if (fileData["file"].length === 0) {
        return NextResponse.json(
          { error: "原始数据中未找到有效行（缺少工号）" },
          { status: 400 },
        );
      }

      // 把 JSON 行转成 xlsx Buffer（以便 pipeline.ts 用 XLSX.read 解析）
      function jsonToBuffer(rows: any[]): Buffer {
        if (!rows || rows.length === 0) return Buffer.alloc(0);
        try {
          const h = Object.keys(rows[0]);
          const ws = XLSX.utils.aoa_to_sheet([
            h,
            ...rows.map((r: any) => h.map((k: string) => r[k] ?? "")),
          ]);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
          return Buffer.from(
            XLSX.write(wb, { type: "buffer", bookType: "xlsx" }),
          );
        } catch (e) {
          console.warn("[pipeline] jsonToBuffer 失败:", e);
          return Buffer.alloc(0);
        }
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
        if (fileData[key]?.length > 0) {
          files[key] = jsonToBuffer(fileData[key]);
        }
      }

      // 校验核心文件是否成功转换
      if (!files.file || files.file.length === 0) {
        return NextResponse.json(
          { error: "原始数据解析失败，请检查文件格式" },
          { status: 400 },
        );
      }

      const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
      const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
      let supabase: any = null;
      if (supabaseUrl && supabaseKey) {
        try {
          const { createClient } = await import("@supabase/supabase-js");
          supabase = createClient(supabaseUrl, supabaseKey, {
            auth: { autoRefreshToken: false, persistSession: false },
          });
        } catch (e) {
          console.warn("[pipeline] Supabase 客户端创建失败:", e);
        }
      }

      const uploadId = crypto.randomUUID();
      const workerType = body["worker_type"] || "GUS";
      const targetDate = body["target_date"] || "";

      // 用 try-catch 包裹整个管道调用
      try {
        const result = await runPipeline(
          files,
          workerType,
          supabase,
          uploadId,
          targetDate,
        );
        return NextResponse.json(result);
      } catch (pipelineErr: any) {
        console.error("[pipeline] 管道执行失败:", pipelineErr);
        return NextResponse.json(
          {
            error: pipelineErr.message || "数据处理管道异常",
            detail: "请检查输入文件格式是否正确，或联系管理员查看日志",
          },
          { status: 500 },
        );
      }
    }

    // multipart/form-data mode: 直接传 Excel（兼容，小文件可用）
    try {
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
        try {
          const { createClient } = await import("@supabase/supabase-js");
          supabase = createClient(supabaseUrl, supabaseKey, {
            auth: { autoRefreshToken: false, persistSession: false },
          });
        } catch (e) {
          console.warn("[pipeline] Supabase 客户端创建失败:", e);
        }
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
    } catch (formErr: any) {
      console.error("[pipeline] multipart 处理失败:", formErr);
      return NextResponse.json(
        {
          error: formErr.message || "文件上传处理异常",
          detail: "请检查上传文件是否损坏，或文件体积是否超过限制",
        },
        { status: 400 },
      );
    }
  } catch (e: any) {
    console.error("[pipeline] 请求处理异常:", e);
    return NextResponse.json(
      {
        error: "请求处理失败",
        detail: e.message || "未知错误",
      },
      { status: 500 },
    );
  }
}
