'use client';

import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const templates = [
  { key: "file", label: "原始考勤数据", hint: "每日打卡工时推送模版", req: true },
  { key: "leave", label: "离职流程", hint: "离职流程表", req: true },
  { key: "roster", label: "花名册", hint: "花名册", req: true },
  { key: "shift_dict", label: "班次", hint: "班次字典", req: true },
  { key: "makeup", label: "补签管理", hint: "补签管理记录", req: true },
  { key: "gus_whitelist", label: "GUS白名单", hint: "GUS需剔除人员", req: true },
  { key: "sign_this", label: "本周签字报表", hint: "GUS+美区签字报表（无括号）", req: true },
  { key: "sign_last", label: "上周签字报表", hint: "GUS+美区签字报表 (2)", req: true },
  { key: "sign_biweek", label: "双周签字报表", hint: "GUS+美区签字报表 (1)", req: false },
];

const rules: [string, string][] = [
  ["file", "每日打卡"], ["leave", "离职"], ["roster", "花名册"],
  ["shift_dict", "班次"], ["makeup", "补签管理"],
  ["sign_this", "本周"], ["sign_last", "上周"], ["sign_biweek", "双周"],
  ["gus_whitelist", "白名单"],
];

async function parseExcel(file: File): Promise<any[]> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  // 多 sheet 合并（GUS白名单有多 sheet 的剔除名单，只取第一个会漏剔）
  const sheets = (wb.SheetNames || []).filter((n) => wb.Sheets[n]);
  if (sheets.length === 0) return [];
  const all: any[] = [];
  for (const sn of sheets) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: "" });
    if (rows.length === 0) continue;
    if (all.length === 0) all.push(...rows);
    else {
      // 后续 sheet 列对齐到第一个 sheet 的列（缺列补空）
      const baseCols = Object.keys(all[0]);
      for (const r of rows) {
        const o: any = {};
        for (const c of baseCols) o[c] = (r as any)[c] ?? "";
        all.push(o);
      }
    }
  }
  return all;
}

export default function HomePage() {
  const [matched, setMatched] = useState<Record<string, string>>({});
  const [uploading, setUploading] = useState(false);
  const [progress, setProgress] = useState("");
  const [result, setResult] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const matchCount = templates.filter((t) => t.req && matched[t.key]).length;
  const canUpload = matchCount >= 7 && !uploading;

  const handleFiles = (list: File[]) => {
    const next = { ...matched };
    // 对每个文件，按优先级寻找最匹配的规则
    // 同一类型有多个文件时，优先选非【】前缀的（排除已完成/审批中等状态前缀）
    for (const [key, kw] of rules) {
      const candidates = list.filter((f) => f.name.includes(kw));
      if (candidates.length === 0) continue;
      // 多个候选时：优先选不以【开头的（排除状态前缀文件）
      const preferred = candidates.filter((f) => !f.name.startsWith("【") && !f.name.includes("- 副本"));
      const chosen = preferred.length > 0 ? preferred[0] : candidates[0];
      if (!next[key]) next[key] = chosen.name;
    }
    setMatched(next); setError(null); setResult(null);
  };

  const handleUpload = async () => {
    setUploading(true);
    setError(null);
    setResult(null);
    setProgress("正在解析 Excel 文件...");

    try {
      const inp = inputRef.current;
      if (!inp?.files) { setError("请选择文件"); setUploading(false); return; }

      const fileMap: Record<string, File> = {};
      for (const [key, kw] of rules) {
        const candidates = Array.from(inp.files).filter((f) => f.name.includes(kw));
        if (candidates.length === 0) continue;
        // 同一类型多个候选时：优先选不以【开头的
        const preferred = candidates.filter((f) => !f.name.startsWith("【") && !f.name.includes("- 副本"));
        fileMap[key] = preferred.length > 0 ? preferred[0] : candidates[0];
      }

      const missing = templates.filter((t) => t.req && !fileMap[t.key]).map((t) => t.label);
      if (missing.length > 0) { setError("缺少: " + missing.join(", ")); setUploading(false); return; }

      const payload: Record<string, any[]> = {};
      const entries = Object.entries(fileMap);
      for (let i = 0; i < entries.length; i++) {
        const [key, file] = entries[i];
        setProgress("解析 " + (i + 1) + "/" + entries.length + ": " + file.name);
        try { payload[key] = await parseExcel(file); }
        catch (e) { payload[key] = []; }
      }

      setProgress("上传中...");

      const signKeys = new Set(["sign_this", "sign_last", "sign_biweek"]);
      const compressed: Record<string, any[]> = {};
      const mainKeep = ["工号","考勤日期","姓名","三级部门","四级部门","五级部门","部门","班次名称",
        "班次上班时间","班次下班时间","首打卡时间","末打卡时间","班次内打卡次数","是否排班正确",
        "每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计","日超8H","是否日超8H",
        "本周加班工时","上周累计加班工时","HUB","备注（GF）","居家办公合计（审批中）","补签数",
        "休息开始时间","休息结束时间",
        // 输出字段所需列（压缩后丢失会导致输出 Excel 空列）
        "职位","居家办公工时（全）合计","累计总工时",
        // 清洗步骤所需列
        "最后工作日","审批状态","入职日期","补签日期","加班合计","加班时间段开始1","加班时间段结束1","加班时间段开始2","加班时间段结束2","OT1.5合计","OT2.0合计","计薪出勤时长合计（REG）",
        // step7 签字报表所需列（工时列全名）
        "每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)"];
      const hoursKw = ["每日总工时", "总工时", "工时", "时长", "小时"];
      for (const [key, rows] of Object.entries(payload)) {
        if (signKeys.has(key)) {
          compressed[key] = rows.map((r: any) => {
            // 优先匹配完整列名"每日总工时(公式：..."
            const allKeys = Object.keys(r);
            const dailyCol = allKeys.find((k) => k.includes("每日总工时"));
            const hc = dailyCol || allKeys.find((k) => hoursKw.some((h) => k.includes(h)));
            // 保留姓名列：step7 依赖「姓名含(合计)」筛选周汇总行
            const nc = allKeys.find((k) => k.includes("姓名"));
            return { 工号: r["工号"], 每日总工时: hc ? r[hc] : 0, 姓名: nc ? r[nc] : "" };
          });
        } else {
          compressed[key] = rows.map((r: any) => {
            const o: any = {};
            for (const f of mainKeep) { if (f in r) o[f] = r[f]; }
            return o;
          });
        }
      }

      const res = await fetch("/api/upload/pipeline", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ files: compressed, worker_type: "GUS" }),
      });

      const text = await res.text();
      let data: any;
      try { data = JSON.parse(text); }
      catch (e2) { setError("服务器异常: " + text.slice(0, 200)); setUploading(false); return; }

      if (!res.ok) {
        setError(data?.detail ? data.error + " (" + data.detail + ")" : data?.error || "处理失败");
        setUploading(false);
        return;
      }

      setResult(data);

      if (data.rows && data.rows.length > 0) {
        const hd = Object.keys(data.rows[0]);
        const ws = XLSX.utils.aoa_to_sheet([hd, ...data.rows.map((r: any) => hd.map((h: string) => r[h]))]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "处理结果");
        const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
        const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = "result_" + data.date + ".xlsx"; a.click();
        URL.revokeObjectURL(url);
      }
    } catch (e) {
      setError((e as Error).message || "网络错误");
    } finally {
      setUploading(false);
      setProgress("");
    }
  };

  return (
    <>
      <header className="border-b bg-white sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-4 h-14 flex items-center justify-between">
          <a href="/" className="flex items-center gap-2 font-semibold text-lg">
            <svg className="w-6 h-6 text-primary-600" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="3" y="3" width="18" height="18" rx="2" /><path d="M3 9h18M9 21V9" />
            </svg>考勤处理平台
          </a>
          <nav className="flex items-center gap-4 text-sm">
            <a href="/" className="hover:text-primary-600 transition-colors">上传处理</a>
            <a href="/history" className="hover:text-primary-600 transition-colors">历史记录</a>
          </nav>
        </div>
      </header>
      <main className="max-w-6xl mx-auto px-4 py-8">
        <div className="text-center mb-8">
          <h1 className="text-2xl font-bold mb-2">考勤数据处理</h1>
          <p className="text-gray-500 text-sm">上传 Excel 文件，自动完成数据清洗、指标计算、透视分析和报告生成</p>
        </div>
        <div className="w-full max-w-2xl mx-auto">
          <div className="border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all border-gray-300 hover:border-primary-300 hover:bg-gray-50"
            onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleFiles(Array.from(e.dataTransfer.files)); }}
            onClick={() => inputRef.current?.click()}>
            <input ref={inputRef} type="file" multiple accept=".xlsx,.xls" className="hidden" onChange={e => e.target.files && handleFiles(Array.from(e.target.files))} />
            <svg className="w-10 h-10 mx-auto mb-3 text-gray-400" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" />
            </svg>
            <p className="text-gray-600 font-medium">拖拽所有 Excel 文件到这里，自动匹配类型</p>
            <p className="text-sm text-gray-400 mt-1">支持 .xlsx / .xls，单次最多 50MB</p>
          </div>
          <div className="mt-4 space-y-2">
            {templates.map(t => {
              const m = !!matched[t.key];
              return (
                <div key={t.key} className={"flex items-center gap-3 px-4 py-2.5 rounded-lg border transition-all " + (m ? "bg-green-50 border-green-200" : "bg-white border-gray-200 hover:border-gray-300")}>
                  <div className="w-36 shrink-0 flex items-center gap-1.5">
                    <div className={"w-4 h-4 rounded-full border-2 " + (m ? "bg-green-400 border-green-400" : "border-amber-300")} />
                    <span className="text-sm font-medium text-gray-600">{t.label}</span>
                    {!t.req && <span className="text-[10px] text-gray-400">可选</span>}
                  </div>
                  <span className={"flex-1 text-sm truncate " + (m ? "text-gray-700" : "text-gray-300")}>{matched[t.key] || t.hint}</span>
                </div>
              );
            })}
          </div>
          <div className="mt-6 flex items-center justify-between">
            <span className="text-xs text-gray-400">已匹配 {matchCount} / 7 类必填</span>
            <button disabled={!canUpload} onClick={handleUpload}
              className={"px-8 py-2.5 rounded-lg font-medium text-white transition-all " + (canUpload ? "bg-primary-600 hover:bg-primary-700" : "bg-gray-300 cursor-not-allowed")}>
              {uploading ? "处理中..." : "上传并处理"}
            </button>
          </div>
          {progress && <p className="mt-3 text-sm text-blue-600">{progress}</p>}
          {result && (
            <div className="mt-6 p-4 bg-green-50 border border-green-200 rounded-lg text-sm">
              <p className="text-green-700 font-medium">✓ 处理完成 · {result.total} 行</p>
              <p className="text-green-600 mt-1">原始 {result.stats?.original} 行 → 最终 {result.total} 行 · 日期: {result.date}</p>
            </div>
          )}
          {error && (
            <div className="mt-6 p-4 bg-red-50 border border-red-200 rounded-lg">
              <p className="text-red-700 font-medium">× 失败</p>
              <p className="text-sm text-red-600 mt-1">{error}</p>
            </div>
          )}
        </div>
      </main>
    </>
  );
}
