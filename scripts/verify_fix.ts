import * as fs from "fs";
import * as XLSX from "xlsx";
import { runPipeline } from "../lib/processing/pipeline";

const DATA_DIR = "D:/Documents/Downloads/722日晾晒数据";
const load = (n: string) => fs.readFileSync(`${DATA_DIR}/${n}`);

// —— 完全模拟修复后的前端 page.tsx 逻辑 ——
// parseExcel: 多 sheet 合并 + 列对齐
function parseExcelLikeFrontend(buf: Buffer): Record<string, unknown>[] {
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheets = (wb.SheetNames || []).filter((n) => wb.Sheets[n]);
  const all: any[] = [];
  for (const sn of sheets) {
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[sn], { defval: "" });
    if (rows.length === 0) continue;
    if (all.length === 0) all.push(...rows);
    else {
      const baseCols = Object.keys(all[0]);
      for (const r of rows) {
        const o: any = {};
        for (const c of baseCols) o[c] = r[c] ?? "";
        all.push(o);
      }
    }
  }
  return all;
}

const signKeys = new Set(["sign_this", "sign_last", "sign_biweek"]);
const mainKeep = ["工号","考勤日期","姓名","三级部门","四级部门","五级部门","部门","班次名称",
  "班次上班时间","班次下班时间","首打卡时间","末打卡时间","班次内打卡次数","是否排班正确",
  "每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计","日超8H","是否日超8H",
  "本周加班工时","上周累计加班工时","HUB","备注（GF）","居家办公合计（审批中）","补签数",
  "休息开始时间","休息结束时间","职位","居家办公工时（全）合计","累计总工时","最后工作日","审批状态","入职日期","补签日期","加班合计",
  "加班时间段开始1","加班时间段结束1","加班时间段开始2","加班时间段结束2","OT1.5合计",
  "OT2.0合计","计薪出勤时长合计（REG）","每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)"];
const hoursKw = ["每日总工时", "总工时", "工时", "时长", "小时"];

function compress(buf: Buffer, key: string): Record<string, unknown>[] {
  const rows = parseExcelLikeFrontend(buf);
  if (signKeys.has(key)) {
    return rows.map((r: any) => {
      const allKeys = Object.keys(r);
      const dailyCol = allKeys.find((k) => k.includes("每日总工时"));
      const hc = dailyCol || allKeys.find((k) => hoursKw.some((h) => k.includes(h)));
      const nc = allKeys.find((k) => k.includes("姓名"));
      return { 工号: r["工号"], 每日总工时: hc ? r[hc] : 0, 姓名: nc ? r[nc] : "" };
    });
  }
  return rows.map((r: any) => {
    const o: any = {};
    for (const f of mainKeep) if (f in r) o[f] = r[f];
    return o;
  });
}

function j2b(rows: Record<string, unknown>[]): Buffer {
  if (!rows?.length) return Buffer.alloc(0);
  const h = Object.keys(rows[0]);
  const ws = XLSX.utils.aoa_to_sheet([h, ...rows.map((r) => h.map((k) => r[k] ?? ""))]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "S");
  return Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
}

const raw: Record<string, Buffer> = {
  file: load("722GUS+每日打卡工时推送模版 (51).xlsx"),
  leave: load("722离职流程 (62).xlsx"),
  roster: load("722花名册 (69).xlsx"),
  shift_dict: load("班次.xlsx"),
  makeup: load("722补签管理 (77).xlsx"),
  gus_whitelist: load("722GUS白名单人员 (3).xlsx"),
  sign_this: load("720-722【本周加班工时】GUS+美区签字报表 (99).xlsx"),
  sign_last: load("713-719【上周加班工时】GUS+美区签字报表 (100).xlsx"),
  sign_biweek: load("713-726【双周累计工时】GUS+美区签字报表 (98).xlsx"),
};

async function main() {
  // 修复后的生产压缩路径
  const prod: any = {};
  for (const k of Object.keys(raw)) prod[k] = j2b(compress(raw[k], k));
  // 直接文件解析（正确基准）
  const direct: any = { ...raw };

  const prodRes = await runPipeline(prod, "GUS", null, "prod-fixed", "2026-07-22");
  const dirRes = await runPipeline(direct, "GUS", null, "direct", "2026-07-22");

  const pr = prodRes.rows!;
  const dr = dirRes.rows!;
  const count = (rows: any[], k: string) => rows.filter((r) => Math.abs(Number(r[k] || 0)) > 0.001).length;
  const sum = (rows: any[], k: string) => rows.reduce((s, r) => s + Math.abs(Number(r[k] || 0)), 0);

  console.log("=== 修复后：生产压缩路径 vs 直接文件解析 ===");
  console.log("最终行数: 生产", prodRes.total, " vs 直接", dirRes.total);
  const gusProd = prodRes.stats.original - prodRes.stats.after_step5_gus;
  const gusDir = dirRes.stats.original - dirRes.stats.after_step5_gus;
  console.log("step5剔除: 生产", gusProd, " vs 直接", gusDir, " (应一致)");

  for (const [k, label] of [["本周累计加班工时", "本周加班"], ["上周累计加班工时", "上周加班"], ["双周累计工时", "双周累计"]]) {
    console.log(`${label}: 生产 非0=${count(pr, k)} 总和=${sum(pr, k).toFixed(1)} | 直接 非0=${count(dr, k)} 总和=${sum(dr, k).toFixed(1)}`);
  }

  // 逐行对比（按工号）
  const pmap = new Map(pr.map((r: any) => [r["工号"], r]));
  const dmap = new Map(dr.map((r: any) => [r["工号"], r]));
  const diffs: string[] = [];
  for (const [code, dr2] of dmap) {
    const pr2 = pmap.get(code);
    if (!pr2) { diffs.push(`${code}: 直接有生产无`); continue; }
    for (const [k, label] of [["本周累计加班工时", "本周"], ["上周累计加班工时", "上周"], ["双周累计工时", "双周"]]) {
      if (Math.abs(Number(pr2[k] || 0) - Number(dr2[k] || 0)) > 0.001) {
        diffs.push(`${code} ${label}: 生产=${pr2[k]} 直接=${dr2[k]}`);
      }
    }
  }
  console.log("\n签字工时字段逐行差异数:", diffs.length);
  diffs.slice(0, 8).forEach((d) => console.log("  ", d));
}

main().catch((e) => { console.error(e); process.exit(1); });
