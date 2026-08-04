import { ProcessingRow } from "./types";
import { readSheetRows } from "./excel";

const HKW = ["每日总工时", "总工时", "工时", "时长", "小时"];
const SUB = 40;

function tn(v: unknown): number {
  const s = String(v ?? "").trim();
  if (!s || ["nan", "NaT", "None", "NaN", "nat", "", "/"].includes(s)) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

/** 工号规范化：去首尾空白后仅保留字母数字 */
function empCode(v: unknown): string {
  const s = String(v ?? "").trim();
  return /^[A-Za-z0-9]+$/.test(s) ? s : "";
}

/**
 * 签字报表处理
 *
 * 差异1修复: 只保留姓名列含"(合计)"的行(蓝色合计行), 其他行删除
 *   比SKILL更精确: SKILL 用 openpyxl 读填充色, 但颜色可能丢失
 *   用"(合计)"文本标记100%准确
 *
 * 差异2修复: 合计行剔除后, 每人只留1行(周汇总)
 *   直接取该行的工时值作为周累计, 不做逐行sum
 *   避免与明细行重复计算
 */
function psr(rows: Record<string, unknown>[], sub: boolean, label: string): Record<string, number> {
  const h = rows.length > 0 ? Object.keys(rows[0]) : [];
  // 优先匹配完整列名，再按关键词匹配
  const dailyCol = h.find((c) => c.includes("每日总工时"));
  const hc = dailyCol || h.find((c) => HKW.some((k) => c.includes(k)));
  const ec = h.find((c) => c.includes("工号"));
  const nc = h.find((c) => c.includes("姓名"));
  if (!hc || !ec) {
    console.log(`  [step7-${label}] 缺少工时列或工号列`);
    return {};
  }

  // 差异1: 只保留姓名含"(合计)"的行
  // 没有姓名列时无法识别合计行，返回空并告警（宁可明确为 0，也不静默取错明细行）
  if (!nc) {
    console.log(`  [step7-${label}] ⚠️ 缺少姓名列，无法识别(合计)行，签字工时置 0`);
    return {};
  }
  const filtered = rows.filter((r) => {
    const name = String(r[nc] ?? "").trim();
    return name.includes("(合计)");
  });

  console.log(`  [step7-${label}] 原始${rows.length}行, 合计行${filtered.length}行`);

  // 差异2: 合计行中每个员工有且只有1行(周累计值)
  // 按工号去重, 取第一条；工号做 trim + 字母数字校验，避免空格/脏数据导致重复
  const seen = new Set<string>();
  const deduped: Record<string, unknown>[] = [];
  for (const row of filtered) {
    const e = empCode(row[ec]);
    if (!e || seen.has(e)) continue;
    seen.add(e);
    deduped.push(row);
  }

  console.log(`  [step7-${label}] 去重后${deduped.length}个员工`);

  // 每人只有1行(周累计工时), 直接取 -40 floor 0
  const r: Record<string, number> = {};
  for (const row of deduped) {
    const e = empCode(row[ec]);
    if (!e) continue;
    const h = tn(row[hc]);
    if (sub) {
      r[e] = Math.max(0, Math.round((h - SUB) * 100) / 100);
    } else {
      // 双周不减40
      r[e] = Math.round(h * 100) / 100;
    }
  }
  return r;
}

export function step7MatchHours(
  rows: ProcessingRow[],
  sThis?: Buffer,
  sLast?: Buffer,
  sBi?: Buffer,
): ProcessingRow[] {
  function readSheet(buf: Buffer): Record<string, unknown>[] {
    try {
      return readSheetRows(buf);
    } catch {
      return [];
    }
  }

  const wData = sThis ? readSheet(sThis) : [];
  const lData = sLast ? readSheet(sLast) : [];
  const bData = sBi ? readSheet(sBi) : [];

  const w = psr(wData, true, "本周");
  const l = psr(lData, true, "上周");
  const b = psr(bData, false, "双周");

  console.log(
    `[step7] 签字报表: 本周${Object.keys(w).length}个/上周${Object.keys(l).length}个/双周${Object.keys(b).length}个`,
  );

  return rows.map((r) => ({
    ...r,
    // 覆盖语义（与 workbench 参考实现一致）：签字报表有值用签字报表，否则保留原始
    sign_report_hours: (w[r.employee_code] ?? r.sign_report_hours ?? 0),
    week_overtime_hours: r.week_overtime_hours,
    last_week_overtime_hours: (l[r.employee_code] ?? r.last_week_overtime_hours ?? 0),
    sign_hours: (b[r.employee_code] ?? r.sign_hours ?? 0),
  }));
}
