import { ProcessingRow } from "./types";
import { readSheetRows } from "./excel";

/**
 * 步骤1: 剔除离职人员（精确匹配 SKILL clean_data.py step1）
 * - 解析 最后工作日 列
 * - 如果 审批状态 列存在，仅保留 审批中/已完成/转交
 * - 仅剔除最后工作日 < 标准日期 且 审批状态符合条件的工号
 */
function parseDate(v: unknown): string | null {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s) return null;
  // Try common formats
  for (const fmt of [/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/, /^(\d{4})-(\d{1,2})-(\d{1,2})$/]) {
    const m = s.match(fmt);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  }
  // Excel serial date number
  const n = Number(s);
  if (!isNaN(n) && n > 40000) {
    const d = new Date((n - 25569) * 86400 * 1000);
    return d.toISOString().slice(0, 10);
  }
  return null;
}

export function step1RemoveResigned(
  rows: ProcessingRow[],
  leaveFile?: Buffer,
): ProcessingRow[] {
  if (!leaveFile || leaveFile.length === 0) {
    console.log("[step1] 跳过");
    return rows;
  }

  // 解析离职文件
  const leaveRows = readSheetRows(leaveFile);
  if (leaveRows.length === 0) return rows;

  // 确定标准日期（取考勤数据中最多的日期）
  const dateCount: Record<string, number> = {};
  for (const r of rows) dateCount[r.date] = (dateCount[r.date] || 0) + 1;
  let standardDate = "";
  let maxCount = 0;
  for (const [d, c] of Object.entries(dateCount)) {
    if (c > maxCount) { maxCount = c; standardDate = d; }
  }
  if (!standardDate) return rows;

  // 检查是否有审批状态列
  const headers = leaveRows.length > 0 ? Object.keys(leaveRows[0]) : [];
  const hasStatus = headers.some((h) => h.includes("审批状态"));

  // 构建剔除集合
  const toRemove = new Set<string>();
  for (const r of leaveRows) {
    const empCode = String(r["工号"] ?? "").trim();
    if (!empCode || !/^[A-Za-z0-9]+$/.test(empCode)) continue;

    const lastWorkDay = parseDate(r["最后工作日"]);
    if (!lastWorkDay) continue;
    if (lastWorkDay >= standardDate) continue; // 最后工作日 >= 标准日期 → 保留

    if (hasStatus) {
      const status = String(r["审批状态"] ?? "").trim();
      const isValid = ["审批中", "已完成", "转交"].some((s) => status.includes(s));
      if (!isValid) continue; // 状态不符合 → 保留
    }

    toRemove.add(empCode);
  }

  if (toRemove.size === 0) return rows;
  const before = rows.length;
  const result = rows.filter((r) => !toRemove.has(r.employee_code));
  console.log(
    `[step1] 剔除离职 ${before - result.length} 人（${before} → ${result.length}）`,
  );
  return result;
}
