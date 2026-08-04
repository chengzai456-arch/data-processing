import * as XLSX from "xlsx";
import { ProcessingRow } from "./types";
import { readWorkbook } from "./excel";

/**
 * GUS 白名单工号集合
 *
 * 只读取「工号 / 员工编号 / ID」类列，避免把姓名、部门等任意
 * 纯字母数字文本误当成工号剔除（原实现扫描了所有单元格，误剔风险高）。
 * 兼容多 sheet，兼容首行无表头（取第一列）的情况。
 */
const ID_HEADER = /^(工号|员工工号|员工编号|员工号|工号编码|员工ID|工号ID|编号|ID|employee\s*id|emp\s*code)$/i;

function normalizeCode(v: unknown): string {
  const s = String(v ?? "").trim();
  return /^[A-Za-z0-9]+$/.test(s) ? s : "";
}

export function gusIds(buf: Buffer): Set<string> {
  const ids = new Set<string>();
  const wb = readWorkbook(buf);
  for (const sn of wb.SheetNames || []) {
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    const rows = (XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
      defval: "",
    }) as Record<string, unknown>[]) || [];
    if (rows.length === 0) continue;

    const headers = Object.keys(rows[0] ?? {});
    // 优先精确匹配工号类列名
    const idCol = headers.find((h) => ID_HEADER.test(h.trim()));
    if (idCol) {
      for (const r of rows) {
        const c = normalizeCode(r[idCol]);
        if (c) ids.add(c);
      }
      continue;
    }
    // 无工号列：回退到第一列（部分白名单文件无表头）
    const firstCol = headers[0];
    if (firstCol) {
      for (const r of rows) {
        const c = normalizeCode(r[firstCol]);
        if (c) ids.add(c);
      }
    }
  }
  return ids;
}

export function step5RemoveGus(
  rows: ProcessingRow[],
  f?: Buffer,
): ProcessingRow[] {
  if (!f || f.length === 0) return rows;
  const ids = gusIds(f);
  if (ids.size === 0) return rows;
  const b = rows.length;
  const r = rows.filter((x) => !ids.has(x.employee_code));
  console.log(`[step5] 剔除GUS白名单 ${b - r.length}人（${b} → ${r.length}）`);
  return r;
}
