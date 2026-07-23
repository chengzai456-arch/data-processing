import * as XLSX from "xlsx";
import { ProcessingRow } from "./types";

function gusIds(buf: Buffer): Set<string> {
  const ids = new Set<string>();
  const wb = XLSX.read(buf, { type: "buffer" });
  for (const sn of wb.SheetNames) {
    for (const row of XLSX.utils.sheet_to_json<Record<string, unknown>>(
      wb.Sheets[sn],
      { defval: "" },
    )) {
      for (const v of Object.values(row)) {
        const s = String(v ?? "").trim();
        if (/^[A-Za-z0-9]+$/.test(s)) ids.add(s);
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
