import * as XLSX from 'xlsx';
import { ProcessingRow } from './types';

function extractIdsFromWorkbook(buf: Buffer): Set<string> {
  const ids = new Set<string>();
  const wb = XLSX.read(buf, { type: 'buffer' });
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[sheetName], { defval: '' });
    for (const row of rows) {
      for (const val of Object.values(row)) {
        const s = String(val ?? '').trim();
        if (/^[A-Za-z0-9]+$/.test(s)) ids.add(s);
      }
    }
  }
  return ids;
}

export function step1RemoveResigned(rows: ProcessingRow[], leaveFile?: Buffer): ProcessingRow[] {
  if (!leaveFile || leaveFile.length === 0) { console.log('[step1] 跳过'); return rows; }
  const ids = extractIdsFromWorkbook(leaveFile);
  if (ids.size === 0) return rows;
  const before = rows.length;
  const result = rows.filter((r) => !ids.has(r.employee_code));
  console.log(`[step1] 剔除离职 ${before - result.length} 人（${before} → ${result.length}）`);
  return result;
}
