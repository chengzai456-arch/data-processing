import * as XLSX from 'xlsx';
import { ProcessingRow } from './types';

export function step2RemoveNotJoined(rows: ProcessingRow[], rosterFile?: Buffer, standardDate?: string): ProcessingRow[] {
  if (!rosterFile || rosterFile.length === 0) { console.log('[step2] 跳过'); return rows; }
  const wb = XLSX.read(rosterFile, { type: 'buffer' });
  const rosterRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[wb.SheetNames[0]], { defval: '' });
  if (rosterRows.length === 0) return rows;
  const ref = new Date(standardDate || new Date().toISOString().slice(0, 10));
  const notJoined = new Set<string>();
  for (const r of rosterRows) {
    const emp = String(r['工号'] ?? '').trim();
    const entry = r['入职日期'];
    if (!emp || !/^[A-Za-z0-9]+$/.test(emp) || !entry) continue;
    let d: Date | null = null;
    if (typeof entry === 'number') { const dd = new Date((entry - 25569) * 86400 * 1000); if (!isNaN(dd.getTime())) d = dd; }
    else { const m = String(entry).trim().match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/); if (m) d = new Date(`${m[1]}-${m[2]}-${m[3]}`); }
    if (d && d > ref) notJoined.add(emp);
  }
  if (notJoined.size === 0) return rows;
  const before = rows.length;
  const result = rows.filter((r) => !notJoined.has(r.employee_code));
  console.log(`[step2] 剔除未入职 ${before - result.length} 人（${before} → ${result.length}）`);
  return result;
}
