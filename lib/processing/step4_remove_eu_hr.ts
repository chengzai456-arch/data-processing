import { ProcessingRow } from './types';

export function step4RemoveEuHr(rows: ProcessingRow[]): ProcessingRow[] {
  const b = rows.length;
  const r = rows.filter((x) => x.department !== 'EU人力资源部');
  console.log(`[step4] 剔除 EU人力资源部 ${b - r.length} 人（${b} → ${r.length}）`);
  return r;
}
