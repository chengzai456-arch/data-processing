import { ProcessingRow } from './types';

const GL00_W = new Set(['GL000001','GL000002','GL000003','GL000004','GL000005','GL000006']);

export function step3Gl00Handle(rows: ProcessingRow[]): ProcessingRow[] {
  const b = rows.length;
  const r1 = rows.filter((r) => r.employee_code !== 'GL502563');
  const r2 = r1.filter((r) => !(r.employee_code.startsWith('GL00') && !GL00_W.has(r.employee_code)));
  console.log(`[step3] 剔除 GL502563(${b-r1.length}) + GL00非白名单${r1.length-r2.length}（${b} → ${r2.length}）`);
  return r2;
}
