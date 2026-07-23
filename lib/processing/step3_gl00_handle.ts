import { ProcessingRow } from "./types";

/**
 * 步骤3: GL00工号处理
 * - 剔除 GL502563
 * - 只保留白名单内的 GL00 工号（其余 GL00 全部剔除）
 *
 * 如果你要修改白名单，直接改下面的 GL00_W 集合即可
 */
const GL00_W = new Set([
  "GL000001",
  "GL000002",
  "GL000003",
  "GL000004",
  "GL000005",
  "GL000006",
]);

export function step3Gl00Handle(rows: ProcessingRow[]): ProcessingRow[] {
  const b = rows.length;

  // 第一步：剔除 GL502563
  const r1 = rows.filter((r) => r.employee_code !== "GL502563");

  // 第二步：剔除不在白名单内的 GL00 工号
  const r2 = r1.filter(
    (r) =>
      !(r.employee_code.startsWith("GL00") && !GL00_W.has(r.employee_code)),
  );

  console.log(
    `[step3] 剔除 GL502563(${b - r1.length}) + GL00非白名单${r1.length - r2.length}（${b} → ${r2.length}）`,
  );
  return r2;
}
