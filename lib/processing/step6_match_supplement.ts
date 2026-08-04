import { ProcessingRow } from "./types";
import { readSheetRows } from "./excel";

const VS = new Set(["已完成", "审批中", "转交"]);

export function step6MatchSupplement(
  rows: ProcessingRow[],
  f?: Buffer,
  sd?: string,
): ProcessingRow[] {
  if (!f || f.length === 0) return rows.map((r) => ({ ...r, makeup_count: 0 }));
  let s = readSheetRows(f);
  if (s.length === 0) return rows.map((r) => ({ ...r, makeup_count: 0 }));
  // 审批状态精确匹配（原为包含匹配，避免"审批中X"等误判）
  if ("审批状态" in s[0])
    s = s.filter((r) => VS.has(String(r["审批状态"] ?? "").trim()));
  const dc = Object.keys(s[0] ?? {}).find((k) => k.includes("日期"));
  if (dc && sd)
    s = s.filter((r) => {
      const x = r[dc];
      if (x == null) return 0;
      const m = String(x)
        .trim()
        .match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
      if (m)
        return (
          `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}` === sd
        );
      return 0;
    });
  const m = new Map<string, number>();
  s.forEach((r) => {
    const e = String(r["工号"] ?? "").trim();
    if (e) m.set(e, (m.get(e) || 0) + 1);
  });
  return rows.map((r) => ({ ...r, makeup_count: m.get(r.employee_code) || 0 }));
}
