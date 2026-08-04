import { toNum, toStr, parseDate } from "../lib/processing/utils";
import { gusIds } from "../lib/processing/step5_remove_gus";
import * as XLSX from "xlsx";

let pass = 0, fail = 0;
function eq(desc: string, got: unknown, want: unknown) {
  const ok = got === want;
  if (ok) pass++; else { fail++; console.log("✗", desc, "got=", got, "want=", want); }
}

// toNum
eq("千位分隔", toNum("1,234.5"), 1234.5);
eq("百分比", toNum("85.5%"), 85.5);
eq("货币", toNum("¥1,200"), 1200);
eq("中文小时", toNum("40小时"), 40);
eq("中文分钟→小时", toNum("90分钟"), 1.5);
eq("h单位", toNum("8h"), 8);
eq("括号负数", toNum("(120)"), -120);
eq("缺失-", toNum("-"), 0);
eq("缺失--", toNum("--"), 0);
eq("时间错填", toNum("08:30"), 510);
eq("正常", toNum("12.5"), 12.5);
eq("空", toNum(""), 0);
eq("nan", toNum("nan"), 0);

// parseDate
eq("斜杠日期", parseDate("2026/7/22"), "2026-07-22");
eq("横杠日期", parseDate("2026-07-22"), "2026-07-22");
eq("点日期", parseDate("2026.7.22"), "2026-07-22");
eq("中文日期", parseDate("2026年7月22日"), "2026-07-22");
eq("excel序列", parseDate(45394), "2024-04-12");
eq("序列字符串", parseDate("45394"), "2024-04-12");
eq("bad", parseDate("abc"), null);
eq("空", parseDate(""), null);

// gusIds: 只扫工号列，姓名/部门里的纯字母数字不误判
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  ["工号", "姓名", "部门"],
  ["GL001", "张三", "TD"],
  ["GL002", "PETER", "F201"],
  ["", "WANG", "LOGISTICS"],
]);
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
const buf = XLSX.write(wb, { type: "buffer" });
const ids = gusIds(buf);
eq("gusIds含GL001", ids.has("GL001"), true);
eq("gusIds含GL002", ids.has("GL002"), true);
eq("gusIds不含姓名PETER", ids.has("PETER"), false);
eq("gusIds不含部门F201", ids.has("F201"), false);
eq("gusIds不含LOGISTICS", ids.has("LOGISTICS"), false);
eq("gusIds不含空格空值", ids.has(""), false);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);
