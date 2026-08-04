import * as fs from "fs";
import * as XLSX from "xlsx";
import { runPipeline } from "../lib/processing/pipeline";

const DATA_DIR = "D:/Documents/Downloads/722日晾晒数据";
const load = (n: string) => fs.readFileSync(`${DATA_DIR}/${n}`);
const mainKeep = ["工号","考勤日期","姓名","三级部门","四级部门","五级部门","部门","班次名称",
  "班次上班时间","班次下班时间","首打卡时间","末打卡时间","班次内打卡次数","是否排班正确",
  "每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计","日超8H","是否日超8H",
  "本周加班工时","上周累计加班工时","HUB","备注（GF）","居家办公合计（审批中）","补签数",
  "休息开始时间","休息结束时间","职位","居家办公工时（全）合计","累计总工时","最后工作日","审批状态","入职日期","补签日期","加班合计",
  "加班时间段开始1","加班时间段结束1","加班时间段开始2","加班时间段结束2","OT1.5合计",
  "OT2.0合计","计薪出勤时长合计（REG）","每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)"];

function compressMain(buf: Buffer): Record<string, unknown>[] {
  const wb = XLSX.read(buf, { type: "buffer" });
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[wb.SheetNames[0]], { defval: "" });
  return rows.map((r: any) => {
    const o: any = {};
    for (const f of mainKeep) if (f in r) o[f] = r[f];
    return o;
  });
}
function j2b(rows: Record<string, unknown>[]): Buffer {
  const h = Object.keys(rows[0]);
  const ws = XLSX.utils.aoa_to_sheet([h, ...rows.map((r) => h.map((k) => r[k] ?? ""))]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "S");
  return Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
}

async function main() {
  const prod: any = {
    file: j2b(compressMain(load("722GUS+每日打卡工时推送模版 (51).xlsx"))),
    leave: j2b(compressMain(load("722离职流程 (62).xlsx"))),
    roster: j2b(compressMain(load("722花名册 (69).xlsx"))),
    shift_dict: j2b(compressMain(load("班次.xlsx"))),
    makeup: j2b(compressMain(load("722补签管理 (77).xlsx"))),
    gus_whitelist: load("722GUS白名单人员 (3).xlsx"),  // 直接传原始多sheet
  };
  // 只验证主文件输出列，不需要签字报表
  const direct: any = {
    file: load("722GUS+每日打卡工时推送模版 (51).xlsx"),
    leave: load("722离职流程 (62).xlsx"),
    roster: load("722花名册 (69).xlsx"),
    shift_dict: load("班次.xlsx"),
    makeup: load("722补签管理 (77).xlsx"),
    gus_whitelist: load("722GUS白名单人员 (3).xlsx"),
  };
  const pr = await runPipeline(prod, "GUS", null, "p", "2026-07-22");
  const dr = await runPipeline(direct, "GUS", null, "d", "2026-07-22");
  const prow = pr.rows!;
  const drow = dr.rows!;
  for (const col of ["职位", "居家办公工时（全）合计", "累计总工时"]) {
    const pNonEmpty = prow.filter((r: any) => String(r[col] ?? "").trim() !== "" && Number(r[col] || 0) !== 0).length;
    const dNonEmpty = drow.filter((r: any) => String(r[col] ?? "").trim() !== "" && Number(r[col] || 0) !== 0).length;
    console.log(`${col}: 生产压缩 非空=${pNonEmpty} | 直接 非空=${dNonEmpty}`);
  }
  // 逐行对比这3列
  const pm = new Map(prow.map((r: any) => [r["工号"], r]));
  let diff = 0;
  for (const r of drow as any[]) {
    const p = pm.get(r["工号"]);
    for (const col of ["职位", "居家办公工时（全）合计", "累计总工时"]) {
      if (String(p?.[col] ?? "") !== String(r[col] ?? "")) { diff++; if (diff <= 3) console.log(`DIFF ${r["工号"]} ${col}: 生产=${p?.[col]} 直接=${r[col]}`); }
    }
  }
  console.log("3列逐行差异:", diff);
}
main().catch((e) => { console.error(e); process.exit(1); });
