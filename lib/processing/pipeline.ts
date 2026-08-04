import * as XLSX from "xlsx";
import {
  ProcessingRow,
  RawRow,
  UploadFiles,
  PipelineResult,
  ShiftDict,
} from "./types";
import { toStr, toNum, parseDate } from "./utils";
import { readSheetRows } from "./excel";
import { step1RemoveResigned } from "./step1_remove_resigned";
import { step2RemoveNotJoined } from "./step2_remove_not_joined";
import { step3Gl00Handle } from "./step3_gl00_handle";
import { step4RemoveEuHr } from "./step4_remove_eu_hr";
import { step5RemoveGus } from "./step5_remove_gus";
import { step6MatchSupplement } from "./step6_match_supplement";
import { step7MatchHours } from "./step7_match_hours";
import {
  rule0RestTime,
  rule1HubMark,
  rule2ShiftCorrect,
  rule3DailyHours,
  rule4Over8h,
  rule5SchedulePunch,
  rule6NoteOverride,
} from "./rules";
import { generateValidationReport, ValidationReport } from "./validation";

function readShiftDict(buf: Buffer): ShiftDict {
  const d: ShiftDict = {};
  try {
    const rows = readSheetRows(buf);
    for (const r of rows) {
      const n = String(r["班次名称"] ?? "").trim();
      if (!n) continue;
      d[n] = {
        rest_start: String(r["休息开始时间"] ?? "").trim(),
        rest_end: String(r["休息结束时间"] ?? "").trim(),
        work_start: String(r["上班时间"] || r["班次上班时间"] || "").trim(),
      };
    }
  } catch {}
  return d;
}

export function mapRow(raw: RawRow, wt: string): ProcessingRow | null {
  const date = parseDate(raw["考勤日期"]);
  if (!date) return null;
  const code = toStr(raw["工号"]);
  if (!code) return null;
  const l = wt === "GUS_LABOR";

  // 逐字段容错：单个单元格格式异常只影响该字段，不丢弃整行
  const R = (fn: () => unknown): unknown => {
    try {
      return fn();
    } catch {
      return "";
    }
  };
  const S = (fn: () => unknown) => String(R(fn) ?? "");
  const N = (fn: () => unknown) => toNum(R(fn));

  return {
    employee_code: code,
    date,
    employee_name: S(() => raw["姓名"] || raw["员工名称"] || ""),
    department_level3: S(() => raw["三级部门"] || raw["区域"] || ""),
    department_level4: S(() => raw["四级部门"] || raw["仓库"] || ""),
    department_level5: S(() => raw["五级部门"] || raw["组"] || ""),
    department: S(() => raw["部门"] || raw["三级部门"] || ""),
    employee_position: S(() => raw["职位"] || ""),
    shift_name: S(() => raw["班次名称"] || raw["班次"] || ""),
    shift_start: S(() => raw["班次上班时间"] || ""),
    shift_end: S(() => raw["班次下班时间"] || ""),
    first_punch: S(() => raw["首打卡时间"] || ""),
    last_punch: S(() => raw["末打卡时间"] || ""),
    punch_count: N(() => raw["班次内打卡次数"] || raw["辅助列"]),
    is_schedule_correct: S(() => raw["是否排班正确"]),
    is_scheduled: S(() => raw["是否排班"]),
    standard_punch_count: N(() => raw["标准打卡数"]),
    miss_count: N(() => raw["缺卡数"]),
    makeup_count: N(() => raw["补签数"]),
    is_over8h: S(() => raw["是否日超8H"]),
    daily_total_hours: N(() =>
      raw["每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计"] ||
      raw["时长总计"] ||
      raw["每日总工时"],
    ),
    overtime_hours: N(() => raw["日超8H"] || raw["加班工时"]),
    week_overtime_hours: N(() => raw["本周加班工时"] || raw["本周累计加班工时"]),
    last_week_overtime_hours: N(() =>
      raw["上周累计加班工时"] || raw["上周加班工时"],
    ),
    sign_hours: N(() => raw["双周加班工时"]),
    sign_report_hours: 0,
    is_hub: S(() => raw["HUB"] || "").length > 0,
    hub_status: S(() => raw["HUB"] || ""),
    note: S(() => raw["备注（GF）"] || ""),
    pending_home_office_hours: N(() => raw["居家办公合计（审批中）"]),
    rest_start: S(() => raw["休息开始时间"] || ""),
    rest_end: S(() => raw["休息结束时间"] || ""),
    rest_time: "",
    supplier_name: l ? S(() => raw["供应商名称"] || "") : "",
    worker_type_label: l ? S(() => raw["工种"] || "") : "",
    first_last_miss: l ? N(() => raw["首末缺卡数"]) : 0,
    mid_punch_1: l ? S(() => raw["中间打卡时间1"] || "") : "",
    mid_punch_2: l ? S(() => raw["中间打卡时间2"] || "") : "",
    lunch_miss: l ? N(() => raw["午休缺卡数"]) : 0,
    sign_start: l ? S(() => raw["首打卡补签时间"] || "") : "",
    sign_rest_start: l ? S(() => raw["休息开始补签时间"] || "") : "",
    sign_rest_end: l ? S(() => raw["休息结束补签时间"] || "") : "",
    sign_end: l ? S(() => raw["末打卡补签时间"] || "") : "",
    helper_col: l ? N(() => raw["辅助列"]) : 0,
    // 原表字段原样保留（注意不要和已有字段重复）
    home_office_hours: N(() => raw["居家办公工时（全）合计"]),
    daily_hours_formula: N(() =>
      raw["每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计"] ||
      raw["时长总计"] ||
      raw["每日总工时"],
    ),
    home_office_pending: N(() => raw["居家办公合计（审批中）"]),
    accumulated_hours: N(() => raw["累计总工时"]),
  };
}

function toDbRow(
  r: ProcessingRow,
  uid: string,
  wt: string,
): Record<string, unknown> {
  return {
    upload_id: uid,
    worker_type: wt,
    date: r.date,
    employee_code: r.employee_code,
    employee_name: r.employee_name,
    department_level3: r.department_level3,
    department_level4: r.department_level4,
    department_level5: r.department_level5,
    region: r.department_level3,
    shift_name: r.shift_name,
    shift_start: r.shift_start || null,
    shift_end: r.shift_end || null,
    first_punch: r.first_punch || null,
    last_punch: r.last_punch || null,
    punch_count: r.punch_count,
    is_schedule_correct: r.is_schedule_correct,
    is_scheduled: r.is_scheduled === "是",
    is_hub: r.is_hub,
    hub_status: r.hub_status || null,
    standard_punch_count: r.standard_punch_count,
    miss_count: r.miss_count,
    makeup_count: r.makeup_count,
    daily_total_hours: r.daily_total_hours,
    overtime_hours: r.overtime_hours,
    week_overtime_hours: r.week_overtime_hours,
    last_week_overtime_hours: r.last_week_overtime_hours,
    sign_hours: r.sign_hours,
    sign_report_hours: r.sign_report_hours,
    note: r.note || null,
    pending_home_office_hours: r.pending_home_office_hours,
    rest_time:
      r.rest_start && r.rest_end ? `${r.rest_start}-${r.rest_end}` : null,
    is_overtime: r.is_over8h === "是",
    supplier_name: r.supplier_name || null,
    worker_type_label: r.worker_type_label || null,
    first_last_miss: r.first_last_miss || 0,
    lunch_miss: r.lunch_miss || 0,
    mid_punch_1: r.mid_punch_1 || null,
    mid_punch_2: r.mid_punch_2 || null,
    sign_start: r.sign_start || null,
    sign_rest_start: r.sign_rest_start || null,
    sign_rest_end: r.sign_rest_end || null,
    sign_end: r.sign_end || null,
    helper_col: r.helper_col || 0,
  };
}

export async function runPipeline(
  files: UploadFiles,
  workerType: string,
  supabase: any,
  uploadId: string,
  dateStr?: string,
): Promise<PipelineResult> {
  // 解析 Excel，容错处理
  let raw: RawRow[] = [];
  try {
    raw = readSheetRows(files.file) as RawRow[];
  } catch (e) {
    console.error("[pipeline] Excel 解析失败:", e);
    throw new Error(`Excel 文件解析失败: ${(e as Error).message}`);
  }
  if (raw.length === 0) throw new Error("原始数据为空");

  // 逐行映射，跳过非法行
  const mapped: ProcessingRow[] = [];
  const errors: string[] = [];
  const freq: Record<string, number> = {};
  for (let i = 0; i < raw.length; i++) {
    try {
      const row = mapRow(raw[i], workerType);
      if (row) {
        mapped.push(row);
        freq[row.date] = (freq[row.date] || 0) + 1;
      } else {
        errors.push(`第 ${i + 1} 行: 缺少工号或考勤日期`);
      }
    } catch (e) {
      errors.push(`第 ${i + 1} 行: 解析异常 - ${(e as Error).message}`);
    }
  }

  if (mapped.length === 0) {
    throw new Error(`所有行均解析失败，共 ${errors.length} 条错误。请检查文件格式`);
  }

  // 输出解析警告
  if (errors.length > 0) {
    console.warn(
      `[pipeline] ${errors.length} 行跳过解析: ${errors.slice(0, 5).join("; ")}${
        errors.length > 5 ? `... 等${errors.length}条` : ""
      }`,
    );
  }
  let md = "",
    mf = 0;
  for (const [d, c] of Object.entries(freq)) {
    if (c > mf) {
      mf = c;
      md = d;
    }
  }
  const dd = dateStr || md || new Date().toISOString().slice(0, 10);

  // 各步骤容错：每个步骤用 try-catch 包裹，失败时跳过该步骤并记录日志
  let s1 = mapped;
  try { s1 = step1RemoveResigned(mapped, files.leave); } catch (e) { console.warn("[pipeline] 步骤1失败，跳过:", (e as Error).message); }

  let s2 = s1;
  try { s2 = step2RemoveNotJoined(s1, files.roster, dd); } catch (e) { console.warn("[pipeline] 步骤2失败，跳过:", (e as Error).message); }

  let s3 = s2;
  try { s3 = step3Gl00Handle(s2); } catch (e) { console.warn("[pipeline] 步骤3失败，跳过:", (e as Error).message); }

  let s4 = s3;
  try { s4 = step4RemoveEuHr(s3); } catch (e) { console.warn("[pipeline] 步骤4失败，跳过:", (e as Error).message); }

  let s5 = s4;
  try { s5 = step5RemoveGus(s4, files.gus_whitelist); } catch (e) { console.warn("[pipeline] 步骤5失败，跳过:", (e as Error).message); }

  let s6 = s5;
  try { s6 = step6MatchSupplement(s5, files.makeup, dd); } catch (e) { console.warn("[pipeline] 步骤6失败，跳过:", (e as Error).message); }

  let s7 = s6;
  try {
    s7 = step7MatchHours(
      s6,
      files.sign_this,
      files.sign_last,
      files.sign_biweek,
    );
  } catch (e) { console.warn("[pipeline] 步骤7失败，跳过:", (e as Error).message); }

  let cur = s7;
  const sd = files.shift_dict ? readShiftDict(files.shift_dict) : {};

  try { cur = rule0RestTime(cur, sd); } catch (e) { console.warn("[pipeline] 规则0失败，跳过:", (e as Error).message); }
  try { cur = rule1HubMark(cur); } catch (e) { console.warn("[pipeline] 规则1失败，跳过:", (e as Error).message); }
  try { cur = rule2ShiftCorrect(cur, sd); } catch (e) { console.warn("[pipeline] 规则2失败，跳过:", (e as Error).message); }
  try { cur = rule3DailyHours(cur); } catch (e) { console.warn("[pipeline] 规则3失败，跳过:", (e as Error).message); }
  try { cur = rule4Over8h(cur); } catch (e) { console.warn("[pipeline] 规则4失败，跳过:", (e as Error).message); }
  try { cur = rule5SchedulePunch(cur); } catch (e) { console.warn("[pipeline] 规则5失败，跳过:", (e as Error).message); }
  try { cur = rule6NoteOverride(cur); } catch (e) { console.warn("[pipeline] 规则6失败，跳过:", (e as Error).message); }

  // 写入 Supabase（如果配置了）+ 生成校验报告
  const resignedCount = mapped.length - s2.length; // step1 + step2
  const gusRemovedCount = s4.length - s5.length; // step5
  const validation = generateValidationReport(
    cur,
    resignedCount,
    gusRemovedCount,
  );

  if (supabase) {
    await supabase
      .from("attendance_data")
      .delete()
      .eq("date", dd)
      .eq("worker_type", workerType);
    const rows = cur.map((r) => toDbRow(r, uploadId, workerType));
    for (let i = 0; i < rows.length; i += 200) {
      const { error } = await supabase
        .from("attendance_data")
        .insert(rows.slice(i, i + 200));
      if (error) throw new Error(`写入失败: ${error.message}`);
    }
    await supabase
      .from("upload_records")
      .update({
        status: "completed",
        data_date: dd,
        rows_written: rows.length,
        completed_at: new Date().toISOString(),
      })
      .eq("id", uploadId);
  }
  return {
    total: cur.length,
    date: dd,
    worker_type: workerType,
    validation,
    rows: cur.map((r) => {
      // 格式化休息时间：HH:MM:SS → 与参考文件一致
      const fmtRest = (s: string) => s ? (s.includes(":") && s.split(":").length === 2 ? s + ":00" : s) : "";
      return {
        考勤日期: r.date.replace(/-/g, "/"),
        姓名: r.employee_name,
        工号: r.employee_code,
        三级部门: r.department_level3,
        四级部门: r.department_level4,
        五级部门: r.department_level5,
        HUB: r.hub_status || "",
        职位: r.employee_position || "",
        班次名称: r.shift_name,
        休息开始时间: fmtRest(r.rest_start),
        休息结束时间: fmtRest(r.rest_end),
        班次上班时间: r.shift_start,
        班次下班时间: r.shift_end,
        首打卡时间: r.first_punch,
        末打卡时间: r.last_punch,
        是否排班正确: r.is_schedule_correct,
        "居家办公工时（全）合计": r.home_office_hours,
        "每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计": r.daily_hours_formula,
        "居家办公合计（审批中）": r.home_office_pending,
        每日总工时计算: r.daily_total_hours,
        日超8H: r.overtime_hours,
        是否日超8H: r.is_over8h,
        累计总工时: r.accumulated_hours,
        本周加班工时: r.week_overtime_hours,
        班次内打卡次数: r.punch_count,
        是否排班: r.is_scheduled,
        标准打卡数: r.standard_punch_count,
        缺卡数: r.miss_count,
        "备注（GF）": r.note,
        补签数: r.makeup_count,
        本周累计加班工时: r.sign_report_hours,
        上周累计加班工时: r.last_week_overtime_hours,
        双周累计工时: r.sign_hours,
      };
    }),
    stats: {
      original: mapped.length,
      after_step1_remove_resigned: s1.length,
      after_step2_remove_not_joined: s2.length,
      after_step3_gl00: s3.length,
      after_step4_eu_hr: s4.length,
      after_step5_gus: s5.length,
      after_step6_supplement: s6.length,
      after_step7_hours: s7.length,
      final: cur.length,
    },
  };
}
