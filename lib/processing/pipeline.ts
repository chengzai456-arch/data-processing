import * as XLSX from "xlsx";
import {
  ProcessingRow,
  RawRow,
  UploadFiles,
  PipelineResult,
  ShiftDict,
} from "./types";
import { toStr, toNum, parseDate } from "./utils";
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
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      XLSX.read(buf, { type: "buffer" }).Sheets[
        XLSX.read(buf, { type: "buffer" }).SheetNames[0]
      ],
      { defval: "" },
    );
    for (const r of rows) {
      const n = String(r["班次名称"] ?? "").trim();
      if (!n) continue;
      d[n] = {
        rest_start: String(r["休息开始"] ?? "").trim(),
        rest_end: String(r["休息结束"] ?? "").trim(),
        work_start: String(r["上班时间"] || r["班次上班时间"] || "").trim(),
      };
    }
  } catch {}
  return d;
}

function mapRow(raw: RawRow, wt: string): ProcessingRow | null {
  const date = parseDate(raw["考勤日期"]);
  if (!date) return null;
  const code = toStr(raw["工号"]);
  if (!code) return null;
  const l = wt === "GUS_LABOR";
  return {
    employee_code: code,
    date,
    employee_name: toStr(raw["姓名"] || raw["员工名称"] || ""),
    department_level3: toStr(raw["三级部门"] || raw["区域"] || ""),
    department_level4: toStr(raw["四级部门"] || raw["仓库"] || ""),
    department_level5: toStr(raw["五级部门"] || raw["组"] || ""),
    department: toStr(raw["部门"] || raw["三级部门"] || ""),
    shift_name: toStr(raw["班次名称"] || raw["班次"] || ""),
    shift_start: toStr(raw["班次上班时间"] || ""),
    shift_end: toStr(raw["班次下班时间"] || ""),
    first_punch: toStr(raw["首打卡时间"] || ""),
    last_punch: toStr(raw["末打卡时间"] || ""),
    punch_count: toNum(raw["班次内打卡次数"] || raw["辅助列"]),
    is_schedule_correct: toStr(raw["是否排班正确"]),
    is_scheduled: toStr(raw["是否排班"]),
    standard_punch_count: toNum(raw["标准打卡数"]),
    miss_count: toNum(raw["缺卡数"]),
    makeup_count: toNum(raw["补签数"]),
    is_over8h: toStr(raw["是否日超8H"]),
    daily_total_hours: toNum(
      raw["每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计"] ||
        raw["时长总计"] ||
        raw["每日总工时"],
    ),
    overtime_hours: toNum(raw["日超8H"] || raw["加班工时"]),
    week_overtime_hours: toNum(raw["本周加班工时"] || raw["本周累计加班工时"]),
    last_week_overtime_hours: toNum(
      raw["上周加班工时"] || raw["上周累计加班工时"],
    ),
    sign_hours: toNum(raw["双周加班工时"]),
    sign_report_hours: 0,
    is_hub: toStr(raw["HUB"] || "").length > 0,
    hub_status: toStr(raw["HUB"] || ""),
    note: toStr(raw["备注（GF）"] || ""),
    pending_home_office_hours: toNum(raw["居家办公合计（审批中）"]),
    rest_start: toStr(raw["休息开始时间"] || ""),
    rest_end: toStr(raw["休息结束时间"] || ""),
    rest_time: "",
    supplier_name: l ? toStr(raw["供应商名称"] || "") : "",
    worker_type_label: l ? toStr(raw["工种"] || "") : "",
    first_last_miss: l ? toNum(raw["首末缺卡数"]) : 0,
    mid_punch_1: l ? toStr(raw["中间打卡时间1"] || "") : "",
    mid_punch_2: l ? toStr(raw["中间打卡时间2"] || "") : "",
    lunch_miss: l ? toNum(raw["午休缺卡数"]) : 0,
    sign_start: l ? toStr(raw["首打卡补签时间"] || "") : "",
    sign_rest_start: l ? toStr(raw["休息开始补签时间"] || "") : "",
    sign_rest_end: l ? toStr(raw["休息结束补签时间"] || "") : "",
    sign_end: l ? toStr(raw["末打卡补签时间"] || "") : "",
    helper_col: l ? toNum(raw["辅助列"]) : 0,
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
    raw = XLSX.utils.sheet_to_json<RawRow>(
      XLSX.read(files.file, { type: "buffer" }).Sheets[
        XLSX.read(files.file, { type: "buffer" }).SheetNames[0]
      ],
      { defval: "" },
    );
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
    rows: cur.map((r) => ({
      工号: r.employee_code,
      姓名: r.employee_name,
      考勤日期: r.date,
      三级部门: r.department_level3,
      四级部门: r.department_level4,
      五级部门: r.department_level5,
      班次名称: r.shift_name,
      首打卡: r.first_punch,
      末打卡: r.last_punch,
      打卡次数: r.punch_count,
      排班正确: r.is_schedule_correct,
      是否排班: r.is_scheduled,
      标准打卡: r.standard_punch_count,
      缺卡: r.miss_count,
      补签: r.makeup_count,
      每日工时: r.daily_total_hours,
      日超8H: r.is_over8h,
      加班工时: r.overtime_hours,
      本周加班: r.week_overtime_hours,
      上周加班: r.last_week_overtime_hours,
      双周加班: r.sign_hours,
      HUB: r.hub_status,
      休息开始: r.rest_start,
      休息结束: r.rest_end,
      备注: r.note,
    })),
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
