import { ProcessingRow, ShiftDict } from "./types";

/**
 * 规则0: 休息开始/结束时间
 * 从班次字典中匹配班次名称，填充休息开始和结束时间
 */
export function rule0RestTime(
  rows: ProcessingRow[],
  sd: ShiftDict,
): ProcessingRow[] {
  let m = 0;
  const r = rows.map((x) => {
    const s = sd[x.shift_name];
    if (s) {
      m++;
      return { ...x, rest_start: s.rest_start, rest_end: s.rest_end };
    }
    return x;
  });
  console.log(`  规则0: 休息时间 ${m}/${rows.length}`);
  return r;
}

/**
 * 规则1: HUB标记
 * 部门含 .H 或等于 EWR.G / CNO.G → is_hub = true
 */
export function rule1HubMark(rows: ProcessingRow[]): ProcessingRow[] {
  let c = 0;
  const r = rows.map((x) => {
    const d = x.department || x.department_level5 || "";
    const h = d === "EWR.G" || d === "CNO.G" || d.includes(".H");
    if (h) c++;
    return {
      ...x,
      is_hub: h || x.is_hub,
      hub_status: h ? "HUB" : x.hub_status || "",
    };
  });
  console.log(`  规则1: HUB ${c}个`);
  return r;
}

/** 解析 "HH:MM" 为分钟数 */
function pt(t: string): number | null {
  const m = t.match(/(\d{1,2}):(\d{2})/);
  return m ? parseInt(m[1]) * 60 + parseInt(m[2]) : null;
}

/**
 * 规则2: 是否排班正确（6级优先级）
 * P1: 班次为空 → "/"
 * P2: 休息日/节假日 + 首末均空 → "正确"
 * P3: 有班次 + 无末打卡 → |首打卡-班次上班时间| ≤1h → "正确", 否则 "不正确"
 * P4: 休息日/节假日 + 单边打卡 → "不正确"
 * P5: 其他 → "不正确"
 */
export function rule2ShiftCorrect(
  rows: ProcessingRow[],
  sd: ShiftDict,
): ProcessingRow[] {
  return rows.map((x) => {
    const sn = (x.shift_name || "").trim(),
      fp = (x.first_punch || "").trim(),
      lp = (x.last_punch || "").trim();
    if (!sn) return { ...x, is_schedule_correct: "/" };
    const s = sd[sn];
    const rest = s ? !s.work_start.trim() : false;
    if (rest && !fp && !lp) return { ...x, is_schedule_correct: "正确" };
    if (sn && !lp && s?.work_start) {
      const d1 = pt(fp),
        d2 = pt(s.work_start);
      if (d1 != null && d2 != null && Math.abs(d1 - d2) <= 60)
        return { ...x, is_schedule_correct: "正确" };
      return { ...x, is_schedule_correct: "不正确" };
    }
    if (rest && (fp || lp)) return { ...x, is_schedule_correct: "不正确" };
    return { ...x, is_schedule_correct: "不正确" };
  });
}

/**
 * 规则3: 每日总工时
 * 总工时 = 原公式值 + 居家办公合计（审批中）
 */
export function rule3DailyHours(rows: ProcessingRow[]): ProcessingRow[] {
  return rows.map((x) => {
    const h = x.daily_total_hours;
    const p = x.pending_home_office_hours;
    return { ...x, daily_total_hours: (Number(h) || 0) + (Number(p) || 0) };
  });
}

/**
 * 规则4: 是否日超8H
 * 每日总工时 > 8 → 是, 否则 → 否
 */
export function rule4Over8h(rows: ProcessingRow[]): ProcessingRow[] {
  let c = 0;
  const r = rows.map((x) => {
    const h = Number(x.daily_total_hours) || 0;
    const o = Math.max(0, h - 8) > 0;
    if (o) c++;
    return { ...x, is_over8h: o ? "是" : "否", overtime_hours: o ? h - 8 : 0 };
  });
  console.log(`  规则4: 超8H ${c}人`);
  return r;
}

/**
 * 规则5: 是否排班 + 标准打卡数 + 缺卡数
 * - 无班次 → 排班=否, 标准打卡=0, 缺卡=0
 * - 有班次 → 排班=是, 标准打卡=2, 缺卡=max(0, 2-打卡次数)
 */
export function rule5SchedulePunch(rows: ProcessingRow[]): ProcessingRow[] {
  return rows.map((x) => {
    if (!(x.shift_name || "").trim())
      return {
        ...x,
        is_scheduled: "否",
        standard_punch_count: 0,
        miss_count: 0,
      };
    const p = Number(x.punch_count) || 0;
    return {
      ...x,
      is_scheduled: "是",
      standard_punch_count: 2,
      miss_count: Math.max(0, 2 - p),
    };
  });
}

/**
 * 规则6: 备注覆盖（在规则5之后执行）
 * 备注含"请假/出差/居家办公"时：
 * - ≥8H → 标准打卡=0, 缺卡=0, 排班正确=正确
 * - ≥4H → 标准打卡=2, 打卡<2→排班正确=不正确
 */
export function rule6NoteOverride(rows: ProcessingRow[]): ProcessingRow[] {
  let c8 = 0,
    c4 = 0;
  return rows.map((x) => {
    const n = (x.note || "").trim();
    if (!n || !["请假", "出差", "居家办公"].some((k) => n.includes(k)))
      return x;
    const m = n.match(/(\d+(?:\.\d+)?)\s*[Hh时]/);
    const h = m ? parseFloat(m[1]) : 0;
    if (h >= 8) {
      c8++;
      return {
        ...x,
        standard_punch_count: 0,
        miss_count: 0,
        is_schedule_correct: "正确",
      };
    }
    if (h >= 4) {
      c4++;
      return {
        ...x,
        standard_punch_count: 2,
        is_schedule_correct:
          (Number(x.punch_count) || 0) < 2 ? "不正确" : x.is_schedule_correct,
      };
    }
    return x;
  });
}
