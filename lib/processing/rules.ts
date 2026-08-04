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
 * 四级/五级部门含 .H 或等于 EWR.G / CNO.G → is_hub = hub
 * SKILL: d5+d4 检查含.H或等于EWR.G/CNO.G
 */
export function rule1HubMark(rows: ProcessingRow[]): ProcessingRow[] {
  let c = 0;
  const r = rows.map((x) => {
    const d4 = x.department_level4 || "";
    const d5 = x.department_level5 || "";
    const h = d5.includes(".H") || d4.includes(".H") || d5 === "EWR.G" || d5 === "CNO.G";
    if (h) c++;
    return { ...x, is_hub: h, hub_status: h ? "hub" : "" };
  });
  console.log(`  规则1: HUB ${c}个`);
  return r;
}

/** 解析 "HH:MM" 为分钟数 */
function pt(t: string): number | null {
  const m = t.match(/(\d{1,2}):(\d{2})/);
  return m ? parseInt(m[1]) * 60 + parseInt(m[2]) : null;
}

/** 解析备注时间段，返回 {startMin, endMin} | null */
// SKILL: 需要请假类型关键词前缀 + 时间段
const WORK_PERIOD_RE = /(?:居家办公|公出|出差|病假|年假|无薪病假|事假|调休|婚假|产假|陪产假|丧假|工伤假).*?(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/;
const PERIOD_FALLBACK_RE = /(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/;

function parseWorkPeriod(n: string): { startMin: number; endMin: number } | null {
  // 优先匹配关键词前缀+时间段（精确匹配SKILL）
  let m = n.match(WORK_PERIOD_RE);
  if (!m) {
    // 回退: 纯时间匹配（备用）
    m = n.match(PERIOD_FALLBACK_RE);
    if (!m) return null;
  }
  try {
    const [sh, sm] = m[1].split(":").map(Number);
    const [eh, em] = m[2].split(":").map(Number);
    let start = sh * 60 + sm;
    let end = eh * 60 + em;
    if (end <= start) end += 24 * 60; // 跨日
    return { startMin: start, endMin: end };
  } catch {
    return null;
  }
}

/**
 * 规则2: 是否排班正确（SKILL 7级优先级，精确匹配 add_metrics.py get_correct）
 *
 * P1: 班次名称空 → "/"
 * P2: 休息日/节假日 + 首末均空 → "正确"
 * P3: 班次不空 + 末打卡空 + |首打卡-班次上班|≤1h → "正确"
 * P4: 休息日/节假日 + 单边打卡 → "不正确"
 * P5: 非休息日 + 首末均有 + |首打卡-班次上班|≤1h → "正确"
 * P6: 其他 → "不正确"
 */
export function rule2ShiftCorrect(
  rows: ProcessingRow[],
  sd: ShiftDict,
): ProcessingRow[] {
  return rows.map((x) => {
    const sn = (x.shift_name || "").trim();
    const fp = (x.first_punch || "").trim();
    const lp = (x.last_punch || "").trim();
    const startT = (x.shift_start || "").trim();

    // P1: 班次为空 → "/"
    if (!sn) return { ...x, is_schedule_correct: "/" };

    const isRest = sn.includes("TY_休息日") || sn.includes("TY_美国节假日");
    const hasFirst = fp.length > 0;
    const hasLast = lp.length > 0;

    // P2: 休息日/节假日 + 首末均空 → "正确"
    if (isRest && !hasFirst && !hasLast) return { ...x, is_schedule_correct: "正确" };

    // P3: 班次不空 + 末打卡空 + |首打卡-班次上班| ≤1h → "正确"
    if (sn && !hasLast) {
      const diff = pt(fp) != null && pt(startT) != null ? Math.abs(pt(fp)! - pt(startT)!) : null;
      if (diff != null && diff <= 60) return { ...x, is_schedule_correct: "正确" };
      return { ...x, is_schedule_correct: "不正确" };
    }

    // P4: 休息日/节假日 + 单边打卡 → "不正确"
    if (isRest && ((hasFirst && !hasLast) || (!hasFirst && hasLast)))
      return { ...x, is_schedule_correct: "不正确" };

    // P5: 非休息日 + 首末均有 + |首打卡-班次上班| ≤1h → "正确"
    if (!isRest && hasFirst && hasLast) {
      const diff = pt(fp) != null && pt(startT) != null ? Math.abs(pt(fp)! - pt(startT)!) : null;
      if (diff != null && diff <= 60) return { ...x, is_schedule_correct: "正确" };
      return { ...x, is_schedule_correct: "不正确" };
    }

    // P6: 其他 → "不正确"
    return { ...x, is_schedule_correct: "不正确" };
  });
}

/**
 * 规则3: 每日总工时 = 原公式值 + 居家办公合计（审批中）
 */
export function rule3DailyHours(rows: ProcessingRow[]): ProcessingRow[] {
  return rows.map((x) => {
    const h = x.daily_total_hours;
    const p = x.pending_home_office_hours;
    return { ...x, daily_total_hours: (Number(h) || 0) + (Number(p) || 0) };
  });
}

/**
 * 规则4: 日超8H = max(0, 每日总工时计算 - 8)
 */
export function rule4Over8h(rows: ProcessingRow[]): ProcessingRow[] {
  let c = 0;
  const r = rows.map((x) => {
    const h = Number(x.daily_total_hours) || 0;
    const ov = Math.max(0, Math.round((h - 8) * 100) / 100);
    const flag = ov > 0 ? "是" : "否";
    if (ov > 0) c++;
    return { ...x, is_over8h: flag, overtime_hours: ov };
  });
  console.log(`  规则4: 超8H ${c}人`);
  return r;
}

/**
 * 规则5: 标准打卡数（SKILL get_standard_count 精确匹配）
 *
 * 1. 未排班(班次空) → 0
 * 2. 休息日/节假日 → 0
 * 3. 备注空 → 4
 * 4. 解析备注时间段，与班次/休息时间交叉计算：
 *    - 时间段完全吻合班次起止 → 0
 *    - 排班正确=正确 + 有休息时间 →
 *      完全在休息区间 → 0, 与休息交叉 → 3, 包含休息区间 → 2, 其他 → 4
 *    - 排班正确=不正确 →
 *      >=7h → 0, >=4h → 2, 其他 → 4
 */
export function rule5SchedulePunch(rows: ProcessingRow[]): ProcessingRow[] {
  return rows.map((x) => {
    const sn = (x.shift_name || "").trim();
    const note = (x.note || "").trim();
    const p = Number(x.punch_count) || 0;
    const isScheduled = sn.length > 0;
    const isRest = sn.includes("TY_休息日") || sn.includes("TY_美国节假日");

    let stdCount = 0;

    // 1: 未排班 → 0
    if (!isScheduled) stdCount = 0;
    // 2: 休息日/节假日 → 0
    else if (isRest) stdCount = 0;
    // 3: 备注空 → 4
    else if (!note) stdCount = 4;
    // 3.5: SKILL v27: 过期备注(仅[已废弃]/[已撤回]且无[审批中]/[已完成]) → 视为无备注 → 4
    else if (isNoteExpired(note)) stdCount = 4;
    else {
      // 4: 有备注 → 解析时间段，与班次计算
      const period = parseWorkPeriod(note);
      if (!period) { stdCount = 4; }
      else {
        const ss = pt(x.shift_start || "");
        const se = pt(x.shift_end || "");
        if (ss == null || se == null) { stdCount = 4; }
        // 4a: 完全吻合班次起止 → 0
        else if (period.startMin === ss && period.endMin === se) { stdCount = 0; }
        else {
          // 4b: 根据排班正确+休息时间计算
          const correct = x.is_schedule_correct;
          const rs = pt(x.rest_start || "");
          const re = pt(x.rest_end || "");
          if (correct === "正确") {
            if (rs != null && re != null) {
              // 完全在休息区间内 → 0
              if (period.startMin >= rs && period.endMin <= re) { stdCount = 0; }
              // 与休息交叉 → 3
              else if (period.startMin < re && period.endMin > rs &&
                       !(period.startMin <= rs && period.endMin >= re) &&
                       period.startMin >= ss && period.endMin <= se) { stdCount = 3; }
              // 完全包含休息区间 → 2
              else if (period.startMin <= rs && period.endMin >= re &&
                       period.startMin >= ss && period.endMin <= se) { stdCount = 2; }
              else { stdCount = 4; }
            } else { stdCount = 4; }
          } else {
            // 排班正确=不正确
            const periodHours = (period.endMin - period.startMin) / 60;
            if (periodHours >= 7) { stdCount = 0; }
            else if (periodHours >= 4) { stdCount = 2; }
            else { stdCount = 4; }
          }
        }
      }
    }

    const miss = Math.max(0, stdCount - p);

    return {
      ...x,
      is_scheduled: isScheduled ? "是" : "否",
      standard_punch_count: stdCount,
      miss_count: miss,
    };
  });
}

/**
 * 规则6: 备注请假/出差/居家办公覆盖（SKILL apply_note_override 精确匹配）
 *
 * 当备注"请假/出差/居家办公"等类型并匹配到时间段时：
 * - 时长 ≥ 8H → 标准打卡=0, 排班正确=正确
 * - 时长 ≥ 4H → 标准打卡=2, 打卡<2→排班正确=不正确
 * - 含[已废弃]/[已撤回]且无[审批中]/[已完成]时不覆盖
 */
const NOTE_ACTIVE = /\[(?:审批中|已完成)\]/;
const NOTE_EXPIRED = /\[(?:已废弃|已撤回)\]/;
const NOTE_LEAVE_TYPES = /居家办公|公出|出差|病假|年假|无薪病假|事假|调休|婚假|产假|陪产假|丧假|工伤假/;

function isNoteExpired(n: string): boolean {
  if (NOTE_ACTIVE.test(n)) return false;
  return NOTE_EXPIRED.test(n);
}

export function rule6NoteOverride(rows: ProcessingRow[]): ProcessingRow[] {
  let c8 = 0, c4 = 0;
  return rows.map((x) => {
    const n = (x.note || "").trim();
    if (!n) return x;
    if (!NOTE_LEAVE_TYPES.test(n)) return x;
    // 过期备注不覆盖
    if (isNoteExpired(n)) return x;

    const period = parseWorkPeriod(n);
    if (!period) return x;

    const durationHours = (period.endMin - period.startMin) / 60;
    let cr = x.is_schedule_correct;
    let st = x.standard_punch_count;

    if (durationHours >= 8) {
      st = 0;
      cr = "正确";
      c8++;
    } else if (durationHours >= 4) {
      st = 2;
      const pc = Number(x.punch_count) || 0;
      if (pc < 2) cr = "不正确";
      c4++;
    }

    return {
      ...x,
      standard_punch_count: st,
      is_schedule_correct: cr,
      miss_count: Math.max(0, st - (Number(x.punch_count) || 0)),
    };
  });
}
