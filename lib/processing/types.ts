import * as XLSX from "xlsx";

export interface ProcessingRow {
  employee_code: string;
  date: string;
  employee_name: string;
  department_level3: string;
  department_level4: string;
  department_level5: string;
  department: string;
  shift_name: string;
  shift_start: string;
  shift_end: string;
  first_punch: string;
  last_punch: string;
  punch_count: number;
  is_schedule_correct: string;
  is_scheduled: string;
  standard_punch_count: number;
  miss_count: number;
  makeup_count: number;
  is_over8h: string;
  daily_total_hours: number;
  overtime_hours: number;
  week_overtime_hours: number;
  last_week_overtime_hours: number;
  sign_hours: number;
  sign_report_hours: number;
  is_hub: boolean;
  hub_status: string;
  note: string;
  pending_home_office_hours: number;
  rest_start: string;
  rest_end: string;
  rest_time: string;
  supplier_name: string;
  worker_type_label: string;
  first_last_miss: number;
  mid_punch_1: string;
  mid_punch_2: string;
  lunch_miss: number;
  sign_start: string;
  sign_rest_start: string;
  sign_rest_end: string;
  sign_end: string;
  helper_col: number;
}

export interface ShiftDict {
  [name: string]: { rest_start: string; rest_end: string; work_start: string };
}

export interface UploadFiles {
  file: Buffer;
  roster?: Buffer;
  leave?: Buffer;
  makeup?: Buffer;
  sign_this?: Buffer;
  sign_last?: Buffer;
  sign_biweek?: Buffer;
  gus_whitelist?: Buffer;
  shift_dict?: Buffer;
}

export interface PipelineResult {
  total: number;
  date: string;
  worker_type: string;
  stats: {
    original: number;
    after_step1_remove_resigned: number;
    after_step2_remove_not_joined: number;
    after_step3_gl00: number;
    after_step4_eu_hr: number;
    after_step5_gus: number;
    after_step6_supplement: number;
    after_step7_hours: number;
    final: number;
  };
  /** 数据校验报告 */
  validation?: import("./validation").ValidationReport;
  /** 处理后的行数据（用于浏览器端下载） */
  rows?: Record<string, unknown>[];
}

export interface RawRow {
  工号?: string;
  考勤日期?: string | number;
  部门?: string;
  三级部门?: string;
  四级部门?: string;
  五级部门?: string;
  班次名称?: string;
  班次上班时间?: string;
  班次下班时间?: string;
  首打卡时间?: string;
  末打卡时间?: string;
  班次内打卡次数?: number | string;
  是否排班正确?: string;
  是否排班?: string;
  标准打卡数?: number | string;
  缺卡数?: number | string;
  补签数?: number | string;
  日超8H?: number | string;
  是否日超8H?: string;
  "每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计"?:
    number | string;
  本周加班工时?: number | string;
  上周累计加班工时?: number | string;
  双周加班工时?: number | string;
  HUB?: string;
  "备注（GF）"?: string;
  "居家办公合计（审批中）"?: number | string;
  休息开始时间?: string;
  休息结束时间?: string;
  姓名?: string;
  供应商名称?: string;
  工种?: string;
  时长总计?: number | string;
  加班工时?: number | string;
  首末缺卡数?: number | string;
  中间打卡时间1?: string;
  中间打卡时间2?: string;
  午休缺卡数?: number | string;
  首打卡补签时间?: string;
  休息开始补签时间?: string;
  休息结束补签时间?: string;
  末打卡补签时间?: string;
  辅助列?: number | string;
  员工名称?: string;
  区域?: string;
  仓库?: string;
  组?: string;
  本周累计加班工时?: number | string;
  上周加班工时?: number | string;
  每日总工时?: number | string;
  班次?: string;
  [key: string]: unknown;
}
