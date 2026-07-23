/**
 * 本地处理脚本：用真实数据运行管道并输出完整 Excel
 * npx tsx scripts/run_local.ts
 */
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { toStr, toNum, parseDate } from '../lib/processing/utils';
import { RawRow, ProcessingRow, UploadFiles } from '../lib/processing/types';
import { step1RemoveResigned } from '../lib/processing/step1_remove_resigned';
import { step2RemoveNotJoined } from '../lib/processing/step2_remove_not_joined';
import { step3Gl00Handle } from '../lib/processing/step3_gl00_handle';
import { step4RemoveEuHr } from '../lib/processing/step4_remove_eu_hr';
import { step5RemoveGus } from '../lib/processing/step5_remove_gus';
import { step6MatchSupplement } from '../lib/processing/step6_match_supplement';
import { step7MatchHours } from '../lib/processing/step7_match_hours';
import { rule0RestTime, rule1HubMark, rule2ShiftCorrect, rule3DailyHours, rule4Over8h, rule5SchedulePunch, rule6NoteOverride } from '../lib/processing/rules';

const DATA_DIR = 'D:/Documents/Downloads/722日晾晒数据';
const OUT_DIR = path.resolve(__dirname, '../../output');

function readShiftDict(buf: Buffer): Record<string, {rest_start:string; rest_end:string; work_start:string}> {
  const d: any = {};
  try {
    for (const r of XLSX.utils.sheet_to_json<any>(XLSX.read(buf,{type:'buffer'}).Sheets[XLSX.read(buf,{type:'buffer'}).SheetNames[0]],{defval:''})) {
      const n = String(r['班次名称']??'').trim(); if (!n) continue;
      d[n] = {rest_start:String(r['休息开始']??'').trim(), rest_end:String(r['休息结束']??'').trim(), work_start:String(r['上班时间']||r['班次上班时间']||'').trim()};
    }
  } catch {}
  return d;
}

function mapRow(raw: RawRow, wt: string): ProcessingRow|null {
  const date = parseDate(raw['考勤日期']); if (!date) return null;
  const code = toStr(raw['工号']); if (!code) return null;
  const l = wt === 'GUS_LABOR';
  return {
    employee_code: code, date, employee_name: toStr(raw['姓名']||raw['员工名称']||''),
    department_level3: toStr(raw['三级部门']||raw['区域']||''), department_level4: toStr(raw['四级部门']||raw['仓库']||''),
    department_level5: toStr(raw['五级部门']||raw['组']||''), department: toStr(raw['部门']||raw['三级部门']||''),
    shift_name: toStr(raw['班次名称']||raw['班次']||''), shift_start: toStr(raw['班次上班时间']||''), shift_end: toStr(raw['班次下班时间']||''),
    first_punch: toStr(raw['首打卡时间']||''), last_punch: toStr(raw['末打卡时间']||''),
    punch_count: toNum(raw['班次内打卡次数']||raw['辅助列']), is_schedule_correct: toStr(raw['是否排班正确']),
    is_scheduled: toStr(raw['是否排班']), standard_punch_count: toNum(raw['标准打卡数']), miss_count: toNum(raw['缺卡数']),
    makeup_count: toNum(raw['补签数']), is_over8h: toStr(raw['是否日超8H']),
    daily_total_hours: toNum(raw['每日总工时(公式：末打卡-首打卡-班次午休时间+居家办公时长)合计']||raw['时长总计']||raw['每日总工时']),
    overtime_hours: toNum(raw['日超8H']||raw['加班工时']), week_overtime_hours: toNum(raw['本周加班工时']||raw['本周累计加班工时']),
    last_week_overtime_hours: toNum(raw['上周加班工时']||raw['上周累计加班工时']), sign_hours: toNum(raw['双周加班工时']),
    sign_report_hours: 0, is_hub: toStr(raw['HUB']||'').length>0, hub_status: toStr(raw['HUB']||''),
    note: toStr(raw['备注（GF）']||''), pending_home_office_hours: toNum(raw['居家办公合计（审批中）']),
    rest_start: toStr(raw['休息开始时间']||''), rest_end: toStr(raw['休息结束时间']||''), rest_time: '',
    supplier_name: l?toStr(raw['供应商名称']||''):'', worker_type_label: l?toStr(raw['工种']||''):'',
    first_last_miss: l?toNum(raw['首末缺卡数']):0, mid_punch_1: l?toStr(raw['中间打卡时间1']||''):'',
    mid_punch_2: l?toStr(raw['中间打卡时间2']||''):'', lunch_miss: l?toNum(raw['午休缺卡数']):0,
    sign_start: l?toStr(raw['首打卡补签时间']||''):'', sign_rest_start: l?toStr(raw['休息开始补签时间']||''):'',
    sign_rest_end: l?toStr(raw['休息结束补签时间']||''):'', sign_end: l?toStr(raw['末打卡补签时间']||''):'',
    helper_col: l?toNum(raw['辅助列']):0,
  };
}

async function main() {
  console.log('加载数据文件...\n');
  const mainBuf = fs.readFileSync(path.join(DATA_DIR, '722GUS+每日打卡工时推送模版 (51).xlsx'));

  // 1. 解析
  const raw = XLSX.utils.sheet_to_json<RawRow>(XLSX.read(mainBuf,{type:'buffer'}).Sheets[XLSX.read(mainBuf,{type:'buffer'}).SheetNames[0]],{defval:''});
  const rows = raw.map(r => mapRow(r, 'GUS')).filter(Boolean) as ProcessingRow[];
  console.log(`原始数据: ${raw.length} 行, 有效: ${rows.length} 行\n`);

  // 2. 各步骤
  const s1 = rows.length;
  const s2 = step1RemoveResigned(rows, fs.readFileSync(path.join(DATA_DIR, '722离职流程 (62).xlsx')));
  const s3 = step2RemoveNotJoined(s2, fs.readFileSync(path.join(DATA_DIR, '722花名册 (69).xlsx')), '2026-07-22');
  const s4 = step3Gl00Handle(s3);
  const s5 = step4RemoveEuHr(s4);
  const s6 = step5RemoveGus(s5, fs.readFileSync(path.join(DATA_DIR, '722GUS白名单人员 (3).xlsx')));
  const s7 = step6MatchSupplement(s6, fs.readFileSync(path.join(DATA_DIR, '722补签管理 (77).xlsx')), '2026-07-22');

  // 签字报表（可选，文件太大）
  let s8 = s7;
  try {
    s8 = step7MatchHours(s7,
      fs.readFileSync(path.join(DATA_DIR, '720-722【本周加班工时】GUS+美区签字报表 (99).xlsx')),
      fs.readFileSync(path.join(DATA_DIR, '713-719【上周加班工时】GUS+美区签字报表 (100).xlsx')),
      fs.readFileSync(path.join(DATA_DIR, '713-726【双周累计工时】GUS+美区签字报表 (98).xlsx')),
    );
  } catch (e) {
    console.log('签字报表加载失败（文件可能过大），跳过:', (e as Error).message);
  }

  // 3. 指标规则
  let cur = s8;
  const sd = readShiftDict(fs.readFileSync(path.join(DATA_DIR, '班次.xlsx')));
  cur = rule0RestTime(cur, sd);
  cur = rule1HubMark(cur);
  cur = rule2ShiftCorrect(cur, sd);
  cur = rule3DailyHours(cur);
  cur = rule4Over8h(cur);
  cur = rule5SchedulePunch(cur);
  cur = rule6NoteOverride(cur);

  console.log(`\n最终行数: ${cur.length}`);

  // 4. 输出 Excel（中文字段名）
  const outHeaders = [
    '工号', '姓名', '考勤日期', '三级部门', '四级部门', '五级部门',
    '班次名称', '首打卡', '末打卡', '打卡次数',
    '是否排班', '标准打卡', '缺卡', '补签数',
    '每日总工时', '加班工时', '日超8H',
    '本周加班', '上周加班', '双周加班',
    '排班正确', 'HUB', '备注',
    '休息开始', '休息结束',
  ];

  const outData = cur.map(r => [
    r.employee_code, r.employee_name, r.date,
    r.department_level3, r.department_level4, r.department_level5,
    r.shift_name, r.first_punch, r.last_punch, r.punch_count,
    r.is_scheduled, r.standard_punch_count, r.miss_count, r.makeup_count,
    r.daily_total_hours, r.overtime_hours, r.is_over8h,
    r.week_overtime_hours, r.last_week_overtime_hours, r.sign_hours,
    r.is_schedule_correct, r.is_hub ? '是' : '否', r.note,
    r.rest_start, r.rest_end,
  ]);

  const sheet1 = XLSX.utils.aoa_to_sheet([outHeaders, ...outData]);
  // 统计表
  const stats = [
    ['数据处理报告', ''],
    ['数据日期', '2026-07-22'],
    ['', ''],
    ['阶段', '行数'],
    ['原始数据', s1],
    ['剔除离职', s1 - s2.length],
    ['GL00处理', s2.length - s4.length],
    ['GUS白名单', s4.length - s6.length],
    ['补签匹配', s6.length - s7.length],
    ['最终行数', cur.length],
  ];
  const sheet2 = XLSX.utils.aoa_to_sheet(stats);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet1, '处理数据');
  XLSX.utils.book_append_sheet(wb, sheet2, '处理统计');

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const outPath = path.join(OUT_DIR, '722_processed_data.xlsx');
  fs.writeFileSync(outPath, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  console.log(`\n✅ 输出文件: ${outPath}`);
}

main().catch(e => console.error('❌ 错误:', e));
