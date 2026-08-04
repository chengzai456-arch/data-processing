/**
 * 本地处理脚本：用真实数据运行管道并输出完整 Excel
 * npx tsx scripts/run_local.ts
 */
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { RawRow, ProcessingRow, UploadFiles } from '../lib/processing/types';
import { readSheetRows } from '../lib/processing/excel';
import { mapRow } from '../lib/processing/pipeline';
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
    for (const r of readSheetRows(buf)) {
      const n = String(r['班次名称']??'').trim(); if (!n) continue;
      d[n] = {rest_start:String(r['休息开始时间']??'').trim(), rest_end:String(r['休息结束时间']??'').trim(), work_start:String(r['上班时间']||r['班次上班时间']||'').trim()};
    }
  } catch {}
  return d;
}

async function main() {
  console.log('加载数据文件...\n');
  const mainBuf = fs.readFileSync(path.join(DATA_DIR, '722GUS+每日打卡工时推送模版 (51).xlsx'));

  // 1. 解析
  const raw = readSheetRows(mainBuf) as RawRow[];
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
