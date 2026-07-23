import * as XLSX from 'xlsx';
import { ProcessingRow } from './types';

const HKW = ['每日总工时','总工时','工时','时长','小时'];
const SUB = 40;

function tn(v:unknown):number{const s=String(v??'').trim();if(!s||['nan','NaT','None','NaN','nat','','/'].includes(s))return 0;const n=Number(s);return Number.isFinite(n)?n:0;}

function psr(buf:Buffer,sub:boolean):Record<string,number>{
  const rows=XLSX.utils.sheet_to_json<Record<string,unknown>>(XLSX.read(buf,{type:'buffer'}).Sheets[XLSX.read(buf,{type:'buffer'}).SheetNames[0]],{defval:''});
  const h=rows.length>0?Object.keys(rows[0]):[];const hc=h.find(h=>HKW.some(k=>h.includes(k)));const ec=h.find(h=>h.includes('工号'));
  if(!hc||!ec)return{};
  const r:Record<string,number>={};
  for(const row of rows){const e=String(row[ec]??'').trim();if(!e)continue;const h=tn(row[hc]);if(sub)r[e]=Math.max(0,(r[e]||0)+h-SUB);else r[e]=(r[e]||0)+h;}
  return r;
}

export function step7MatchHours(rows:ProcessingRow[],sThis?:Buffer,sLast?:Buffer,sBi?:Buffer):ProcessingRow[]{
  const w=sThis?psr(sThis,true):{},l=sLast?psr(sLast,true):{},b=sBi?psr(sBi,false):{};
  console.log(`[step7] 签字报表: 本周${Object.keys(w).length}/上周${Object.keys(l).length}/双周${Object.keys(b).length}`);
  return rows.map(r=>({...r,sign_report_hours:w[r.employee_code]??r.sign_report_hours??0,week_overtime_hours:w[r.employee_code]??r.week_overtime_hours??0,last_week_overtime_hours:l[r.employee_code]??r.last_week_overtime_hours??0,sign_hours:b[r.employee_code]??r.sign_hours??0}));
}
