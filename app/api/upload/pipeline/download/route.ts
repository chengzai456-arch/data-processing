import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

export const dynamic = 'force-dynamic';

/** 分批拉取全量数据（避免单次 limit 5000 截断） */
async function fetchAll(
  supabase: any,
  query: any,
): Promise<Record<string, unknown>[]> {
  const all: Record<string, unknown>[] = [];
  const PAGE = 2000;
  let from = 0;
  for (;;) {
    const { data, error } = await query.range(from, from + PAGE - 1);
    if (error) throw error;
    if (!data || data.length === 0) break;
    all.push(...data);
    from += PAGE;
    if (data.length < PAGE) break;
  }
  return all;
}

export async function GET(req: NextRequest) {
  const date = req.nextUrl.searchParams.get('date');
  const worker = req.nextUrl.searchParams.get('worker') || 'GUS';

  const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

  // 数据库查询模式（有配置时）
  if (supabaseUrl && supabaseKey) {
    try {
      const { createClient } = await import('@supabase/supabase-js');
      const supabase = createClient(supabaseUrl, supabaseKey, {
        auth: { autoRefreshToken: false, persistSession: false },
      });

      let query = supabase.from('attendance_data').select('*').eq('worker_type', worker);
      if (date) query = query.eq('date', date);
      query = query.order('employee_code');

      const data = await fetchAll(supabase, query);
      if (data && data.length > 0) {
        const headers = Object.keys(data[0]);
        const rows = data.map((r: any) => headers.map((h) => r[h]));
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '处理结果');
        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        return new NextResponse(buf, {
          status: 200,
          headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename="attendance_processed_${date || 'all'}.xlsx"`,
          },
        });
      }
    } catch {}
  }

  // 无数据库时，返回一个包含处理统计的简单 Excel
  const stats = [
    ['考勤数据处理报告'],
    [''],
    ['数据日期', date || '未知'],
    ['工种', worker],
    [''],
    ['说明', '数据库未配置，无法下载完整数据。请设置 SUPABASE_SERVICE_ROLE_KEY 环境变量。'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(stats);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '报告');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  return new NextResponse(buf, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="report_${date || 'all'}.xlsx"`,
    },
  });
}
