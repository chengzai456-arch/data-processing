import { createClient } from '@supabase/supabase-js';

interface UploadRecord {
  id: string;
  worker_type?: string;
  status?: string;
  data_date?: string;
  rows_written?: number;
  created_at?: string;
  completed_at?: string;
}

// 服务端直连 Supabase（仅服务端执行，service key 不会泄露到浏览器）
async function fetchUploads(): Promise<{ records: UploadRecord[] | null; error: string | null }> {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!url || !key) return { records: null, error: null };

  try {
    const supabase = createClient(url, key, {
      auth: { autoRefreshToken: false, persistSession: false },
    });
    const { data, error } = await supabase
      .from('upload_records')
      .select('id, worker_type, status, data_date, rows_written, created_at, completed_at')
      .order('created_at', { ascending: false })
      .limit(200);
    if (error) return { records: null, error: error.message };
    return { records: (data as UploadRecord[]) || [], error: null };
  } catch (e) {
    return { records: null, error: (e as Error).message };
  }
}

function fmtDate(iso?: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso;
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export default async function HistoryPage() {
  const { records, error } = await fetchUploads();

  return (
    <main className="max-w-6xl mx-auto px-4 py-8">
      <h1 className="text-2xl font-bold mb-1">处理历史</h1>
      <p className="text-gray-500 text-sm mb-6">已上传数据的处理记录，可重新下载当天结果</p>

      {error ? (
        <div className="p-4 bg-red-50 border border-red-200 rounded-lg">
          <p className="text-red-700 font-medium">读取历史记录失败</p>
          <p className="text-sm text-red-600 mt-1">{error}</p>
        </div>
      ) : records === null ? (
        <div className="p-8 text-center border border-dashed rounded-xl">
          <p className="text-gray-500">数据库未配置，无法查询历史记录。</p>
          <p className="text-xs text-gray-400 mt-1">请在 Vercel 环境变量中配置 NEXT_PUBLIC_SUPABASE_URL 和 SUPABASE_SERVICE_ROLE_KEY</p>
        </div>
      ) : records.length === 0 ? (
        <div className="p-8 text-center border border-dashed rounded-xl">
          <p className="text-gray-500">暂无历史记录</p>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 overflow-hidden">
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-200 text-left text-gray-500">
                <th className="px-4 py-3 font-medium">数据日期</th>
                <th className="px-4 py-3 font-medium">工种</th>
                <th className="px-4 py-3 font-medium">行数</th>
                <th className="px-4 py-3 font-medium">状态</th>
                <th className="px-4 py-3 font-medium">处理时间</th>
                <th className="px-4 py-3 font-medium text-right">操作</th>
              </tr>
            </thead>
            <tbody>
              {records.map((r) => (
                <tr key={r.id} className="border-b border-gray-100 hover:bg-gray-50">
                  <td className="px-4 py-3">{r.data_date || '—'}</td>
                  <td className="px-4 py-3">{r.worker_type || 'GUS'}</td>
                  <td className="px-4 py-3">{r.rows_written ?? '—'}</td>
                  <td className="px-4 py-3">
                    <span
                      className={
                        r.status === 'completed'
                          ? 'inline-flex px-2 py-0.5 rounded-full text-xs bg-green-50 text-green-700'
                          : 'inline-flex px-2 py-0.5 rounded-full text-xs bg-gray-100 text-gray-600'
                      }
                    >
                      {r.status || '—'}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-gray-500">{fmtDate(r.completed_at || r.created_at)}</td>
                  <td className="px-4 py-3 text-right">
                    <a
                      href={`/api/upload/pipeline/download?date=${encodeURIComponent(r.data_date || '')}&worker=${encodeURIComponent(r.worker_type || 'GUS')}`}
                      className="inline-flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-medium text-white bg-blue-600 hover:bg-blue-700 transition-colors"
                    >
                      下载当天数据
                    </a>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </main>
  );
}
