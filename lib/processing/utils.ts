export function toStr(v: unknown): string {
  if (v == null) return '';
  const s = String(v).trim();
  if (['nan', 'NaT', 'None', 'NaN', 'nat', '', '/'].includes(s)) return '';
  return s;
}

export function toNum(v: unknown): number {
  const s = toStr(v);
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

export function parseDate(v: unknown): string | null {
  const s = toStr(v);
  if (!s) return null;
  if (typeof v === 'number') {
    const ms = (v - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]|$)/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  return null;
}
