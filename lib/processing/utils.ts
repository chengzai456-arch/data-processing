/**
 * 字符串/数值归一化：把 Excel 里常见的脏值统一处理成空字符串
 * 兼容 "nan"、"NaT"、"None"、"NaN"、"nat"、"/" 等缺失标记
 */
export function toStr(v: unknown): string {
  if (v == null) return "";
  const s = String(v).trim();
  if (["nan", "NaT", "None", "NaN", "nat", "", "/"].includes(s)) return "";
  return s;
}

/**
 * 宽松数字解析
 *
 * 相比直接 Number(s)，这里兼容了 Excel / 手工表格里常见的格式：
 * - 千位分隔符: "1,234.5"
 * - 百分比: "85.5%"
 * - 货币符号: "¥1,200" / "$45.9"
 * - 中文单位: "40小时" / "1.5H" / "120分"
 * - 括号负数: "(120)" → -120
 * - 纯时间 "HH:MM[:SS]"（如把时间错填到数字列）→ 秒数 / 分钟数
 * - 缺失标记 "-"、"--"、"#"、空格 → 0
 */
export function toNum(v: unknown): number {
  const s = toStr(v);
  if (!s) return 0;
  if (["-", "--", "#", "#N/A", "#REF!", "#VALUE!"].includes(s)) return 0;

  let t = s;

  // 百分比
  if (t.endsWith("%")) {
    const n = Number(t.slice(0, -1).trim());
    return Number.isFinite(n) ? n : 0;
  }

  // 括号负数
  const paren = t.match(/^\((.+)\)$/);
  if (paren) {
    const n = toNum(paren[1]);
    return -Math.abs(n);
  }

  // 纯时间 "HH:MM[:SS]" → 转成秒数（用于时间错填到数字列的情况）
  const timeM = t.match(/^\d{1,2}:\d{2}(:\d{2})?$/);
  if (timeM) {
    const parts = timeM[0].split(":").map(Number);
    if (parts.length === 2) return parts[0] * 60 + parts[1]; // 分钟
    return parts[0] * 3600 + parts[1] * 60 + parts[2]; // 秒
  }

  // 去掉货币符号和千位分隔符
  t = t.replace(/[¥$€£￥]/g, "").replace(/,/g, "");

  // 中文单位 / 英文字母单位
  const cnUnit = t.match(/^(-?\d+(?:\.\d+)?)\s*(小时|时|分钟|分|天|h|hr|hrs|小时数|人|元)?$/i);
  if (cnUnit) {
    const n = Number(cnUnit[1]);
    if (!Number.isFinite(n)) return 0;
    const unit = (cnUnit[2] || "").toLowerCase();
    if (unit === "分钟" || unit === "分") return n / 60; // 统一成小时
    if (unit === "h" || unit === "hr" || unit === "hrs" || unit === "小时" || unit === "时") return n;
    return n;
  }

  const n = Number(t);
  return Number.isFinite(n) ? n : 0;
}

/**
 * 日期解析
 * - 支持 "2026/7/22"、"2026-07-22"、"2026.7.22"、"2026年7月22日"、"2026-07-22 08:00:00" 等文本
 * - 支持 Excel 序列日期数字（如 45394）
 * - 解析失败返回 null
 */
export function parseDate(v: unknown): string | null {
  // Excel 序列日期数字
  if (typeof v === "number" && Number.isFinite(v) && v > 1) {
    const ms = (v - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  const s = toStr(v);
  if (!s) return null;

  // 数字字符串（"45394"）
  const serial = Number(s);
  if (/^\d{5,6}$/.test(s) && Number.isFinite(serial)) {
    const ms = (serial - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }

  // 常规文本格式
  const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;

  // 中文日期 "2026年7月22日"
  const cn = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (cn) return `${cn[1]}-${cn[2].padStart(2, "0")}-${cn[3].padStart(2, "0")}`;

  return null;
}
