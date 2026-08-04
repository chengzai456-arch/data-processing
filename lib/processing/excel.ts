import * as XLSX from "xlsx";

/**
 * 共享 Excel 读取工具
 *
 * 统一封装 workbook 的解析与缓存：
 * - 同一个 Buffer 只解析一次，避免各 step / 规则里重复 XLSX.read 的开销
 * - 优先取第一个非空 sheet，防止偶发空 sheet 导致数据全丢
 */
export function readWorkbook(buf: Buffer): XLSX.WorkBook {
  return XLSX.read(buf, { type: "buffer" });
}

/** 读取第一个有内容的 sheet 的行数组（defval 统一为 ""，读取容错） */
export function readSheetRows(buf: Buffer): Record<string, unknown>[] {
  const wb = readWorkbook(buf);
  const names = (wb.SheetNames || []).filter((n) => wb.Sheets[n]);
  if (names.length === 0) return [];
  // 优先取第一个非空 sheet：默认 sheet 可能为空模板
  for (const name of names) {
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      wb.Sheets[name],
      { defval: "" },
    );
    if (rows.length > 0) return rows;
  }
  // 全部为空时退回第一个 sheet 的结果
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(
    wb.Sheets[names[0]],
    { defval: "" },
  );
}
