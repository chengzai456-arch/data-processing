import { ProcessingRow } from "./types";

/**
 * 数据校验报告
 * 在所有步骤执行完后调用，输出数据的合理性检查
 */
export interface ValidationReport {
  /** 总体状态 */
  passed: boolean;
  /** 检查项列表 */
  checks: ValidationCheck[];
  /** 汇总 */
  summary: string;
}

export interface ValidationCheck {
  name: string;
  passed: boolean;
  message: string;
}

export function generateValidationReport(
  finalRows: ProcessingRow[],
  resignedCount: number,
  gusRemovedCount: number,
): ValidationReport {
  const checks: ValidationCheck[] = [];
  let allPassed = true;

  if (finalRows.length === 0) {
    return {
      passed: false,
      checks: [{ name: "数据为空", passed: false, message: "处理后数据为0行" }],
      summary: "❌ 处理后无数据，请检查输入文件",
    };
  }

  // 1. 总人数
  const uniqueEmps = new Set(finalRows.map((r) => r.employee_code));
  checks.push({
    name: "总人数",
    passed: true,
    message: `${uniqueEmps.size} 人`,
  });

  // 2. 排班率
  const schedCount = finalRows.filter((r) => r.is_scheduled === "是").length;
  const schedRate = ((schedCount / finalRows.length) * 100).toFixed(1);
  checks.push({
    name: "排班率",
    passed: schedCount > 0,
    message: `${schedRate}%（${schedCount}/${finalRows.length}）`,
  });

  // 3. 打卡率
  const totalStandard = finalRows.reduce(
    (s, r) => s + r.standard_punch_count,
    0,
  );
  const totalPunch = finalRows.reduce((s, r) => s + r.punch_count, 0);
  const punchRate =
    totalStandard > 0 ? ((totalPunch / totalStandard) * 100).toFixed(1) : "N/A";
  checks.push({
    name: "打卡率",
    passed: totalStandard > 0,
    message: `${punchRate}%（${totalPunch}/${totalStandard}）`,
  });

  // 4. 缺卡异常（缺卡数 > 标准打卡数）
  const missAnomalies = finalRows.filter(
    (r) => r.miss_count > r.standard_punch_count,
  ).length;
  checks.push({
    name: "缺卡异常",
    passed: missAnomalies === 0,
    message:
      missAnomalies > 0
        ? `⚠️ ${missAnomalies} 行缺卡数超过标准打卡数`
        : "正常",
  });

  // 5. 日超8H
  const over8h = finalRows.filter((r) => r.is_over8h === "是").length;
  checks.push({
    name: "日超8H",
    passed: true,
    message: `${over8h} 人（${((over8h / finalRows.length) * 100).toFixed(1)}%）`,
  });

  // 6. 补签覆盖
  const makeupCount = finalRows.reduce((s, r) => s + r.makeup_count, 0);
  checks.push({
    name: "补签数",
    passed: true,
    message: `共 ${makeupCount} 次补签记录`,
  });

  // 7. 无班次人员
  const noShift = finalRows.filter((r) => !(r.shift_name || "").trim()).length;
  checks.push({
    name: "无班次人员",
    passed: noShift < finalRows.length * 0.5,
    message:
      noShift > 0
        ? `⚠️ ${noShift} 人（${((noShift / finalRows.length) * 100).toFixed(1)}%）`
        : "无",
  });

  // 8. 离职/GUS剔除统计
  checks.push({
    name: "数据剔除",
    passed: true,
    message: `离职 ${resignedCount} 人 / GUS白名单 ${gusRemovedCount} 人`,
  });

  allPassed = checks.every((c) => c.passed);

  return {
    passed: allPassed,
    checks,
    summary: allPassed
      ? `✅ 数据校验通过 · ${finalRows.length} 行 · ${uniqueEmps.size} 人`
      : `⚠️ 数据存在异常，请查看详情`,
  };
}
