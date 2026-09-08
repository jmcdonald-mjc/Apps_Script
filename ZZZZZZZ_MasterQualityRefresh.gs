/**
 * One-button master Quality refresh.
 *
 * calculateFPYSummary_FINAL() has historically been the function run by the
 * user to refresh the complete Quality reporting workflow. The source
 * implementation in Code.gs rebuilds the SafetyCulture/FPY source workbook,
 * but it currently returns after the FPY/Pareto tabs are written and therefore
 * never reaches the monthly report package.
 *
 * This file intentionally sorts last. Capture the existing data-only function,
 * then replace the public entry point with a wrapper that:
 *   1. rebuilds the FPY/SafetyCulture source data;
 *   2. flushes those writes;
 *   3. updates the prior completed month's monthly workbook and Slides report;
 *   4. propagates any report error so the run cannot appear successful when the
 *      presentation was not updated.
 */

var calculateFPYSummary_FINAL_DATA_ONLY_ = calculateFPYSummary_FINAL;

calculateFPYSummary_FINAL = function() {
  const startedAt = new Date();

  Logger.log('Master Quality Refresh: rebuilding SafetyCulture / FPY source data.');
  calculateFPYSummary_FINAL_DATA_ONLY_();
  SpreadsheetApp.flush();

  Logger.log(
    'Master Quality Refresh: FPY source data complete. ' +
    'Updating prior completed month workbook and Slides report.'
  );

  // With no asOfDate argument, getMonthlyQualityPackageContext_() intentionally
  // targets the prior completed month. Example: a September run updates August.
  const reportResult = updateMonthlyQualityReport();

  Logger.log(JSON.stringify({
    status: 'COMPLETE',
    startedAt: startedAt.toISOString(),
    completedAt: new Date().toISOString(),
    monthlyReport: reportResult
  }));

  return reportResult;
};
