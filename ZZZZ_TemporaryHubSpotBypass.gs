/**
 * TEMPORARY HubSpot bypass for the monthly quality report.
 *
 * Remove this file after the HubSpot service token is available again.
 *
 * ZZZ_DPPMChartOverride.gs normally performs a live HubSpot API sync and then
 * rebuilds HubSpot Chart Data and DPPM Inputs from the synchronized raw ticket
 * tab. While the service token is unavailable, that path fails with 401.
 *
 * This temporary override intentionally does NOT call HubSpot and does NOT
 * rebuild from the stale "HubSpot Support Tickets" tab. It uses the already
 * validated values currently stored in DPPM Inputs and HubSpot Chart Data,
 * then updates the linked charts and issue-summary tables in Slides.
 *
 * This file is named ZZZZ_* so it loads after ZZZ_DPPMChartOverride.gs and this
 * implementation wins for updateMonthlyQualityPackageDPPM_().
 */

function updateMonthlyQualityPackageDPPM_(packageResult) {
  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());

  const hubSpotSyncResult = {
    status: 'SKIPPED',
    reason: 'Temporary bypass until HubSpot service token is available'
  };
  Logger.log(JSON.stringify(hubSpotSyncResult));

  const validatedIssueChartResult =
    readMonthlyQualityExistingIssueChartData_(spreadsheet);
  Logger.log(JSON.stringify(validatedIssueChartResult));

  SpreadsheetApp.flush();

  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const updatedSections = [];

  MONTHLY_QUALITY_VISIBLE_DPPM_SECTIONS_.forEach(function(section) {
    const chart = findMonthlyQualityVisibleDPPMChart_(spreadsheet, section);
    const slide = findMonthlyQualitySlide_(presentation, section.slideLabel);
    replaceMonthlyQualityVisibleDPPMChart_(slide, chart);
    updatedSections.push(section.slideLabel);
  });

  normalizeMonthlyQualityServiceChartSlides_(spreadsheet, presentation);
  updateMonthlyQualityIssueSummaryTables_(
    presentation,
    validatedIssueChartResult
  );

  const linkedChartsRefreshed = refreshMonthlyQualityLinkedSheetsCharts_(
    presentation
  );

  presentation.saveAndClose();
  return {
    updatedSections: updatedSections,
    hubSpotSync: hubSpotSyncResult,
    validatedIssueChartData: validatedIssueChartResult,
    linkedChartsRefreshed: linkedChartsRefreshed
  };
}

/**
 * Read the already-validated issue counts from HubSpot Chart Data without
 * contacting HubSpot or rebuilding DPPM Inputs.
 */
function readMonthlyQualityExistingIssueChartData_(spreadsheet) {
  const configSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_CONFIG_SHEET_
  );
  if (!configSheet) {
    throw new Error('Issue chart source is missing "DPPM Config".');
  }

  const chartDataSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_CHART_DATA_SHEET_
  );
  if (!chartDataSheet) {
    throw new Error(
      'Issue chart source is missing "' +
      VALIDATED_ISSUE_CHART_DATA_SHEET_ + '".'
    );
  }

  const reportMonth = validatedIssueMonthStart_(
    configSheet.getRange('B7').getValue()
  );
  if (!reportMonth) {
    throw new Error('DPPM Config!B7 does not contain a valid report month.');
  }

  const previousMonth = new Date(
    reportMonth.getFullYear(),
    reportMonth.getMonth() - 1,
    1
  );

  const summaries = {};

  VALIDATED_ISSUE_BREAKDOWN_BLOCKS_.forEach(function(block) {
    const rows = chartDataSheet
      .getRange(2, block.startColumn, 12, 4)
      .getValues();

    summaries[block.product] = {
      current: readMonthlyQualityExistingIssueCounts_(
        rows,
        reportMonth,
        block.product
      ),
      previous: readMonthlyQualityExistingIssueCounts_(
        rows,
        previousMonth,
        block.product
      )
    };
  });

  summaries['All Lines'] = {
    current: sumMonthlyQualityExistingIssueCounts_([
      summaries.MSC.current,
      summaries.ARU.current,
      summaries.CSC.current
    ]),
    previous: sumMonthlyQualityExistingIssueCounts_([
      summaries.MSC.previous,
      summaries.ARU.previous,
      summaries.CSC.previous
    ])
  };

  return {
    status: 'READY',
    source: VALIDATED_ISSUE_CHART_DATA_SHEET_ + ' (existing validated values)',
    method: 'Temporary offline mode; HubSpot API sync skipped',
    reportMonth: validatedIssueMonthKey_(reportMonth),
    currentMonthLabel: Utilities.formatDate(reportMonth, 'UTC', 'MMM'),
    previousMonthLabel: Utilities.formatDate(previousMonth, 'UTC', 'MMM'),
    dppmInputsUpdated: 0,
    summaries: summaries
  };
}

function readMonthlyQualityExistingIssueCounts_(rows, targetMonth, product) {
  const targetKey = validatedIssueMonthKey_(targetMonth);

  for (let index = 0; index < rows.length; index++) {
    const month = validatedIssueMonthStart_(rows[index][0]);
    if (!month || validatedIssueMonthKey_(month) !== targetKey) continue;

    const startup = Number(rows[index][1]) || 0;
    const warranty = Number(rows[index][2]) || 0;
    const service = Number(rows[index][3]) || 0;

    return {
      startup: startup,
      warranty: warranty,
      service: service,
      total: startup + warranty + service
    };
  }

  throw new Error(
    'Existing HubSpot Chart Data is missing ' + targetKey +
    ' for ' + product + '.'
  );
}

function sumMonthlyQualityExistingIssueCounts_(records) {
  const total = {
    startup: 0,
    warranty: 0,
    service: 0,
    total: 0
  };

  records.forEach(function(record) {
    total.startup += Number(record.startup) || 0;
    total.warranty += Number(record.warranty) || 0;
    total.service += Number(record.service) || 0;
    total.total += Number(record.total) || 0;
  });

  return total;
}
