/**
 * TEMPORARY HubSpot bypass for the monthly quality report.
 *
 * Remove this file after the HubSpot service token is available again.
 *
 * While HubSpot API authentication is unavailable, the data-first stage uses
 * the manually audited monthly values. This file owns the report-finalization
 * stage: it reads those validated values, rebinds every linked report chart to
 * the CURRENT monthly workbook, updates the issue-summary tables, refreshes the
 * linked charts, and saves the presentation.
 *
 * Important: do not identify report charts by one exact historical title. The
 * July canonical template and the repaired August workbook use slightly
 * different wording (for example "12-Mth" vs "12-Month" and "All" vs
 * "All Lines"). Instead, source charts are indexed from the visible monthly
 * report sheets and matched using a normalized title.
 */

const MONTHLY_QUALITY_REPORT_CHART_SOURCE_SHEETS_ = Object.freeze([
  'Plant Summary',
  'MSC DPPM',
  'CSC DPPM',
  'ARU DPPM',
  'Mods',
  'Gas Heat',
  'Coatings',
  'Bard Coatings'
]);

/**
 * Manual/offline report finalizer.
 *
 * The data-first stage has already:
 *   - skipped the broken HubSpot API token,
 *   - applied the manually audited month,
 *   - synchronized DPPM Inputs,
 *   - rebuilt the DPPM model.
 *
 * This function therefore performs no HubSpot API calls and no issue-data
 * rebuilding. It only consumes the validated monthly workbook and updates the
 * Slides report.
 */
function updateMonthlyQualityPackageDPPM_(packageResult) {
  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());

  const hubSpotSyncResult = {
    status: 'SKIPPED',
    reason: 'Temporary manual HubSpot mode until Apps Script API authentication is repaired'
  };
  Logger.log(JSON.stringify(hubSpotSyncResult));

  const validatedIssueChartResult =
    readMonthlyQualityExistingIssueChartData_(spreadsheet);
  Logger.log(JSON.stringify(validatedIssueChartResult));

  SpreadsheetApp.flush();

  const presentation = SlidesApp.openById(packageResult.deckFile.getId());

  // Rebind ALL linked report charts to the current monthly workbook. This fixes
  // both the July/August title wording mismatch and stale/deleted chart IDs in
  // an already-created monthly deck.
  const chartRebindResult = rebindMonthlyQualityReportCharts_(
    spreadsheet,
    presentation
  );

  updateMonthlyQualityIssueSummaryTables_(
    presentation,
    validatedIssueChartResult
  );

  const linkedChartsRefreshed = refreshMonthlyQualityLinkedSheetsCharts_(
    presentation
  );

  verifyMonthlyQualityReportChartBindings_(
    presentation,
    spreadsheet.getId()
  );

  presentation.saveAndClose();

  return {
    updatedSections: ['All Lines', 'MSC', 'CSC', 'ARU'],
    hubSpotSync: hubSpotSyncResult,
    validatedIssueChartData: validatedIssueChartResult,
    chartRebind: chartRebindResult,
    linkedChartsRefreshed: linkedChartsRefreshed
  };
}

/**
 * Build a source-chart index from only the visible report sheets. Hidden
 * dashboard charts are intentionally excluded so a similarly named legacy
 * chart can never win over the visible July-style chart.
 */
function buildMonthlyQualityReportChartIndex_(spreadsheet) {
  const exact = {};
  const normalized = {};
  const sourceTitles = [];

  MONTHLY_QUALITY_REPORT_CHART_SOURCE_SHEETS_.forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('Monthly report chart source sheet not found: ' + sheetName);
    }

    sheet.getCharts().forEach(function(chart) {
      const title = String(chart.getOptions().get('title') || '').trim();
      if (!title) return;

      const record = {
        chart: chart,
        title: title,
        sheetName: sheetName,
        chartId: chart.getChartId()
      };

      exact[title] = record;

      const key = normalizeMonthlyQualityChartTitle_(title);
      if (!normalized[key]) {
        normalized[key] = record;
      }
      sourceTitles.push(title);
    });
  });

  return {
    exact: exact,
    normalized: normalized,
    sourceTitles: sourceTitles
  };
}

/**
 * Normalize harmless wording differences between the July canonical template
 * and repaired monthly workbooks without weakening matching enough to confuse
 * different chart types.
 */
function normalizeMonthlyQualityChartTitle_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\u2013|\u2014/g, '-')
    .replace(/12\s*-?\s*mth/g, '12 month')
    .replace(/12\s*-?\s*month/g, '12 month')
    .replace(/\ball\s+lines\b/g, 'all')
    .replace(/\s+-\s+filtered\b/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isMonthlyQualityReportChartTitle_(title) {
  const text = String(title || '').toLowerCase();
  return (
    text.indexOf('rolling dppm') !== -1 ||
    text.indexOf('service tickets') !== -1 ||
    text.indexOf('quality pipeline') !== -1
  );
}

/**
 * Replace every linked report chart in Slides with the corresponding chart
 * object from the current monthly spreadsheet. Geometry is copied from the
 * existing slide element, so the July-approved slide layout is preserved.
 */
function rebindMonthlyQualityReportCharts_(spreadsheet, presentation) {
  const index = buildMonthlyQualityReportChartIndex_(spreadsheet);
  const spreadsheetId = spreadsheet.getId();
  let rebound = 0;
  let alreadyCurrent = 0;
  const unmatched = [];

  presentation.getSlides().forEach(function(slide) {
    const elements = slide.getPageElements();

    elements.forEach(function(element) {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) {
        return;
      }

      const title = String(element.getTitle ? element.getTitle() : '').trim();
      if (!isMonthlyQualityReportChartTitle_(title)) return;

      const source = index.exact[title] ||
        index.normalized[normalizeMonthlyQualityChartTitle_(title)];

      if (!source) {
        unmatched.push(title || '(untitled linked chart)');
        return;
      }

      const linked = element.asSheetsChart();
      let currentSpreadsheetId = '';
      let currentChartId = null;

      try {
        currentSpreadsheetId = String(linked.getSpreadsheetId() || '');
        currentChartId = linked.getChartId();
      } catch (error) {
        currentSpreadsheetId = '';
        currentChartId = null;
      }

      if (
        currentSpreadsheetId === spreadsheetId &&
        Number(currentChartId) === Number(source.chartId)
      ) {
        alreadyCurrent++;
        return;
      }

      const left = element.getLeft();
      const top = element.getTop();
      const width = element.getWidth();
      const height = element.getHeight();

      element.remove();
      slide.insertSheetsChart(
        source.chart,
        left,
        top,
        width,
        height
      );
      rebound++;

      Logger.log(
        'Rebound report chart: ' + title +
        ' -> ' + source.sheetName + ' / chart ' + source.chartId
      );
    });
  });

  if (unmatched.length) {
    throw new Error(
      'Monthly report contains linked charts that could not be matched to the current workbook: ' +
      unmatched.join(' | ') +
      '. Available source chart titles: ' + index.sourceTitles.join(' | ')
    );
  }

  Logger.log(JSON.stringify({
    status: 'READY',
    chartReboundCount: rebound,
    chartAlreadyCurrentCount: alreadyCurrent,
    spreadsheetId: spreadsheetId
  }));

  return {
    status: 'READY',
    rebound: rebound,
    alreadyCurrent: alreadyCurrent
  };
}

/**
 * Final fail-safe: every report-linked DPPM/service/quality chart in the deck
 * must point to the current monthly workbook before the run may succeed.
 */
function verifyMonthlyQualityReportChartBindings_(presentation, spreadsheetId) {
  const wrongWorkbook = [];
  let reportChartCount = 0;

  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) {
        return;
      }

      const title = String(element.getTitle ? element.getTitle() : '').trim();
      if (!isMonthlyQualityReportChartTitle_(title)) return;

      reportChartCount++;
      const linked = element.asSheetsChart();
      const linkedSpreadsheetId = String(linked.getSpreadsheetId() || '');

      if (linkedSpreadsheetId !== spreadsheetId) {
        wrongWorkbook.push(
          title + ' -> ' + (linkedSpreadsheetId || '(no spreadsheet id)')
        );
      }
    });
  });

  if (wrongWorkbook.length) {
    throw new Error(
      'Monthly report chart binding verification failed. Charts still linked to the wrong workbook: ' +
      wrongWorkbook.join(' | ')
    );
  }

  if (reportChartCount < 19) {
    throw new Error(
      'Monthly report chart binding verification found only ' + reportChartCount +
      ' report charts; expected at least 19 DPPM/service/quality linked charts.'
    );
  }

  Logger.log(
    'Monthly report chart binding verification READY: ' +
    reportChartCount + ' report charts linked to current workbook.'
  );
  return true;
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
