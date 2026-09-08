/**
 * Temporary manual-HubSpot mode for the Monthly Quality report.
 *
 * Until the Apps Script HubSpot API authentication is repaired, the monthly
 * report must NOT call HubSpot directly. The manually audited counts below are
 * the temporary source of truth for months that have been reviewed.
 *
 * This file intentionally sorts after the live-sync data guard so this version
 * of populateMonthlyQualityDataStage_() wins while manual mode is active.
 */

const MONTHLY_QUALITY_MANUAL_HUBSPOT_OVERRIDES_ = Object.freeze({
  '2026-08': Object.freeze({
    'MSC': Object.freeze({ startup: 2, warranty: 4, service: 1 }),
    'ARU': Object.freeze({ startup: 0, warranty: 5, service: 0 }),
    'CSC': Object.freeze({ startup: 1, warranty: 4, service: 0 }),
    'Mods': Object.freeze({ startup: 0, warranty: 0, service: 0 }),
    'Gas Heat': Object.freeze({ startup: 0, warranty: 0, service: 0 }),
    'Coatings': Object.freeze({ startup: 0, warranty: 1, service: 0 }),
    'Bard Coatings': Object.freeze({ startup: 0, warranty: 0, service: 0 })
  })
});

function populateMonthlyQualityDataStage_(context, dataStage) {
  const spreadsheetId = dataStage.dataFile.getId();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  Logger.log(
    'Monthly Quality: HubSpot API sync intentionally bypassed; using manually audited values.'
  );

  // Read the existing 12-month table first so prior months remain unchanged.
  const issueChartData = readMonthlyQualityExistingIssueChartData_(spreadsheet);
  if (!issueChartData || issueChartData.status !== 'READY') {
    throw new Error(
      'Monthly Quality data stage stopped: manually validated HubSpot chart data is not READY.'
    );
  }

  // If the report month has been manually audited, apply that audited month to
  // the workbook source tables before any chart or DPPM refresh occurs.
  applyMonthlyQualityManualHubSpotOverride_(spreadsheet, issueChartData);

  // Keep DPPM Inputs aligned with the same manually validated values used by
  // the service-ticket summary tables and charts. Units Shipped is preserved.
  syncMonthlyQualityDPPMInputsFromExistingSummary_(spreadsheet, issueChartData);

  ensureMonthlyQualityPackageDPPMModel_(spreadsheet, context);
  SpreadsheetApp.flush();

  validateMonthlyQualityCurrentMonthDPPMInputs_(spreadsheet, issueChartData);

  const configSheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_
  );
  if (configSheet) {
    configSheet.getRange('B12').setValue(
      'Manual HubSpot mode: audited values used; live HubSpot API skipped'
    );
  }

  return {
    spreadsheetId: spreadsheetId,
    spreadsheetName: dataStage.dataFile.getName(),
    hubSpotSync: {
      status: 'SKIPPED',
      reason: 'Temporary manual HubSpot mode until Apps Script API authentication is repaired'
    },
    issueChartData: issueChartData,
    dppmInputValidation: 'READY',
    reportCreatedAfterDataStage: true
  };
}

/**
 * Apply a manually audited report-month override to HubSpot Chart Data and the
 * in-memory summary that is later written to Slides.
 */
function applyMonthlyQualityManualHubSpotOverride_(spreadsheet, issueResult) {
  const reportMonthKey = String(issueResult.reportMonth || '').trim();
  const override = MONTHLY_QUALITY_MANUAL_HUBSPOT_OVERRIDES_[reportMonthKey];
  if (!override) {
    Logger.log(
      'Manual HubSpot mode: no explicit override for ' + reportMonthKey +
      '; retaining existing validated workbook values.'
    );
    return false;
  }

  const chartDataSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_CHART_DATA_SHEET_
  );
  if (!chartDataSheet) {
    throw new Error('Manual HubSpot mode is missing "HubSpot Chart Data".');
  }

  const reportMonth = validatedIssueMonthStart_(reportMonthKey + '-01');
  if (!reportMonth) {
    throw new Error('Invalid manual HubSpot override month: ' + reportMonthKey);
  }

  VALIDATED_ISSUE_PRODUCT_LINES_.forEach(function(product) {
    const values = override[product] || { startup: 0, warranty: 0, service: 0 };
    const record = {
      startup: Number(values.startup) || 0,
      warranty: Number(values.warranty) || 0,
      service: Number(values.service) || 0
    };
    record.total = record.startup + record.warranty + record.service;

    if (!issueResult.summaries[product]) {
      issueResult.summaries[product] = {};
    }
    issueResult.summaries[product].current = record;

    const block = VALIDATED_ISSUE_BREAKDOWN_BLOCKS_.filter(function(item) {
      return item.product === product;
    })[0];

    if (block) {
      const monthValues = chartDataSheet
        .getRange(2, block.startColumn, 12, 1)
        .getValues();
      let targetRow = null;

      monthValues.forEach(function(row, index) {
        const month = validatedIssueMonthStart_(row[0]);
        if (month && validatedIssueMonthKey_(month) === reportMonthKey) {
          targetRow = index + 2;
        }
      });

      if (!targetRow) {
        throw new Error(
          'HubSpot Chart Data is missing ' + reportMonthKey + ' for ' + product + '.'
        );
      }

      chartDataSheet.getRange(targetRow, block.startColumn + 1, 1, 3).setValues([[
        record.startup,
        record.warranty,
        record.service
      ]]);
    }
  });

  // A:H contains monthly total countable tickets by product.
  const summaryMonths = chartDataSheet.getRange(2, 1, 12, 1).getValues();
  let summaryRow = null;
  summaryMonths.forEach(function(row, index) {
    const month = validatedIssueMonthStart_(row[0]);
    if (month && validatedIssueMonthKey_(month) === reportMonthKey) {
      summaryRow = index + 2;
    }
  });

  if (!summaryRow) {
    throw new Error('HubSpot Chart Data summary is missing ' + reportMonthKey + '.');
  }

  const totals = VALIDATED_ISSUE_PRODUCT_LINES_.map(function(product) {
    const record = issueResult.summaries[product].current;
    return Number(record.total) || 0;
  });
  chartDataSheet.getRange(summaryRow, 2, 1, totals.length).setValues([totals]);

  // Plant reporting intentionally includes MSC + ARU + CSC only.
  const plant = { startup: 0, warranty: 0, service: 0, total: 0 };
  ['MSC', 'ARU', 'CSC'].forEach(function(product) {
    const record = issueResult.summaries[product].current;
    plant.startup += Number(record.startup) || 0;
    plant.warranty += Number(record.warranty) || 0;
    plant.service += Number(record.service) || 0;
    plant.total += Number(record.total) || 0;
  });
  issueResult.summaries['All Lines'].current = plant;

  issueResult.source = 'Manual audited HubSpot override';
  issueResult.method =
    'Support Pipeline + Startup/Warranty/Service + MJC No Fault blank; manually audited while API offline';

  SpreadsheetApp.flush();
  Logger.log(
    'Manual HubSpot override applied for ' + reportMonthKey +
    ': All Lines=' + plant.total +
    ' (' + plant.startup + '/' + plant.warranty + '/' + plant.service + ').'
  );
  return true;
}

/**
 * Copy the report month's manually validated Startup / Warranty / Service
 * summary into DPPM Inputs. Units Shipped is preserved.
 */
function syncMonthlyQualityDPPMInputsFromExistingSummary_(spreadsheet, issueResult) {
  const inputSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_INPUTS_SHEET_
  );
  if (!inputSheet) {
    throw new Error('Manual HubSpot mode is missing "DPPM Inputs".');
  }

  const reportMonthKey = String(issueResult.reportMonth || '').trim();
  if (!reportMonthKey) {
    throw new Error('Manual HubSpot mode could not determine the report month.');
  }

  const reportMonth = validatedIssueMonthStart_(reportMonthKey + '-01');
  if (!reportMonth) {
    throw new Error('Manual HubSpot mode has an invalid report month: ' + reportMonthKey);
  }

  const products = VALIDATED_ISSUE_PRODUCT_LINES_.slice();
  const lastRow = inputSheet.getLastRow();
  const data = lastRow >= 2
    ? inputSheet.getRange(2, 1, lastRow - 1, 6).getValues()
    : [];

  products.forEach(function(product) {
    const matches = [];

    data.forEach(function(row, index) {
      const month = validatedIssueMonthStart_(row[0]);
      if (
        month &&
        validatedIssueMonthKey_(month) === reportMonthKey &&
        String(row[1] || '').trim() === product
      ) {
        matches.push(index + 2);
      }
    });

    if (matches.length > 1) {
      throw new Error(
        'Manual HubSpot mode found duplicate DPPM Inputs rows for ' +
        reportMonthKey + ' / ' + product + '.'
      );
    }

    const summary = issueResult.summaries[product] &&
      issueResult.summaries[product].current
      ? issueResult.summaries[product].current
      : { startup: 0, warranty: 0, service: 0 };

    if (matches.length === 1) {
      inputSheet.getRange(matches[0], 4, 1, 3).setValues([[
        Number(summary.startup) || 0,
        Number(summary.warranty) || 0,
        Number(summary.service) || 0
      ]]);
    } else {
      inputSheet.appendRow([
        new Date(reportMonth.getTime()),
        product,
        '',
        Number(summary.startup) || 0,
        Number(summary.warranty) || 0,
        Number(summary.service) || 0
      ]);
    }
  });

  SpreadsheetApp.flush();
  return true;
}
