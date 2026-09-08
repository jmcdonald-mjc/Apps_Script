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
  const reportMonthKey = String(context.monthKey || '').trim();

  Logger.log(
    'Monthly Quality: HubSpot API sync intentionally bypassed; using manually audited values.'
  );

  // IMPORTANT: a July-based template does not yet contain an August row.
  // Roll the 12-month HubSpot chart-data window forward and inject the audited
  // current month BEFORE any code tries to read the report month.
  prepareMonthlyQualityManualHubSpotWindow_(spreadsheet, reportMonthKey);

  const issueChartData = readMonthlyQualityExistingIssueChartData_(spreadsheet);
  if (!issueChartData || issueChartData.status !== 'READY') {
    throw new Error(
      'Monthly Quality data stage stopped: manually validated HubSpot chart data is not READY.'
    );
  }

  // Keep DPPM Inputs aligned with exactly the same manually audited values used
  // by the service-ticket summary tables and charts. Units Shipped is preserved.
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
 * Rebuild HubSpot Chart Data to the exact 12 months ending in reportMonthKey.
 * Existing historical values are retained by month. The current report month
 * is taken from the manually audited override table above.
 */
function prepareMonthlyQualityManualHubSpotWindow_(spreadsheet, reportMonthKey) {
  const override = MONTHLY_QUALITY_MANUAL_HUBSPOT_OVERRIDES_[reportMonthKey];
  if (!override) {
    throw new Error(
      'Manual HubSpot mode has no audited values configured for ' +
      reportMonthKey + '. Add the month to MONTHLY_QUALITY_MANUAL_HUBSPOT_OVERRIDES_ before running.'
    );
  }

  const chartDataSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_CHART_DATA_SHEET_
  );
  if (!chartDataSheet) {
    throw new Error('Manual HubSpot mode is missing "HubSpot Chart Data".');
  }

  const reportMonth = validatedIssueMonthStart_(reportMonthKey + '-01');
  if (!reportMonth) {
    throw new Error('Invalid manual HubSpot report month: ' + reportMonthKey);
  }

  const months = [];
  for (let offset = 11; offset >= 0; offset--) {
    months.push(new Date(
      reportMonth.getFullYear(),
      reportMonth.getMonth() - offset,
      1
    ));
  }

  const currentByProduct = {};

  VALIDATED_ISSUE_BREAKDOWN_BLOCKS_.forEach(function(block) {
    const oldRows = chartDataSheet
      .getRange(2, block.startColumn, 12, 4)
      .getValues();
    const oldByMonth = {};

    oldRows.forEach(function(row) {
      const month = validatedIssueMonthStart_(row[0]);
      if (!month) return;
      oldByMonth[validatedIssueMonthKey_(month)] = {
        startup: Number(row[1]) || 0,
        warranty: Number(row[2]) || 0,
        service: Number(row[3]) || 0
      };
    });

    const newRows = months.map(function(month) {
      const key = validatedIssueMonthKey_(month);
      let record = oldByMonth[key] || { startup: 0, warranty: 0, service: 0 };

      if (key === reportMonthKey) {
        const audited = override[block.product] || {
          startup: 0,
          warranty: 0,
          service: 0
        };
        record = {
          startup: Number(audited.startup) || 0,
          warranty: Number(audited.warranty) || 0,
          service: Number(audited.service) || 0
        };
        currentByProduct[block.product] = record;
      }

      return [month, record.startup, record.warranty, record.service];
    });

    chartDataSheet
      .getRange(2, block.startColumn, 12, 4)
      .setValues(newRows);
    chartDataSheet
      .getRange(2, block.startColumn, 12, 1)
      .setNumberFormat('mmm yyyy');
  });

  // A:H is the total count by product. Rebuild it from the same breakdown
  // blocks so the product totals can never disagree with Startup/Warranty/Service.
  const productMaps = {};
  VALIDATED_ISSUE_BREAKDOWN_BLOCKS_.forEach(function(block) {
    const rows = chartDataSheet
      .getRange(2, block.startColumn, 12, 4)
      .getValues();
    productMaps[block.product] = {};

    rows.forEach(function(row) {
      const month = validatedIssueMonthStart_(row[0]);
      if (!month) return;
      productMaps[block.product][validatedIssueMonthKey_(month)] =
        (Number(row[1]) || 0) +
        (Number(row[2]) || 0) +
        (Number(row[3]) || 0);
    });
  });

  const summaryRows = months.map(function(month) {
    const key = validatedIssueMonthKey_(month);
    const row = [month];
    VALIDATED_ISSUE_PRODUCT_LINES_.forEach(function(product) {
      row.push(Number(productMaps[product] && productMaps[product][key]) || 0);
    });
    return row;
  });

  chartDataSheet
    .getRange(2, 1, 12, 1 + VALIDATED_ISSUE_PRODUCT_LINES_.length)
    .setValues(summaryRows);
  chartDataSheet.getRange(2, 1, 12, 1).setNumberFormat('mmm yyyy');

  SpreadsheetApp.flush();

  const plant = { startup: 0, warranty: 0, service: 0, total: 0 };
  ['MSC', 'ARU', 'CSC'].forEach(function(product) {
    const record = currentByProduct[product] || {
      startup: 0,
      warranty: 0,
      service: 0
    };
    plant.startup += Number(record.startup) || 0;
    plant.warranty += Number(record.warranty) || 0;
    plant.service += Number(record.service) || 0;
  });
  plant.total = plant.startup + plant.warranty + plant.service;

  Logger.log(
    'Manual HubSpot window prepared through ' + reportMonthKey +
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
