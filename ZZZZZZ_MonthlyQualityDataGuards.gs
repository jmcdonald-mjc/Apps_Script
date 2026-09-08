/**
 * Final monthly-quality data guards.
 *
 * This file intentionally sorts after the other monthly-report overrides.
 * It prevents a monthly report from being created when HubSpot issue counts
 * have not actually reached DPPM Inputs.
 */

/**
 * Populate HubSpot data before Slides are created and verify the data stage.
 */
function populateMonthlyQualityDataStage_(context, dataStage) {
  const spreadsheetId = dataStage.dataFile.getId();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const hubSpotSync = syncHubSpotSupportTicketsToSpreadsheet_(spreadsheetId);
  if (!hubSpotSync || hubSpotSync.status !== 'READY') {
    throw new Error(
      'Monthly Quality data stage stopped: HubSpot sync is not READY. ' +
      String(hubSpotSync && (hubSpotSync.reason || hubSpotSync.status) || 'Unknown error')
    );
  }

  const issueChartData = refreshMonthlyQualityValidatedIssueChartData_(spreadsheet);
  if (!issueChartData || issueChartData.status !== 'READY') {
    throw new Error('Monthly Quality data stage stopped: issue data refresh failed.');
  }

  // Rebuild the formulas only after HubSpot has populated DPPM Inputs.
  ensureMonthlyQualityPackageDPPMModel_(spreadsheet, context);
  SpreadsheetApp.flush();

  validateMonthlyQualityCurrentMonthDPPMInputs_(
    spreadsheet,
    issueChartData
  );

  const configSheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_
  );
  if (configSheet) {
    configSheet.getRange('B12').setValue(
      'Monthly sheet populated and validated before report creation'
    );
  }

  return {
    spreadsheetId: spreadsheetId,
    spreadsheetName: dataStage.dataFile.getName(),
    hubSpotSync: hubSpotSync,
    issueChartData: issueChartData,
    dppmInputValidation: 'READY',
    reportCreatedAfterDataStage: true
  };
}

/**
 * Synchronize HubSpot issue counts to DPPM Inputs.
 *
 * The previous implementation only updated rows that already existed. A copied
 * monthly template can legitimately stop at the previous month, which meant the
 * new report month could be absent and therefore silently receive no HubSpot
 * counts. This implementation guarantees one current-month row per product,
 * preserves Units Shipped, updates the 12-month issue columns, and verifies the
 * current-month write before returning.
 */
function syncMonthlyQualityDPPMInputsFromHubSpot_(
  spreadsheet,
  months,
  byMonthAndProduct
) {
  const inputSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_INPUTS_SHEET_
  );
  if (!inputSheet) {
    throw new Error('Issue source is missing "DPPM Inputs".');
  }
  if (!months || !months.length) {
    throw new Error('No report months were supplied for DPPM input synchronization.');
  }

  const reportMonth = validatedIssueMonthStart_(months[months.length - 1]);
  const reportMonthKey = validatedIssueMonthKey_(reportMonth);
  const products = VALIDATED_ISSUE_PRODUCT_LINES_.slice();

  // Guarantee the copied template has a row for the new report month.
  const existingLastRow = inputSheet.getLastRow();
  const existing = existingLastRow >= 2
    ? inputSheet.getRange(2, 1, existingLastRow - 1, 2).getValues()
    : [];
  const currentMonthRows = {};

  existing.forEach(function(row) {
    const month = validatedIssueMonthStart_(row[0]);
    const product = String(row[1] || '').trim();
    if (!month || validatedIssueMonthKey_(month) !== reportMonthKey) return;
    if (products.indexOf(product) === -1) return;
    currentMonthRows[product] = (currentMonthRows[product] || 0) + 1;
  });

  const duplicateProducts = products.filter(function(product) {
    return (currentMonthRows[product] || 0) > 1;
  });
  if (duplicateProducts.length) {
    throw new Error(
      'Duplicate DPPM Inputs rows for ' + reportMonthKey + ': ' +
      duplicateProducts.join(', ')
    );
  }

  const missingRows = products.filter(function(product) {
    return !currentMonthRows[product];
  }).map(function(product) {
    // A Month, B Product, C Units Shipped, D Startup, E Warranty, F Service.
    // Units Shipped is intentionally left blank for the shipment-data routine.
    return [new Date(reportMonth.getTime()), product, '', 0, 0, 0];
  });

  if (missingRows.length) {
    const startRow = inputSheet.getLastRow() + 1;
    inputSheet.getRange(startRow, 1, missingRows.length, 6).setValues(missingRows);
    inputSheet.getRange(startRow, 1, missingRows.length, 1).setNumberFormat('mmm yyyy');
  }

  const lastRow = inputSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('DPPM Inputs has no data rows after current-month setup.');
  }

  const data = inputSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const monthKeys = {};
  months.forEach(function(month) {
    monthKeys[validatedIssueMonthKey_(month)] = true;
  });

  const output = [];
  let rowsUpdated = 0;

  data.forEach(function(row) {
    const month = validatedIssueMonthStart_(row[0]);
    const product = String(row[1] || '').trim();
    let startup = row[3];
    let warranty = row[4];
    let service = row[5];

    if (
      month &&
      monthKeys[validatedIssueMonthKey_(month)] &&
      products.indexOf(product) !== -1
    ) {
      const record = byMonthAndProduct[
        validatedIssueMonthKey_(month) + '|' + product
      ];
      startup = record ? record.startup : 0;
      warranty = record ? record.warranty : 0;
      service = record ? record.service : 0;
      rowsUpdated++;
    }

    output.push([startup, warranty, service]);
  });

  inputSheet.getRange(2, 4, output.length, 3).setValues(output);
  SpreadsheetApp.flush();

  // Read back the current month. Never allow a silent zero/missing write.
  const readback = inputSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const verifiedProducts = [];

  products.forEach(function(product) {
    const matches = readback.filter(function(row) {
      const month = validatedIssueMonthStart_(row[0]);
      return month &&
        validatedIssueMonthKey_(month) === reportMonthKey &&
        String(row[1] || '').trim() === product;
    });

    if (matches.length !== 1) {
      throw new Error(
        'Expected exactly one DPPM Inputs row for ' + reportMonthKey +
        ' / ' + product + '; found ' + matches.length + '.'
      );
    }

    const expected = byMonthAndProduct[reportMonthKey + '|' + product] || {
      startup: 0,
      warranty: 0,
      service: 0
    };
    const actual = matches[0];

    if (
      Number(actual[3] || 0) !== Number(expected.startup || 0) ||
      Number(actual[4] || 0) !== Number(expected.warranty || 0) ||
      Number(actual[5] || 0) !== Number(expected.service || 0)
    ) {
      throw new Error(
        'DPPM Inputs verification failed for ' + reportMonthKey + ' / ' + product +
        '. Expected Startup/Warranty/Service ' +
        [expected.startup || 0, expected.warranty || 0, expected.service || 0].join('/') +
        ', found ' + [actual[3] || 0, actual[4] || 0, actual[5] || 0].join('/') + '.'
      );
    }

    verifiedProducts.push(product);
  });

  return {
    rowsUpdated: rowsUpdated,
    rowsCreated: missingRows.length,
    reportMonth: reportMonthKey,
    verifiedProducts: verifiedProducts
  };
}

/**
 * Cross-check the report month's DPPM Inputs against the exact HubSpot summary
 * that will be used in the Slides report. MSC/ARU/CSC are mandatory because
 * they drive the plant DPPM metric.
 */
function validateMonthlyQualityCurrentMonthDPPMInputs_(spreadsheet, issueResult) {
  const inputSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_INPUTS_SHEET_
  );
  if (!inputSheet) throw new Error('Missing DPPM Inputs during validation.');

  const reportMonthKey = String(issueResult.reportMonth || '').trim();
  const requiredProducts = ['MSC', 'ARU', 'CSC'];
  const data = inputSheet.getRange(2, 1, Math.max(1, inputSheet.getLastRow() - 1), 6)
    .getValues();

  requiredProducts.forEach(function(product) {
    const rows = data.filter(function(row) {
      const month = validatedIssueMonthStart_(row[0]);
      return month &&
        validatedIssueMonthKey_(month) === reportMonthKey &&
        String(row[1] || '').trim() === product;
    });

    if (rows.length !== 1) {
      throw new Error(
        'Monthly report blocked: expected one DPPM Inputs row for ' +
        reportMonthKey + ' / ' + product + ', found ' + rows.length + '.'
      );
    }

    const expected = issueResult.summaries[product].current;
    const row = rows[0];
    if (
      Number(row[3] || 0) !== Number(expected.startup || 0) ||
      Number(row[4] || 0) !== Number(expected.warranty || 0) ||
      Number(row[5] || 0) !== Number(expected.service || 0)
    ) {
      throw new Error(
        'Monthly report blocked: HubSpot and DPPM Inputs disagree for ' +
        reportMonthKey + ' / ' + product + '.'
      );
    }
  });

  return true;
}
