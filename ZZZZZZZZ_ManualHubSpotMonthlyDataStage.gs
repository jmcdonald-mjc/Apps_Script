/**
 * Temporary manual-HubSpot mode for the Monthly Quality report.
 *
 * Until the Apps Script HubSpot API authentication is repaired, the monthly
 * report must NOT call HubSpot directly. The validated HubSpot counts already
 * stored in the monthly workbook are the temporary source of truth.
 *
 * This file intentionally sorts after the live-sync data guard so this version
 * of populateMonthlyQualityDataStage_() wins while manual mode is active.
 */

function populateMonthlyQualityDataStage_(context, dataStage) {
  const spreadsheetId = dataStage.dataFile.getId();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  Logger.log(
    'Monthly Quality: HubSpot API sync intentionally bypassed; using manually validated workbook values.'
  );

  // Read the manually validated Startup / Warranty / Service values already
  // stored in HubSpot Chart Data. Do not contact HubSpot here.
  const issueChartData = readMonthlyQualityExistingIssueChartData_(spreadsheet);
  if (!issueChartData || issueChartData.status !== 'READY') {
    throw new Error(
      'Monthly Quality data stage stopped: manually validated HubSpot chart data is not READY.'
    );
  }

  // Keep DPPM Inputs aligned with the same manually validated values used by
  // the report tables. This prevents the DPPM charts and the service-ticket
  // summary tables from disagreeing with one another.
  syncMonthlyQualityDPPMInputsFromExistingSummary_(spreadsheet, issueChartData);

  ensureMonthlyQualityPackageDPPMModel_(spreadsheet, context);
  SpreadsheetApp.flush();

  validateMonthlyQualityCurrentMonthDPPMInputs_(spreadsheet, issueChartData);

  const configSheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_
  );
  if (configSheet) {
    configSheet.getRange('B12').setValue(
      'Manual HubSpot mode: validated workbook values used; live HubSpot API skipped'
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
  let lastRow = inputSheet.getLastRow();
  let data = lastRow >= 2
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
