/**
 * Builds the monthly service/quality issue chart source from the validated
 * DPPM model rather than from raw HubSpot ticket counts.
 *
 * IMPORTANT:
 * - DPPM Monthly is the controlled source for these charts.
 * - Startup / Warranty / Service counts therefore reflect the reviewed issue
 *   counts used by the quality report, including MJC responsibility decisions,
 *   deduplication and carryover handling already represented in DPPM Inputs.
 * - Raw HubSpot tickets are NOT counted directly here.
 */

const VALIDATED_ISSUE_CHART_DATA_SHEET_ = 'HubSpot Chart Data';
const VALIDATED_ISSUE_DPPM_MONTHLY_SHEET_ = 'DPPM Monthly';
const VALIDATED_ISSUE_DPPM_CONFIG_SHEET_ = 'DPPM Config';

const VALIDATED_ISSUE_PRODUCT_LINES_ = Object.freeze([
  'MSC',
  'ARU',
  'CSC',
  'Mods',
  'Gas Heat',
  'Coatings'
]);

const VALIDATED_ISSUE_BREAKDOWN_BLOCKS_ = Object.freeze([
  { product: 'MSC', startColumn: 9 },       // I:L
  { product: 'ARU', startColumn: 14 },      // N:Q
  { product: 'CSC', startColumn: 19 },      // S:V
  { product: 'Mods', startColumn: 24 },     // X:AA
  { product: 'Gas Heat', startColumn: 29 }, // AC:AF
  { product: 'Coatings', startColumn: 34 }  // AH:AK
]);

/**
 * Refresh the 12-month service-ticket chart source tables from DPPM Monthly.
 * Existing pipeline-status data below row 15 is intentionally left alone.
 */
function refreshMonthlyQualityValidatedIssueChartData_(spreadsheet) {
  const monthlySheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_MONTHLY_SHEET_
  );
  if (!monthlySheet) {
    throw new Error('Validated issue chart source is missing "DPPM Monthly".');
  }

  const configSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_CONFIG_SHEET_
  );
  if (!configSheet) {
    throw new Error('Validated issue chart source is missing "DPPM Config".');
  }

  let chartDataSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_CHART_DATA_SHEET_
  );
  if (!chartDataSheet) {
    chartDataSheet = spreadsheet.insertSheet(VALIDATED_ISSUE_CHART_DATA_SHEET_);
    chartDataSheet.hideSheet();
  }

  const data = monthlySheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('DPPM Monthly does not contain any data rows.');
  }

  const headers = data[0].map(function(value) {
    return String(value || '').trim();
  });
  const columns = validatedIssueHeaderMap_(headers);

  const requiredHeaders = [
    'Month',
    'Product Line',
    'Startup Issues',
    'Warranty Issues',
    'Service Issues',
    'Total Issues'
  ];
  requiredHeaders.forEach(function(header) {
    if (columns[header] === undefined) {
      throw new Error('DPPM Monthly is missing required column: ' + header);
    }
  });

  const reportMonthValue = configSheet.getRange('B7').getValue();
  const reportMonth = validatedIssueMonthStart_(reportMonthValue);
  if (!reportMonth) {
    throw new Error('DPPM Config!B7 does not contain a valid report month.');
  }

  const months = [];
  for (let offset = 11; offset >= 0; offset--) {
    months.push(new Date(
      reportMonth.getFullYear(),
      reportMonth.getMonth() - offset,
      1
    ));
  }

  const byMonthAndProduct = {};
  data.slice(1).forEach(function(row) {
    const month = validatedIssueMonthStart_(row[columns['Month']]);
    const product = String(row[columns['Product Line']] || '').trim();
    if (!month || !product) return;

    const key = validatedIssueMonthKey_(month) + '|' + product;
    const startup = validatedIssueNumberOrBlank_(
      row[columns['Startup Issues']]
    );
    const warranty = validatedIssueNumberOrBlank_(
      row[columns['Warranty Issues']]
    );
    const service = validatedIssueNumberOrBlank_(
      row[columns['Service Issues']]
    );
    const total = validatedIssueNumberOrBlank_(
      row[columns['Total Issues']]
    );

    byMonthAndProduct[key] = {
      startup: startup,
      warranty: warranty,
      service: service,
      total: total
    };
  });

  // A:G = monthly validated issue totals by product.
  // This drives the all-products service-ticket chart.
  const productSummary = [
    ['Month'].concat(VALIDATED_ISSUE_PRODUCT_LINES_)
  ];

  months.forEach(function(month) {
    const monthKey = validatedIssueMonthKey_(month);
    const row = [month];
    VALIDATED_ISSUE_PRODUCT_LINES_.forEach(function(product) {
      const record = byMonthAndProduct[monthKey + '|' + product];
      row.push(record ? record.total : '');
    });
    productSummary.push(row);
  });

  chartDataSheet.getRange(1, 1, 13, 7).clearContent();
  chartDataSheet.getRange(1, 1, productSummary.length, 7)
    .setValues(productSummary);
  chartDataSheet.getRange(2, 1, 12, 1).setNumberFormat('mmm yyyy');

  // Individual product chart blocks.
  VALIDATED_ISSUE_BREAKDOWN_BLOCKS_.forEach(function(block) {
    const rows = [['Month', 'Startup', 'Warranty', 'Service']];

    months.forEach(function(month) {
      const record = byMonthAndProduct[
        validatedIssueMonthKey_(month) + '|' + block.product
      ];
      rows.push([
        month,
        record ? record.startup : '',
        record ? record.warranty : '',
        record ? record.service : ''
      ]);
    });

    chartDataSheet.getRange(1, block.startColumn, 13, 4).clearContent();
    chartDataSheet.getRange(1, block.startColumn, rows.length, 4)
      .setValues(rows);
    chartDataSheet.getRange(2, block.startColumn, 12, 1)
      .setNumberFormat('mmm yyyy');
  });

  SpreadsheetApp.flush();

  return {
    status: 'READY',
    source: VALIDATED_ISSUE_DPPM_MONTHLY_SHEET_,
    chartDataSheet: VALIDATED_ISSUE_CHART_DATA_SHEET_,
    reportMonth: validatedIssueMonthKey_(reportMonth),
    monthsWritten: months.length,
    productLines: VALIDATED_ISSUE_PRODUCT_LINES_.slice()
  };
}

function validatedIssueHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    if (header) map[header] = index;
  });
  return map;
}

function validatedIssueMonthStart_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), 1);
  }
  if (!value) return null;

  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), 1);
}

function validatedIssueMonthKey_(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM');
}

function validatedIssueNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const number = Number(value);
  return isNaN(number) ? '' : number;
}
