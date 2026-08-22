/**
 * Builds the monthly service/quality issue chart source directly from the
 * synchronized HubSpot Support Pipeline ticket data.
 *
 * Counting rule (mirrors the HubSpot quality reports):
 * - ticket is in the Support Pipeline (the raw sync already enforces this)
 * - Product matches the product chart
 * - primary Ticket Category is Startup, Warranty, or Service
 * - MJC No Fault is NOT set
 *
 * The exact same filtered counts are also written back to DPPM Inputs for
 * products that already have a DPPM denominator row. This keeps the service
 * ticket charts, current/previous month tables, and DPPM calculations on the
 * same HubSpot source-of-truth method.
 */

const VALIDATED_ISSUE_CHART_DATA_SHEET_ = 'HubSpot Chart Data';
const VALIDATED_ISSUE_HUBSPOT_RAW_SHEET_ = 'HubSpot Support Tickets';
const VALIDATED_ISSUE_DPPM_INPUTS_SHEET_ = 'DPPM Inputs';
const VALIDATED_ISSUE_DPPM_CONFIG_SHEET_ = 'DPPM Config';

const VALIDATED_ISSUE_COUNTED_CATEGORIES_ = Object.freeze([
  'Startup',
  'Warranty',
  'Service'
]);

const VALIDATED_ISSUE_PRODUCT_LINES_ = Object.freeze([
  'MSC',
  'ARU',
  'CSC',
  'Mods',
  'Gas Heat',
  'Coatings',
  'Bard Coatings'
]);

const VALIDATED_ISSUE_BREAKDOWN_BLOCKS_ = Object.freeze([
  { product: 'MSC', startColumn: 9 },             // I:L
  { product: 'ARU', startColumn: 14 },            // N:Q
  { product: 'CSC', startColumn: 19 },            // S:V
  { product: 'Mods', startColumn: 24 },           // X:AA
  { product: 'Gas Heat', startColumn: 29 },       // AC:AF
  { product: 'Coatings', startColumn: 34 },       // AH:AK
  { product: 'Bard Coatings', startColumn: 39 }   // AM:AP
]);

/**
 * Refresh the 12-month issue chart source from the complete synchronized
 * HubSpot Support Pipeline dataset and synchronize DPPM Inputs to the same
 * MJC-No-Fault exclusion logic.
 */
function refreshMonthlyQualityValidatedIssueChartData_(spreadsheet) {
  const rawSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_HUBSPOT_RAW_SHEET_
  );
  if (!rawSheet) {
    throw new Error(
      'HubSpot ticket source is missing "' +
      VALIDATED_ISSUE_HUBSPOT_RAW_SHEET_ + '". Run the HubSpot sync first.'
    );
  }

  const configSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_DPPM_CONFIG_SHEET_
  );
  if (!configSheet) {
    throw new Error('Issue chart source is missing "DPPM Config".');
  }

  let chartDataSheet = spreadsheet.getSheetByName(
    VALIDATED_ISSUE_CHART_DATA_SHEET_
  );
  if (!chartDataSheet) {
    chartDataSheet = spreadsheet.insertSheet(VALIDATED_ISSUE_CHART_DATA_SHEET_);
    chartDataSheet.hideSheet();
  }

  const requiredColumns = VALIDATED_ISSUE_BREAKDOWN_BLOCKS_[
    VALIDATED_ISSUE_BREAKDOWN_BLOCKS_.length - 1
  ].startColumn + 3;
  if (chartDataSheet.getMaxColumns() < requiredColumns) {
    chartDataSheet.insertColumnsAfter(
      chartDataSheet.getMaxColumns(),
      requiredColumns - chartDataSheet.getMaxColumns()
    );
  }

  const rawData = rawSheet.getDataRange().getValues();
  if (rawData.length < 1) {
    throw new Error('HubSpot Support Tickets does not contain a header row.');
  }

  const headers = rawData[0].map(function(value) {
    return String(value || '').trim();
  });
  const columns = validatedIssueHeaderMap_(headers);

  [
    'Created Date',
    'Ticket Category',
    'Product',
    'MJC No Fault'
  ].forEach(function(header) {
    if (columns[header] === undefined) {
      throw new Error(
        'HubSpot Support Tickets is missing required column: ' + header
      );
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

  const monthKeys = {};
  months.forEach(function(month) {
    monthKeys[validatedIssueMonthKey_(month)] = true;
  });

  const byMonthAndProduct = {};
  let countableTicketCount = 0;
  let excludedNoFaultCount = 0;
  let excludedCategoryCount = 0;

  rawData.slice(1).forEach(function(row) {
    const month = validatedIssueMonthStart_(row[columns['Created Date']]);
    if (!month) return;

    const monthKey = validatedIssueMonthKey_(month);
    if (!monthKeys[monthKey]) return;

    const product = String(row[columns['Product']] || '').trim();
    const category = String(row[columns['Ticket Category']] || '').trim();
    const noFault = String(row[columns['MJC No Fault']] || '').trim();

    if (noFault) {
      excludedNoFaultCount++;
      return;
    }

    if (VALIDATED_ISSUE_COUNTED_CATEGORIES_.indexOf(category) === -1) {
      excludedCategoryCount++;
      return;
    }

    // HubSpot currently has a single Coatings product enum and no separate
    // Bard Coatings service enum. Bard Coatings therefore remains zero until
    // a distinct HubSpot classification exists; do not duplicate Coatings.
    if (product === 'Bard Coatings') return;
    if (VALIDATED_ISSUE_PRODUCT_LINES_.indexOf(product) === -1) return;

    const key = monthKey + '|' + product;
    if (!byMonthAndProduct[key]) {
      byMonthAndProduct[key] = {
        startup: 0,
        warranty: 0,
        service: 0,
        total: 0
      };
    }

    const record = byMonthAndProduct[key];
    if (category === 'Startup') record.startup++;
    if (category === 'Warranty') record.warranty++;
    if (category === 'Service') record.service++;
    record.total++;
    countableTicketCount++;
  });

  // A:H = monthly countable HubSpot tickets by product.
  const productSummary = [
    ['Month'].concat(VALIDATED_ISSUE_PRODUCT_LINES_)
  ];

  months.forEach(function(month) {
    const monthKey = validatedIssueMonthKey_(month);
    const row = [month];
    VALIDATED_ISSUE_PRODUCT_LINES_.forEach(function(product) {
      const record = byMonthAndProduct[monthKey + '|' + product];
      row.push(record ? record.total : 0);
    });
    productSummary.push(row);
  });

  const productSummaryWidth = productSummary[0].length;
  chartDataSheet.getRange(1, 1, 13, productSummaryWidth).clearContent();
  chartDataSheet.getRange(1, 1, productSummary.length, productSummaryWidth)
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
        record ? record.startup : 0,
        record ? record.warranty : 0,
        record ? record.service : 0
      ]);
    });

    chartDataSheet.getRange(1, block.startColumn, 13, 4).clearContent();
    chartDataSheet.getRange(1, block.startColumn, rows.length, 4)
      .setValues(rows);
    chartDataSheet.getRange(2, block.startColumn, 12, 1)
      .setNumberFormat('mmm yyyy');
  });

  const dppmInputResult = syncMonthlyQualityDPPMInputsFromHubSpot_(
    spreadsheet,
    months,
    byMonthAndProduct
  );

  SpreadsheetApp.flush();

  const currentMonth = reportMonth;
  const previousMonth = new Date(
    reportMonth.getFullYear(),
    reportMonth.getMonth() - 1,
    1
  );
  const summaries = buildMonthlyQualityIssueSummaries_(
    currentMonth,
    previousMonth,
    byMonthAndProduct
  );

  return {
    status: 'READY',
    source: VALIDATED_ISSUE_HUBSPOT_RAW_SHEET_,
    method: 'Support Pipeline + Product + Startup/Warranty/Service + MJC No Fault blank',
    chartDataSheet: VALIDATED_ISSUE_CHART_DATA_SHEET_,
    reportMonth: validatedIssueMonthKey_(reportMonth),
    currentMonthLabel: Utilities.formatDate(reportMonth, 'UTC', 'MMM'),
    previousMonthLabel: Utilities.formatDate(previousMonth, 'UTC', 'MMM'),
    monthsWritten: months.length,
    productLines: VALIDATED_ISSUE_PRODUCT_LINES_.slice(),
    countableTicketCount: countableTicketCount,
    excludedNoFaultCount: excludedNoFaultCount,
    excludedCategoryCount: excludedCategoryCount,
    dppmInputsUpdated: dppmInputResult.rowsUpdated,
    summaries: summaries
  };
}

/**
 * Update only Startup/Warranty/Service columns on existing DPPM Inputs rows.
 * Unit shipment denominators are deliberately untouched. This is performed for
 * the same 12-month window shown in the report, so historical chart points in
 * that window use the same HubSpot filter method as the service-ticket charts.
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

  const lastRow = inputSheet.getLastRow();
  if (lastRow < 2) return { rowsUpdated: 0 };

  const data = inputSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const monthKeys = {};
  months.forEach(function(month) {
    monthKeys[validatedIssueMonthKey_(month)] = true;
  });

  const dppmProducts = {
    'MSC': true,
    'ARU': true,
    'CSC': true,
    'Mods': true,
    'Gas Heat': true,
    'Coatings': true,
    'Bard Coatings': true
  };

  const output = [];
  let rowsUpdated = 0;

  data.forEach(function(row) {
    const month = validatedIssueMonthStart_(row[0]);
    const product = String(row[1] || '').trim();
    let startup = row[3];
    let warranty = row[4];
    let service = row[5];

    if (month && monthKeys[validatedIssueMonthKey_(month)] && dppmProducts[product]) {
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
  return { rowsUpdated: rowsUpdated };
}

function buildMonthlyQualityIssueSummaries_(
  currentMonth,
  previousMonth,
  byMonthAndProduct
) {
  const summaries = {};
  VALIDATED_ISSUE_PRODUCT_LINES_.forEach(function(product) {
    summaries[product] = {
      current: monthlyQualityIssueRecord_(
        byMonthAndProduct,
        currentMonth,
        product
      ),
      previous: monthlyQualityIssueRecord_(
        byMonthAndProduct,
        previousMonth,
        product
      )
    };
  });

  const allProductsForDPPM = ['MSC', 'ARU', 'CSC'];
  summaries['All Lines'] = {
    current: monthlyQualityIssueSumProducts_(
      byMonthAndProduct,
      currentMonth,
      allProductsForDPPM
    ),
    previous: monthlyQualityIssueSumProducts_(
      byMonthAndProduct,
      previousMonth,
      allProductsForDPPM
    )
  };
  return summaries;
}

function monthlyQualityIssueRecord_(byMonthAndProduct, month, product) {
  const record = byMonthAndProduct[
    validatedIssueMonthKey_(month) + '|' + product
  ];
  return record ? {
    startup: record.startup,
    warranty: record.warranty,
    service: record.service,
    total: record.total
  } : {
    startup: 0,
    warranty: 0,
    service: 0,
    total: 0
  };
}

function monthlyQualityIssueSumProducts_(
  byMonthAndProduct,
  month,
  products
) {
  const total = { startup: 0, warranty: 0, service: 0, total: 0 };
  products.forEach(function(product) {
    const record = monthlyQualityIssueRecord_(
      byMonthAndProduct,
      month,
      product
    );
    total.startup += record.startup;
    total.warranty += record.warranty;
    total.service += record.service;
    total.total += record.total;
  });
  return total;
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
