/**
 * Override for monthly quality DPPM slide chart replacement.
 *
 * The original monthly package routine targeted the hidden "DPPM Dashboard"
 * charts. Those are the legacy line-only charts. The monthly deck should use
 * the visible workbook charts on Plant Summary / MSC DPPM / CSC DPPM / ARU DPPM.
 *
 * This override also performs a complete HubSpot ticket sync before rebuilding
 * the service-ticket chart source. Countable issues come directly from the
 * Support Pipeline using Startup/Warranty/Service and excluding tickets where
 * MJC No Fault is set. The same counts feed DPPM Inputs and the Slides summary
 * tables so HubSpot, Sheets, and Slides all use the same method.
 *
 * This file is intentionally named ZZZ_* so clasp/Apps Script loads it after
 * MonthlyReportPackage.gs and this implementation wins for the shared function
 * name used by updateMonthlyQualityReport().
 */

const MONTHLY_QUALITY_VISIBLE_DPPM_CHART_LAYOUT_ = Object.freeze({
  left: 75.384,
  top: 36,
  width: 564.33,
  height: 211.05
});

const MONTHLY_QUALITY_SERVICE_CHART_LAYOUT_ = Object.freeze({
  left: 23.72,
  top: 14.2,
  width: 398.9,
  height: 212.35
});

const MONTHLY_QUALITY_VISIBLE_DPPM_SECTIONS_ = Object.freeze([
  {
    slideLabel: 'All Lines',
    sheetName: 'Plant Summary',
    chartTitle: 'All 12-Mth Rolling DPPM'
  },
  {
    slideLabel: 'MSC',
    sheetName: 'MSC DPPM',
    chartTitle: 'MSC 12-Mth Rolling DPPM'
  },
  {
    slideLabel: 'CSC',
    sheetName: 'CSC DPPM',
    chartTitle: 'CSC 12-Mth Rolling DPPM'
  },
  {
    slideLabel: 'ARU',
    sheetName: 'ARU DPPM',
    chartTitle: 'ARU 12-Mth Rolling DPPM'
  }
]);

/**
 * Refresh HubSpot, rebuild issue/DPPM source data, replace the DPPM chart
 * placeholders, update the current/previous issue tables, and refresh every
 * linked Sheets chart in the deck.
 */
function updateMonthlyQualityPackageDPPM_(packageResult) {
  const hubSpotSyncResult = syncHubSpotSupportTicketsToSpreadsheet_(
    packageResult.dataFile.getId()
  );
  Logger.log(JSON.stringify(hubSpotSyncResult));

  if (hubSpotSyncResult.status !== 'READY') {
    throw new Error(
      'HubSpot sync is required for monthly quality issue reporting: ' +
      String(hubSpotSyncResult.reason || hubSpotSyncResult.status)
    );
  }

  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());

  const validatedIssueChartResult =
    refreshMonthlyQualityValidatedIssueChartData_(spreadsheet);
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

  // Refresh every linked Sheets chart in the deck. This includes the recreated
  // service-ticket charts, so the freshly synchronized HubSpot data is rendered
  // into Slides during the same monthly-report run.
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

function findMonthlyQualityVisibleDPPMChart_(spreadsheet, section) {
  const sheet = spreadsheet.getSheetByName(section.sheetName);
  if (!sheet) {
    throw new Error('DPPM chart sheet not found: ' + section.sheetName);
  }

  const charts = sheet.getCharts();
  for (let index = 0; index < charts.length; index++) {
    const title = String(charts[index].getOptions().get('title') || '').trim();
    if (title === section.chartTitle) return charts[index];
  }

  throw new Error(
    'Visible DPPM chart not found: ' + section.sheetName + ' / ' + section.chartTitle
  );
}

function replaceMonthlyQualityVisibleDPPMChart_(slide, chart) {
  const elements = slide.getPageElements();
  let dppmPlaceholder = null;
  let largestPlaceholder = null;
  let largestArea = 0;

  elements.forEach(function(element) {
    const type = element.getPageElementType();
    if (
      type !== SlidesApp.PageElementType.IMAGE &&
      type !== SlidesApp.PageElementType.SHEETS_CHART
    ) {
      return;
    }

    let title = '';
    try {
      title = String(element.getTitle ? element.getTitle() : '').trim();
    } catch (error) {
      title = '';
    }

    if (title.indexOf('12-Mth Rolling DPPM') !== -1) {
      dppmPlaceholder = element;
    }

    const area = element.getWidth() * element.getHeight();
    if (area > largestArea) {
      largestArea = area;
      largestPlaceholder = element;
    }
  });

  const placeholder = dppmPlaceholder || largestPlaceholder;
  if (!placeholder) {
    throw new Error('No DPPM chart image or linked chart was found on the slide.');
  }

  placeholder.remove();
  slide.insertSheetsChart(
    chart,
    MONTHLY_QUALITY_VISIBLE_DPPM_CHART_LAYOUT_.left,
    MONTHLY_QUALITY_VISIBLE_DPPM_CHART_LAYOUT_.top,
    MONTHLY_QUALITY_VISIBLE_DPPM_CHART_LAYOUT_.width,
    MONTHLY_QUALITY_VISIBLE_DPPM_CHART_LAYOUT_.height
  );
}

function normalizeMonthlyQualityServiceChartSlides_(spreadsheet, presentation) {
  ['Coatings', 'Bard Coatings'].forEach(function(slideLabel) {
    const slide = findMonthlyQualitySlide_(presentation, slideLabel);
    removeMonthlyQualityNoServiceText_(slide);
  });

  const bardChart = ensureMonthlyQualityBardCoatingsServiceChart_(spreadsheet);
  const bardSlide = findMonthlyQualitySlide_(presentation, 'Bard Coatings');
  replaceMonthlyQualityServiceChartArea_(bardSlide, bardChart);
}

function ensureMonthlyQualityBardCoatingsServiceChart_(spreadsheet) {
  const chartTitle = 'Service Tickets by Month - Bard Coatings';
  const existingChart = findMonthlyQualityChartByTitle_(spreadsheet, chartTitle);
  if (existingChart) return existingChart;

  const dataSheet = spreadsheet.getSheetByName(VALIDATED_ISSUE_CHART_DATA_SHEET_);
  if (!dataSheet) {
    throw new Error('Missing HubSpot issue chart data sheet.');
  }

  let sheet = spreadsheet.getSheetByName('Bard Coatings');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Bard Coatings');
    sheet.hideGridlines(true);
  }

  sheet.getRange('A1').setValue('Quality Dashboard - Bard Coatings');
  sheet.getRange('A2').setValue(
    'HubSpot charts below are driven from the hidden HubSpot Chart Data tab.'
  );

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataSheet.getRange(1, 39, 13, 4))
    .setOption('title', chartTitle)
    .setOption('isStacked', true)
    .setOption('legend', { position: 'bottom' })
    .setOption('hAxis', { title: 'Create date - Monthly' })
    .setOption('vAxis', { title: 'Count of tickets' })
    .setPosition(4, 2, 0, 0)
    .build();
  sheet.insertChart(chart);
  SpreadsheetApp.flush();

  const insertedChart = findMonthlyQualityChartByTitle_(spreadsheet, chartTitle);
  if (!insertedChart) {
    throw new Error('Unable to create Bard Coatings service ticket chart.');
  }
  return insertedChart;
}

function findMonthlyQualityChartByTitle_(spreadsheet, chartTitle) {
  const sheets = spreadsheet.getSheets();
  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const charts = sheets[sheetIndex].getCharts();
    for (let chartIndex = 0; chartIndex < charts.length; chartIndex++) {
      const title = String(charts[chartIndex].getOptions().get('title') || '').trim();
      if (title === chartTitle) return charts[chartIndex];
    }
  }
  return null;
}

function removeMonthlyQualityNoServiceText_(slide) {
  slide.getPageElements().forEach(function(element) {
    if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

    let text = '';
    try {
      text = element.asShape().getText().asString();
    } catch (error) {
      text = '';
    }

    if (text.indexOf('This means no service tickets') !== -1) {
      element.remove();
    }
  });
}

function replaceMonthlyQualityServiceChartArea_(slide, chart) {
  slide.getPageElements().forEach(function(element) {
    const type = element.getPageElementType();
    if (
      type !== SlidesApp.PageElementType.IMAGE &&
      type !== SlidesApp.PageElementType.SHEETS_CHART
    ) {
      return;
    }

    if (element.getLeft() < 430 && element.getTop() < 260) {
      element.remove();
    }
  });

  slide.insertSheetsChart(
    chart,
    MONTHLY_QUALITY_SERVICE_CHART_LAYOUT_.left,
    MONTHLY_QUALITY_SERVICE_CHART_LAYOUT_.top,
    MONTHLY_QUALITY_SERVICE_CHART_LAYOUT_.width,
    MONTHLY_QUALITY_SERVICE_CHART_LAYOUT_.height
  );
}

/**
 * Keep the two small current/previous month issue tables on every quality slide
 * synchronized to the exact same HubSpot-derived counts used by the charts.
 */
function updateMonthlyQualityIssueSummaryTables_(presentation, issueResult) {
  if (!issueResult || !issueResult.summaries) return;

  const slideLabels = [
    'All Lines',
    'MSC',
    'ARU',
    'CSC',
    'Mods',
    'Gas Heat',
    'Coatings',
    'Bard Coatings'
  ];

  slideLabels.forEach(function(slideLabel) {
    const summary = issueResult.summaries[slideLabel];
    if (!summary) return;

    const slide = findMonthlyQualitySlide_(presentation, slideLabel);
    const tables = slide.getTables().filter(function(table) {
      return table.getNumRows() === 5 && table.getNumColumns() === 2;
    });

    if (tables.length < 2) {
      Logger.log(
        'Expected two 5x2 issue summary tables on slide ' + slideLabel +
        '; found ' + tables.length + '.'
      );
      return;
    }

    tables.sort(function(a, b) {
      return a.getLeft() - b.getLeft();
    });

    const displayProduct = slideLabel === 'All Lines' ? 'All' : slideLabel;
    writeMonthlyQualityIssueSummaryTable_(
      tables[0],
      issueResult.currentMonthLabel,
      displayProduct,
      summary.current
    );
    writeMonthlyQualityIssueSummaryTable_(
      tables[1],
      issueResult.previousMonthLabel,
      displayProduct,
      summary.previous
    );
  });
}

function writeMonthlyQualityIssueSummaryTable_(
  table,
  monthLabel,
  productLabel,
  counts
) {
  table.getCell(0, 0).getText().setText(String(monthLabel || ''));
  table.getCell(0, 1).getText().setText(String(productLabel || ''));

  const rows = [
    ['Total', counts.total],
    ['Startup', counts.startup],
    ['Warranty', counts.warranty],
    ['Service', counts.service]
  ];

  rows.forEach(function(row, index) {
    table.getCell(index + 1, 0).getText().setText(row[0]);
    table.getCell(index + 1, 1).getText().setText(String(row[1] || 0));
  });
}

function refreshMonthlyQualityLinkedSheetsCharts_(presentation) {
  let refreshed = 0;

  presentation.getSlides().forEach(function(slide) {
    slide.getPageElements().forEach(function(element) {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) {
        return;
      }

      try {
        element.asSheetsChart().refresh();
        refreshed++;
      } catch (error) {
        Logger.log(
          'Unable to refresh linked Sheets chart ' + element.getObjectId() +
          ': ' + error
        );
      }
    });
  });

  return refreshed;
}
