/**
 * Override for monthly quality DPPM slide chart replacement.
 *
 * The original monthly package routine targeted the hidden "DPPM Dashboard"
 * charts. Those are the legacy line-only charts. The monthly deck should use
 * the visible workbook charts on Plant Summary / MSC DPPM / CSC DPPM / ARU DPPM.
 *
 * This file is intentionally named ZZZ_* so clasp/Apps Script loads it after
 * MonthlyReportPackage.gs and this implementation wins for the shared function
 * name used by updateMonthlyQualityReport().
 */

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
 * Replaces the DPPM chart placeholders with the current visible workbook charts.
 * This intentionally does not use the hidden DPPM Dashboard tab.
 */
function updateMonthlyQualityPackageDPPM_(packageResult) {
  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());
  SpreadsheetApp.flush();

  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const updatedSections = [];

  MONTHLY_QUALITY_VISIBLE_DPPM_SECTIONS_.forEach(function(section) {
    const chart = findMonthlyQualityVisibleDPPMChart_(spreadsheet, section);
    const slide = findMonthlyQualitySlide_(presentation, section.slideLabel);
    replaceMonthlyQualityPackageChart_(slide, chart);
    updatedSections.push(section.slideLabel);
  });

  presentation.saveAndClose();
  return { updatedSections: updatedSections };
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
