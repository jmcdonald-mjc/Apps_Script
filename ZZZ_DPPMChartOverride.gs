/**
 * Override for monthly quality DPPM slide chart replacement.
 *
 * The original monthly package routine targeted the hidden "DPPM Dashboard"
 * charts. Those are the legacy line-only charts. The monthly deck should use
 * the visible workbook charts on Plant Summary / MSC DPPM / CSC DPPM / ARU DPPM.
 *
 * This override also refreshes the service-ticket chart source from the
 * validated DPPM issue model before any linked Sheets charts are refreshed.
 * Raw HubSpot ticket counts must never overwrite these reviewed issue counts.
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
 * Before touching Slides, rebuild the service-ticket chart data from DPPM
 * Monthly so the deck always uses the reviewed quality issue counts rather
 * than the raw number of HubSpot tickets.
 */
function updateMonthlyQualityPackageDPPM_(packageResult) {
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
    replaceMonthlyQualityPackageChart_(slide, chart);
    updatedSections.push(section.slideLabel);
  });

  // Refresh every linked Sheets chart in the deck. This includes the recreated
  // service-ticket charts, so changes made to HubSpot Chart Data are rendered
  // into Slides during the same monthly-report run.
  const linkedChartsRefreshed = refreshMonthlyQualityLinkedSheetsCharts_(
    presentation
  );

  presentation.saveAndClose();
  return {
    updatedSections: updatedSections,
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
