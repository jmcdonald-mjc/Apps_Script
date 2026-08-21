/**
 * Monthly Quality report package automation.
 *
 * Creates one month folder containing the Google Slides report and its linked
 * DPPM Google Sheet. Uses the existing SlidesAutomation.gs configuration and
 * FPY helpers. The DPPM template is discovered by exact name in the configured
 * report root so no additional Drive identifier is stored in source control.
 */

const MONTHLY_QUALITY_PACKAGE_DPPM_TEMPLATE_NAME_ =
  'Monthly Quality DPPM Data - TEMPLATE';
const MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_ = 'DPPM Config';
const MONTHLY_QUALITY_PACKAGE_DPPM_MONTHLY_SHEET_ = 'DPPM Monthly';
const MONTHLY_QUALITY_PACKAGE_DPPM_DASHBOARD_SHEET_ = 'DPPM Dashboard';

const MONTHLY_QUALITY_PACKAGE_DPPM_SECTIONS_ = Object.freeze([
  { slideLabel: 'All Lines', chartTitle: 'All Lines 12-Month Rolling DPPM' },
  { slideLabel: 'MSC', chartTitle: 'MSC 12-Month Rolling DPPM' },
  { slideLabel: 'CSC', chartTitle: 'CSC 12-Month Rolling DPPM' },
  { slideLabel: 'ARU', chartTitle: 'ARU 12-Month Rolling DPPM' }
]);

/**
 * Read-only setup check for the new monthly folder + sheet + deck workflow.
 */
function validateMonthlyQualityReportPackageAutomation() {
  const config = getMonthlyQualityAutomationConfig_();
  const rootFolder = DriveApp.getFolderById(config.outputFolderId);
  const dppmTemplate = findMonthlyQualityDPPMTemplate_(rootFolder);
  const spreadsheet = SpreadsheetApp.openById(dppmTemplate.getId());
  const configSheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_
  );
  const dashboard = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_DASHBOARD_SHEET_
  );
  if (!configSheet || !dashboard) {
    throw new Error(
      'The DPPM template must contain "DPPM Config" and "DPPM Dashboard".'
    );
  }
  if (dashboard.getCharts().length < MONTHLY_QUALITY_PACKAGE_DPPM_SECTIONS_.length) {
    throw new Error('The DPPM template does not contain all four dashboard charts.');
  }

  const result = {
    reportsFolderName: rootFolder.getName(),
    dppmTemplateName: dppmTemplate.getName(),
    dppmChartCount: dashboard.getCharts().length,
    timeZone: config.timeZone
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * Creates or reuses the prior completed month's subfolder, deck, and data
 * spreadsheet. Repeated runs return the same three Drive objects.
 */
function createOrGetMonthlyQualityReportPackage(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const context = getMonthlyQualityPackageContext_(asOfDate);
    const packageResult = ensureMonthlyQualityReportPackageUnlocked_(context);
    const result = describeMonthlyQualityReportPackage_(packageResult, context);
    Logger.log(JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Primary monthly runner. Creates/reuses the monthly package, updates FPY
 * tables, and replaces the old DPPM screenshots with linked Sheets charts.
 */
function updateMonthlyQualityReport(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const context = getMonthlyQualityPackageContext_(asOfDate);
    const packageResult = ensureMonthlyQualityReportPackageUnlocked_(context);
    const fpyResult = updateMonthlyQualityPackageFPY_(context, packageResult);
    const dppmResult = updateMonthlyQualityPackageDPPM_(packageResult);
    const result = describeMonthlyQualityReportPackage_(packageResult, context);
    result.fpySections = fpyResult.updatedSections;
    result.plantAverage = fpyResult.plantAverage;
    result.dppmSections = dppmResult.updatedSections;
    Logger.log(JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates only the four linked DPPM charts in the monthly deck.
 */
function updateMonthlyQualityDPPMSlides(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const context = getMonthlyQualityPackageContext_(asOfDate);
    const packageResult = ensureMonthlyQualityReportPackageUnlocked_(context);
    const updateResult = updateMonthlyQualityPackageDPPM_(packageResult);
    const result = describeMonthlyQualityReportPackage_(packageResult, context);
    result.updatedSections = updateResult.updatedSections;
    Logger.log(JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function getMonthlyQualityPackageContext_(asOfDate) {
  const context = getPreviousCompletedQualityMonth_(asOfDate);
  context.folderName = 'Monthly Quality Status - ' + context.monthLabel;
  context.dataName = 'Monthly Quality Data - ' + context.monthLabel;
  return context;
}

function ensureMonthlyQualityReportPackageUnlocked_(context) {
  const config = getMonthlyQualityAutomationConfig_();
  const rootFolder = DriveApp.getFolderById(config.outputFolderId);
  const dppmTemplate = findMonthlyQualityDPPMTemplate_(rootFolder);
  const folderResult = ensureMonthlyQualityPackageFolder_(rootFolder, context);
  const deckResult = ensureMonthlyQualityPackageFile_(
    folderResult.folder,
    rootFolder,
    context.deckName,
    MimeType.GOOGLE_SLIDES,
    DriveApp.getFileById(config.templateId),
    monthlyQualityPackagePropertyKey_('DECK', context.monthKey)
  );
  const dataResult = ensureMonthlyQualityPackageFile_(
    folderResult.folder,
    rootFolder,
    context.dataName,
    MimeType.GOOGLE_SHEETS,
    dppmTemplate,
    monthlyQualityPackagePropertyKey_('DATA', context.monthKey)
  );

  configureMonthlyQualityPackageData_(dataResult.file, context, dataResult.created);
  PropertiesService.getScriptProperties().setProperty(
    monthlyQualityPackagePropertyKey_('FOLDER', context.monthKey),
    folderResult.folder.getId()
  );

  return {
    folder: folderResult.folder,
    deckFile: deckResult.file,
    dataFile: dataResult.file,
    createdFolder: folderResult.created,
    createdDeck: deckResult.created,
    createdData: dataResult.created
  };
}

function findMonthlyQualityDPPMTemplate_(rootFolder) {
  const files = collectMonthlyQualityPackageFiles_(
    rootFolder,
    MONTHLY_QUALITY_PACKAGE_DPPM_TEMPLATE_NAME_,
    MimeType.GOOGLE_SHEETS
  );
  if (files.length !== 1) {
    throw new Error(
      'Expected exactly one native Google Sheet named "' +
      MONTHLY_QUALITY_PACKAGE_DPPM_TEMPLATE_NAME_ +
      '" in the configured report folder; found ' + files.length + '.'
    );
  }
  return files[0];
}

function ensureMonthlyQualityPackageFolder_(rootFolder, context) {
  const matches = rootFolder.getFoldersByName(context.folderName);
  const folders = [];
  while (matches.hasNext()) folders.push(matches.next());
  if (folders.length > 1) {
    throw new Error(
      'Multiple report folders named "' + context.folderName + '" exist.'
    );
  }
  if (folders.length === 1) return { folder: folders[0], created: false };
  return { folder: rootFolder.createFolder(context.folderName), created: true };
}

function ensureMonthlyQualityPackageFile_(
  monthFolder,
  rootFolder,
  fileName,
  mimeType,
  templateFile,
  propertyKey
) {
  const properties = PropertiesService.getScriptProperties();
  const storedId = String(properties.getProperty(propertyKey) || '').trim();
  if (storedId) {
    try {
      const storedFile = DriveApp.getFileById(storedId);
      if (storedFile.getName() === fileName && storedFile.getMimeType() === mimeType) {
        if (!monthlyQualityPackageFileIsInFolder_(storedFile, monthFolder.getId())) {
          storedFile.moveTo(monthFolder);
        }
        return { file: storedFile, created: false };
      }
    } catch (error) {
      // Recover from a stale Script Property by searching by exact name.
    }
    properties.deleteProperty(propertyKey);
  }

  const inMonth = collectMonthlyQualityPackageFiles_(monthFolder, fileName, mimeType);
  if (inMonth.length > 1) {
    throw new Error('Multiple files named "' + fileName + '" exist in the month folder.');
  }
  if (inMonth.length === 1) {
    properties.setProperty(propertyKey, inMonth[0].getId());
    return { file: inMonth[0], created: false };
  }

  const inRoot = collectMonthlyQualityPackageFiles_(rootFolder, fileName, mimeType);
  if (inRoot.length > 1) {
    throw new Error('Multiple files named "' + fileName + '" exist in the report root.');
  }
  if (inRoot.length === 1) {
    inRoot[0].moveTo(monthFolder);
    properties.setProperty(propertyKey, inRoot[0].getId());
    return { file: inRoot[0], created: false };
  }

  const newFile = templateFile.makeCopy(fileName, monthFolder);
  newFile.setDescription('Monthly Quality report artifact created automatically.');
  properties.setProperty(propertyKey, newFile.getId());
  return { file: newFile, created: true };
}

function collectMonthlyQualityPackageFiles_(folder, fileName, mimeType) {
  const matches = folder.getFilesByName(fileName);
  const files = [];
  while (matches.hasNext()) {
    const file = matches.next();
    if (file.getMimeType() === mimeType) files.push(file);
  }
  return files;
}

function monthlyQualityPackageFileIsInFolder_(file, folderId) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return true;
  }
  return false;
}

function monthlyQualityPackagePropertyKey_(type, monthKey) {
  return 'QUALITY_REPORT_PACKAGE_' + type + '_' +
    String(monthKey).replace('-', '_');
}

function configureMonthlyQualityPackageData_(dataFile, context, isNewCopy) {
  const spreadsheet = SpreadsheetApp.openById(dataFile.getId());
  const configSheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_
  );
  if (!configSheet) throw new Error('Monthly data sheet is missing "DPPM Config".');
  configSheet.getRange('B7').setValue(new Date(context.year, context.month - 1, 1));
  configSheet.getRange('B7').setNumberFormat('mmm yyyy');
  if (isNewCopy) {
    configSheet.getRange('B12').setValue('Epicor and HubSpot credentials pending');
  }
  ensureMonthlyQualityPackageDPPMModel_(spreadsheet, context);
  SpreadsheetApp.flush();
}

function ensureMonthlyQualityPackageDPPMModel_(spreadsheet, context) {
  const sheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_MONTHLY_SHEET_
  );
  if (!sheet) throw new Error('Monthly data sheet is missing "DPPM Monthly".');

  const productLines = ['MSC', 'ARU', 'CSC', 'All Lines'];
  let lastRow = sheet.getLastRow();
  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 2).getValues()
    : [];
  const present = {};
  let latestDate = null;

  existing.forEach(function(row) {
    const date = row[0];
    const productLine = String(row[1] || '').trim();
    if (!(date instanceof Date) || isNaN(date.getTime()) || !productLine) return;
    const key = Utilities.formatDate(date, 'UTC', 'yyyy-MM') + '|' + productLine;
    present[key] = true;
    if (!latestDate || date.getTime() > latestDate.getTime()) latestDate = date;
  });

  if (!latestDate) latestDate = new Date(Date.UTC(context.year, context.month - 1, 1));
  let cursor = new Date(Date.UTC(
    latestDate.getUTCFullYear(), latestDate.getUTCMonth() + 1, 1
  ));
  const reportDate = new Date(Date.UTC(context.year, context.month - 1, 1));
  const rowsToAppend = [];

  while (cursor.getTime() <= reportDate.getTime()) {
    productLines.forEach(function(productLine) {
      const monthKey = Utilities.formatDate(cursor, 'UTC', 'yyyy-MM');
      if (!present[monthKey + '|' + productLine]) {
        rowsToAppend.push([new Date(cursor.getTime()), productLine]);
      }
    });
    cursor = new Date(Date.UTC(cursor.getUTCFullYear(), cursor.getUTCMonth() + 1, 1));
  }

  if (rowsToAppend.length) {
    const startRow = lastRow + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
    sheet.getRange(startRow, 1, rowsToAppend.length, 1).setNumberFormat('mmm yyyy');
    lastRow += rowsToAppend.length;
  }

  const modelRows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const formulas = modelRows.map(function(row, index) {
    return buildMonthlyQualityPackageDPPMFormulaRow_(
      index + 2,
      String(row[1] || '').trim()
    );
  });
  sheet.getRange(2, 3, formulas.length, 10).setFormulas(formulas);
  normalizeMonthlyQualityPackageDashboardFormulas_(spreadsheet);
}

function buildMonthlyQualityPackageDPPMFormulaRow_(rowNumber, productLine) {
  const row = String(rowNumber);
  let inputFormulas;
  if (productLine === 'All Lines') {
    const firstProductRow = rowNumber - 3;
    const lastProductRow = rowNumber - 1;
    inputFormulas = ['C', 'D', 'E', 'F'].map(function(column) {
      return '=IF(COUNT(' + column + firstProductRow + ':' + column + lastProductRow +
        ')=0,"",SUM(' + column + firstProductRow + ':' + column + lastProductRow + '))';
    });
  } else {
    inputFormulas = ['C', 'D', 'E', 'F'].map(function(column) {
      return '=IF(COUNTIFS(\'DPPM Inputs\'!$A$2:$A,$A' + row +
        ',\'DPPM Inputs\'!$B$2:$B,$B' + row +
        ',\'DPPM Inputs\'!$' + column + '$2:$' + column + ',"<>")=0,"",' +
        'SUMIFS(\'DPPM Inputs\'!$' + column + '$2:$' + column +
        ',\'DPPM Inputs\'!$A$2:$A,$A' + row +
        ',\'DPPM Inputs\'!$B$2:$B,$B' + row + '))';
    });
  }

  return inputFormulas.concat([
    '=IF(COUNT(D' + row + ':E' + row + ')=0,"",SUM(D' + row + ':E' + row + '))',
    '=IF(COUNT(D' + row + ':F' + row + ')=0,"",SUM(D' + row + ':F' + row + '))',
    '=IF(OR(C' + row + '="",C' + row + '=0,G' + row + '=""),"",G' + row + '/C' + row + '*1000000)',
    '=IF(C' + row + '="","",IFERROR(SUMIFS($G$2:$G,$B$2:$B,B' + row +
      ',$A$2:$A,">"&EDATE(A' + row + ',-12),$A$2:$A,"<="&A' + row +
      ')/SUMIFS($C$2:$C,$B$2:$B,B' + row + ',$A$2:$A,">"&EDATE(A' +
      row + ',-12),$A$2:$A,"<="&A' + row + ')*1000000,""))',
    '=IFERROR(VLOOKUP(B' + row + ',\'DPPM Config\'!$A$2:$B$5,2,FALSE),"")',
    '=IF(C' + row + '="","Pending API data","Ready")'
  ]);
}

function normalizeMonthlyQualityPackageDashboardFormulas_(spreadsheet) {
  const dashboard = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_DASHBOARD_SHEET_
  );
  if (!dashboard) throw new Error('Monthly data sheet is missing "DPPM Dashboard".');
  ['A4:D15', 'A21:D32', 'A38:D49', 'A55:D66'].forEach(function(a1Notation) {
    const range = dashboard.getRange(a1Notation);
    const formulas = range.getFormulas().map(function(row) {
      return row.map(function(formula) {
        return formula.replace(/:\$([A-Z]+)\$317/g, function(match, column) {
          return ':$' + column;
        });
      });
    });
    range.setFormulas(formulas);
  });
}

function updateMonthlyQualityPackageFPY_(context, packageResult) {
  const fpyData = buildMonthlyQualityFPYData_(context);
  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const updatedSections = [];
  MONTHLY_QUALITY_FPY_DECK_SECTIONS_.forEach(function(section) {
    const slide = findMonthlyQualitySlide_(presentation, section.slideLabel);
    const table = findMonthlyQualityFPYTable_(slide);
    writeMonthlyQualityFPYTable_(table, section, fpyData);
    updatedSections.push(section.slideLabel);
  });
  presentation.saveAndClose();
  return {
    updatedSections: updatedSections,
    plantAverage: formatMonthlyQualityPercent_(fpyData.plantAverage)
  };
}

function updateMonthlyQualityPackageDPPM_(packageResult) {
  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());
  const dashboard = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_DASHBOARD_SHEET_
  );
  if (!dashboard) throw new Error('Monthly data sheet is missing "DPPM Dashboard".');
  SpreadsheetApp.flush();
  const charts = dashboard.getCharts();
  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const updatedSections = [];

  MONTHLY_QUALITY_PACKAGE_DPPM_SECTIONS_.forEach(function(section) {
    const chart = findMonthlyQualityPackageChart_(charts, section.chartTitle);
    const slide = findMonthlyQualitySlide_(presentation, section.slideLabel);
    replaceMonthlyQualityPackageChart_(slide, chart);
    updatedSections.push(section.slideLabel);
  });
  presentation.saveAndClose();
  return { updatedSections: updatedSections };
}

function findMonthlyQualityPackageChart_(charts, chartTitle) {
  for (let index = 0; index < charts.length; index++) {
    const title = String(charts[index].getOptions().get('title') || '').trim();
    if (title === chartTitle) return charts[index];
  }
  throw new Error('DPPM dashboard chart not found: ' + chartTitle);
}

function replaceMonthlyQualityPackageChart_(slide, chart) {
  const elements = slide.getPageElements();
  let placeholder = null;
  let largestArea = 0;
  elements.forEach(function(element) {
    const type = element.getPageElementType();
    if (
      type !== SlidesApp.PageElementType.IMAGE &&
      type !== SlidesApp.PageElementType.SHEETS_CHART
    ) return;
    const area = element.getWidth() * element.getHeight();
    if (area > largestArea) {
      largestArea = area;
      placeholder = element;
    }
  });
  if (!placeholder) {
    throw new Error('No DPPM chart image or linked chart was found on the slide.');
  }

  const left = placeholder.getLeft();
  const top = placeholder.getTop();
  const width = placeholder.getWidth();
  const height = placeholder.getHeight();
  placeholder.remove();
  slide.insertSheetsChart(chart, left, top, width, height);
}

function describeMonthlyQualityReportPackage_(packageResult, context) {
  return {
    reportMonth: context.monthLabel,
    folderId: packageResult.folder.getId(),
    folderName: packageResult.folder.getName(),
    folderUrl: packageResult.folder.getUrl(),
    deckId: packageResult.deckFile.getId(),
    deckName: packageResult.deckFile.getName(),
    deckUrl: packageResult.deckFile.getUrl(),
    dataSpreadsheetId: packageResult.dataFile.getId(),
    dataSpreadsheetName: packageResult.dataFile.getName(),
    dataSpreadsheetUrl: packageResult.dataFile.getUrl(),
    createdFolder: Boolean(packageResult.createdFolder),
    createdDeck: Boolean(packageResult.createdDeck),
    createdDataSpreadsheet: Boolean(packageResult.createdData)
  };
}
