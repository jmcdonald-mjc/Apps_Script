/**
 * Monthly Quality Status Google Slides automation.
 *
 * Owns configuration, monthly deck creation, and FPY table updates.
 */

const MONTHLY_QUALITY_PROPERTY_KEYS_ = Object.freeze({
  TEMPLATE_ID: 'QUALITY_SLIDES_TEMPLATE_ID',
  OUTPUT_FOLDER_ID: 'QUALITY_REPORTS_FOLDER_ID',
  DATA_SPREADSHEET_ID: 'QUALITY_DATA_SPREADSHEET_ID',
  TIME_ZONE: 'QUALITY_REPORT_TIME_ZONE'
});

const MONTHLY_QUALITY_INITIAL_CONFIGURATION_ = Object.freeze({
  templateId: '1KOmS8kIPn9fGzcIhk5l0Xr-tNLjX1-jc-V0GGnIoJ7c',
  outputFolderId: '1vBoyVpqKYi102sStqLAaskQfKHKcDvO-',
  dataSpreadsheetId: '1XVR755InumWMzcMehoxjw9EYuliDAb5rCKl_kE321XM',
  timeZone: 'America/New_York'
});

const MONTHLY_QUALITY_FPY_SHEET_NAME_ = 'Sheet1';

const MONTHLY_QUALITY_FPY_DECK_SECTIONS_ = Object.freeze([
  { slideLabel: 'All Lines', tableLabel: 'All', sheetHeader: 'All Lines' },
  { slideLabel: 'MSC', tableLabel: 'MSC', sheetHeader: 'MSC' },
  { slideLabel: 'CSC', tableLabel: 'CSC', sheetHeader: 'CSC' },
  { slideLabel: 'ARU', tableLabel: 'ARU', sheetHeader: 'ARU' },
  { slideLabel: 'Mods', tableLabel: 'Mods', sheetHeader: 'HGRH' },
  { slideLabel: 'Coatings', tableLabel: 'Coatings', sheetHeader: 'Coatings' },
  {
    slideLabel: 'Bard Coatings',
    tableLabel: 'Bard Coatings',
    sheetHeader: 'Bard Coatings'
  },
  { slideLabel: 'Gas Heat', tableLabel: 'Gas Heat', sheetHeader: 'Gas Heat' }
]);

/**
 * One-time setup function. Run this after the file is pushed to Apps Script.
 * The stored values are Drive resource identifiers, not credentials.
 */
function configureMonthlyQualityAutomation() {
  const keys = MONTHLY_QUALITY_PROPERTY_KEYS_;

  PropertiesService.getScriptProperties().setProperties({
    [keys.TEMPLATE_ID]: MONTHLY_QUALITY_INITIAL_CONFIGURATION_.templateId,
    [keys.OUTPUT_FOLDER_ID]: MONTHLY_QUALITY_INITIAL_CONFIGURATION_.outputFolderId,
    [keys.DATA_SPREADSHEET_ID]:
      MONTHLY_QUALITY_INITIAL_CONFIGURATION_.dataSpreadsheetId,
    [keys.TIME_ZONE]: MONTHLY_QUALITY_INITIAL_CONFIGURATION_.timeZone
  }, false);

  return validateMonthlyQualityAutomationConfiguration();
}

/**
 * Confirms that all required configuration exists and that the configured
 * template and output folder can be opened by the account running the script.
 * This function is read-only with respect to Drive and Slides.
 */
function validateMonthlyQualityAutomationConfiguration() {
  const config = getMonthlyQualityAutomationConfig_();
  const templateFile = DriveApp.getFileById(config.templateId);
  const outputFolder = DriveApp.getFolderById(config.outputFolderId);
  const dataSpreadsheet = SpreadsheetApp.openById(config.dataSpreadsheetId);
  const dataSheet = dataSpreadsheet.getSheetByName(MONTHLY_QUALITY_FPY_SHEET_NAME_);

  if (templateFile.getMimeType() !== MimeType.GOOGLE_SLIDES) {
    throw new Error(
      'QUALITY_SLIDES_TEMPLATE_ID must identify a native Google Slides presentation.'
    );
  }

  const presentation = SlidesApp.openById(config.templateId);
  if (!dataSheet) {
    throw new Error(
      'The configured spreadsheet does not contain the required "' +
      MONTHLY_QUALITY_FPY_SHEET_NAME_ +
      '" sheet.'
    );
  }
  const result = {
    templateId: config.templateId,
    templateName: templateFile.getName(),
    templateSlideCount: presentation.getSlides().length,
    outputFolderId: config.outputFolderId,
    outputFolderName: outputFolder.getName(),
    dataSpreadsheetId: config.dataSpreadsheetId,
    dataSpreadsheetName: dataSpreadsheet.getName(),
    dataSheetName: dataSheet.getName(),
    timeZone: config.timeZone
  };

  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * Returns validated configuration for internal automation functions.
 */
function getMonthlyQualityAutomationConfig_() {
  const keys = MONTHLY_QUALITY_PROPERTY_KEYS_;
  const properties = PropertiesService.getScriptProperties();
  const config = {
    templateId: String(properties.getProperty(keys.TEMPLATE_ID) || '').trim(),
    outputFolderId: String(properties.getProperty(keys.OUTPUT_FOLDER_ID) || '').trim(),
    dataSpreadsheetId: String(
      properties.getProperty(keys.DATA_SPREADSHEET_ID) || ''
    ).trim(),
    timeZone: String(
      properties.getProperty(keys.TIME_ZONE) || Session.getScriptTimeZone() || ''
    ).trim()
  };

  const missing = [];
  if (!config.templateId) missing.push(keys.TEMPLATE_ID);
  if (!config.outputFolderId) missing.push(keys.OUTPUT_FOLDER_ID);
  if (!config.dataSpreadsheetId) missing.push(keys.DATA_SPREADSHEET_ID);
  if (!config.timeZone) missing.push(keys.TIME_ZONE);

  if (missing.length) {
    throw new Error(
      'Monthly Quality automation is not configured. Missing Script Properties: ' +
      missing.join(', ') +
      '. Run configureMonthlyQualityAutomation() once.'
    );
  }

  return config;
}

/**
 * Produces the stable naming context for the previous completed month.
 * Example: an August 2026 run returns July 2026 and key 2026-07.
 */
function getPreviousCompletedQualityMonth_(asOfDate) {
  const config = getMonthlyQualityAutomationConfig_();
  const parts = Utilities.formatDate(
    asOfDate || new Date(),
    config.timeZone,
    'yyyy-MM'
  ).split('-');
  const currentYear = Number(parts[0]);
  const currentMonth = Number(parts[1]);
  const reportYear = currentMonth === 1 ? currentYear - 1 : currentYear;
  const reportMonth = currentMonth === 1 ? 12 : currentMonth - 1;
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const monthKey = reportYear + '-' + String(reportMonth).padStart(2, '0');
  const monthLabel = monthNames[reportMonth - 1] + ' ' + reportYear;

  return {
    year: reportYear,
    month: reportMonth,
    monthKey: monthKey,
    monthLabel: monthLabel,
    deckName: 'Monthly Quality Status - ' + monthLabel
  };
}

/**
 * Creates the previous completed month's deck if it does not already exist.
 * Repeated runs return the same deck instead of making duplicates.
 */
function createOrGetMonthlyQualityDeck(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const context = getPreviousCompletedQualityMonth_(asOfDate);
    const file = ensureMonthlyQualityDeckUnlocked_(context);
    return describeMonthlyQualityDeck_(file, context, false);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Main entry point for this batch. Creates/reuses the monthly deck and updates
 * all native FPY summary tables using data through the report month.
 */
function updateMonthlyQualityFPYSlides(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const context = getPreviousCompletedQualityMonth_(asOfDate);
    const fileResult = ensureMonthlyQualityDeckUnlocked_(context, true);
    const file = fileResult.file;
    const fpyData = buildMonthlyQualityFPYData_(context);
    const presentation = SlidesApp.openById(file.getId());
    const updatedSections = [];

    MONTHLY_QUALITY_FPY_DECK_SECTIONS_.forEach(function(section) {
      const slide = findMonthlyQualitySlide_(presentation, section.slideLabel);
      const table = findMonthlyQualityFPYTable_(slide);
      writeMonthlyQualityFPYTable_(table, section, fpyData);
      updatedSections.push(section.slideLabel);
    });

    presentation.saveAndClose();

    const result = describeMonthlyQualityDeck_(
      file,
      context,
      fileResult.created
    );
    result.updatedSections = updatedSections;
    result.plantAverage = formatMonthlyQualityPercent_(fpyData.plantAverage);
    Logger.log(JSON.stringify(result));
    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Read-only data preview for checking the report cutoff before editing a deck.
 */
function previewMonthlyQualityFPYData(asOfDate) {
  const context = getPreviousCompletedQualityMonth_(asOfDate);
  const data = buildMonthlyQualityFPYData_(context);
  const result = {
    reportMonth: context.monthLabel,
    deckName: context.deckName,
    monthKeys: data.months.map(function(month) { return month.monthKey; }),
    plantAverage: formatMonthlyQualityPercent_(data.plantAverage),
    sections: {}
  };

  MONTHLY_QUALITY_FPY_DECK_SECTIONS_.forEach(function(section) {
    const sectionData = data.sections[section.sheetHeader];
    result.sections[section.tableLabel] = {
      currentYearTotal: formatMonthlyQualityFPYRecord_(sectionData.ytd),
      months: sectionData.months.map(formatMonthlyQualityFPYRecord_)
    };
  });

  Logger.log(JSON.stringify(result));
  return result;
}

function ensureMonthlyQualityDeckUnlocked_(context, includeCreationStatus) {
  const config = getMonthlyQualityAutomationConfig_();
  const properties = PropertiesService.getScriptProperties();
  const propertyKey = getMonthlyQualityDeckPropertyKey_(context.monthKey);
  const storedId = String(properties.getProperty(propertyKey) || '').trim();

  if (storedId) {
    try {
      const storedFile = DriveApp.getFileById(storedId);
      if (
        storedFile.getName() === context.deckName &&
        storedFile.getMimeType() === MimeType.GOOGLE_SLIDES &&
        fileIsInMonthlyQualityFolder_(storedFile, config.outputFolderId)
      ) {
        return includeCreationStatus
          ? { file: storedFile, created: false }
          : storedFile;
      }
    } catch (error) {
      // The saved reference is stale; recover by checking the output folder.
    }
    properties.deleteProperty(propertyKey);
  }

  const outputFolder = DriveApp.getFolderById(config.outputFolderId);
  const matches = outputFolder.getFilesByName(context.deckName);
  const matchingSlides = [];
  while (matches.hasNext()) {
    const candidate = matches.next();
    if (candidate.getMimeType() === MimeType.GOOGLE_SLIDES) {
      matchingSlides.push(candidate);
    }
  }

  if (matchingSlides.length > 1) {
    throw new Error(
      'Multiple monthly decks named "' + context.deckName +
      '" already exist in the output folder. Resolve the duplicates before rerunning.'
    );
  }

  if (matchingSlides.length === 1) {
    properties.setProperty(propertyKey, matchingSlides[0].getId());
    return includeCreationStatus
      ? { file: matchingSlides[0], created: false }
      : matchingSlides[0];
  }

  const templateFile = DriveApp.getFileById(config.templateId);
  const newFile = templateFile.makeCopy(context.deckName, outputFolder);
  newFile.setDescription(
    'Monthly Quality Status report for ' + context.monthLabel +
    '. Created automatically from template ' + config.templateId + '.'
  );
  properties.setProperty(propertyKey, newFile.getId());

  return includeCreationStatus
    ? { file: newFile, created: true }
    : newFile;
}

function getMonthlyQualityDeckPropertyKey_(monthKey) {
  return 'QUALITY_REPORT_DECK_' + String(monthKey).replace('-', '_');
}

function fileIsInMonthlyQualityFolder_(file, folderId) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return true;
  }
  return false;
}

function describeMonthlyQualityDeck_(file, context, created) {
  return {
    reportMonth: context.monthLabel,
    deckId: file.getId(),
    deckName: file.getName(),
    deckUrl: file.getUrl(),
    created: Boolean(created)
  };
}

function buildMonthlyQualityFPYData_(context) {
  const config = getMonthlyQualityAutomationConfig_();
  const spreadsheet = SpreadsheetApp.openById(config.dataSpreadsheetId);
  const sheet = spreadsheet.getSheetByName(MONTHLY_QUALITY_FPY_SHEET_NAME_);
  if (!sheet) {
    throw new Error(
      'Required FPY sheet not found: ' + MONTHLY_QUALITY_FPY_SHEET_NAME_
    );
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 25);
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  if (values.length < 6) {
    throw new Error('The FPY summary sheet does not contain the expected data rows.');
  }

  const header = values[0];
  const startColumns = {};
  MONTHLY_QUALITY_FPY_DECK_SECTIONS_.forEach(function(section) {
    const index = header.findIndex(function(value) {
      return String(value || '').trim() === section.sheetHeader;
    });
    if (index < 0) {
      throw new Error(
        'FPY summary header not found for product line: ' + section.sheetHeader
      );
    }
    startColumns[section.sheetHeader] = index;
  });

  const monthlyRows = {};
  values.slice(5).forEach(function(row) {
    const monthKey = getMonthlyQualitySheetMonthKey_(row[0], config.timeZone);
    if (monthKey) monthlyRows[monthKey] = row;
  });

  const monthWindow = getMonthlyQualityMonthWindow_(context.year, context.month, 3);
  const sections = {};

  MONTHLY_QUALITY_FPY_DECK_SECTIONS_.forEach(function(section) {
    const startColumn = startColumns[section.sheetHeader];
    const ytd = { inspected: 0, defects: 0, fpy: 0 };

    Object.keys(monthlyRows).forEach(function(monthKey) {
      const parts = monthKey.split('-').map(Number);
      if (parts[0] === context.year && parts[1] <= context.month) {
        const record = readMonthlyQualityFPYRecord_(
          monthlyRows[monthKey],
          startColumn
        );
        ytd.inspected += record.inspected;
        ytd.defects += record.defects;
      }
    });
    ytd.fpy = calculateMonthlyQualityFPY_(ytd.inspected, ytd.defects);

    sections[section.sheetHeader] = {
      ytd: ytd,
      months: monthWindow.map(function(month) {
        const row = monthlyRows[month.monthKey];
        const record = row
          ? readMonthlyQualityFPYRecord_(row, startColumn)
          : { inspected: 0, defects: 0, fpy: 0 };
        record.monthKey = month.monthKey;
        record.displayLabel = month.displayLabel;
        return record;
      })
    };
  });

  return {
    months: monthWindow,
    sections: sections,
    plantAverage: sections['All Lines'].ytd.fpy
  };
}

function getMonthlyQualitySheetMonthKey_(value, timeZone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timeZone, 'yyyy-MM');
  }

  const text = String(value || '').trim();
  const match = text.match(/^(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!match) return '';
  const year = match[2].length === 2 ? 2000 + Number(match[2]) : Number(match[2]);
  return year + '-' + String(Number(match[1])).padStart(2, '0');
}

function getMonthlyQualityMonthWindow_(year, month, count) {
  const result = [];
  for (let offset = count - 1; offset >= 0; offset--) {
    const date = new Date(Date.UTC(year, month - 1 - offset, 1));
    const resultYear = date.getUTCFullYear();
    const resultMonth = date.getUTCMonth() + 1;
    result.push({
      monthKey: resultYear + '-' + String(resultMonth).padStart(2, '0'),
      displayLabel:
        String(resultMonth).padStart(2, '0') + '/' + String(resultYear).slice(-2)
    });
  }
  return result;
}

function readMonthlyQualityFPYRecord_(row, startColumn) {
  const inspected = toMonthlyQualityNumber_(row[startColumn]);
  const defects = toMonthlyQualityNumber_(row[startColumn + 1]);
  return {
    inspected: inspected,
    defects: defects,
    fpy: calculateMonthlyQualityFPY_(inspected, defects)
  };
}

function toMonthlyQualityNumber_(value) {
  const number = Number(value);
  return isFinite(number) ? number : 0;
}

function calculateMonthlyQualityFPY_(inspected, defects) {
  return inspected > 0 ? (inspected - defects) / inspected : 0;
}

function findMonthlyQualitySlide_(presentation, slideLabel) {
  const expectedTitle = 'Quality Status Updates - ' + slideLabel;
  const slides = presentation.getSlides();

  for (let index = 0; index < slides.length; index++) {
    const elements = slides[index].getPageElements();
    for (let elementIndex = 0; elementIndex < elements.length; elementIndex++) {
      const element = elements[elementIndex];
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
      const text = element.asShape().getText().asString().trim();
      if (text === expectedTitle) return slides[index];
    }
  }

  throw new Error('Monthly Quality slide not found: ' + expectedTitle);
}

function findMonthlyQualityFPYTable_(slide) {
  const elements = slide.getPageElements();
  for (let index = 0; index < elements.length; index++) {
    const element = elements[index];
    if (element.getPageElementType() !== SlidesApp.PageElementType.TABLE) continue;
    const table = element.asTable();
    if (table.getNumRows() < 7 || table.getNumColumns() < 4) continue;
    const firstCell = table.getCell(0, 0).getText().asString().trim();
    if (firstCell === 'Month/Year') return table;
  }

  throw new Error('The expected 7-row FPY table was not found on a quality slide.');
}

function writeMonthlyQualityFPYTable_(table, section, data) {
  const sectionData = data.sections[section.sheetHeader];
  setMonthlyQualityTableText_(table, 0, 0, 'Month/Year');
  setMonthlyQualityTableText_(table, 0, 1, section.tableLabel);
  setMonthlyQualityTableText_(table, 1, 0, 'Plant Average');
  setMonthlyQualityTableText_(
    table,
    2,
    0,
    formatMonthlyQualityPercent_(data.plantAverage)
  );
  setMonthlyQualityTableText_(table, 2, 1, 'Total Inspected');
  setMonthlyQualityTableText_(table, 2, 2, 'Defects');
  setMonthlyQualityTableText_(table, 2, 3, 'FPY Inspection');
  writeMonthlyQualityFPYRecordRow_(table, 3, 'Current Year Total', sectionData.ytd);

  sectionData.months.forEach(function(record, index) {
    writeMonthlyQualityFPYRecordRow_(table, 4 + index, record.displayLabel, record);
  });
}

function writeMonthlyQualityFPYRecordRow_(table, rowIndex, label, record) {
  setMonthlyQualityTableText_(table, rowIndex, 0, label);
  setMonthlyQualityTableText_(table, rowIndex, 1, String(Math.round(record.inspected)));
  setMonthlyQualityTableText_(table, rowIndex, 2, String(Math.round(record.defects)));
  setMonthlyQualityTableText_(
    table,
    rowIndex,
    3,
    formatMonthlyQualityPercent_(record.fpy)
  );
}

function setMonthlyQualityTableText_(table, rowIndex, columnIndex, value) {
  table.getCell(rowIndex, columnIndex).getText().setText(String(value));
}

function formatMonthlyQualityPercent_(value) {
  return (toMonthlyQualityNumber_(value) * 100).toFixed(2) + '%';
}

function formatMonthlyQualityFPYRecord_(record) {
  return {
    monthKey: record.monthKey || null,
    label: record.displayLabel || null,
    inspected: Math.round(record.inspected),
    defects: Math.round(record.defects),
    fpy: formatMonthlyQualityPercent_(record.fpy)
  };
}
