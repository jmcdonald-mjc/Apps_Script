/**
 * Monthly Quality Status Google Slides automation.
 *
 * Batch 3-4 scope:
 *   - Own the Slides-specific configuration in a separate module.
 *   - Store the template, output folder, and time zone in Script Properties.
 *   - Validate that the configured Google Drive resources are accessible.
 *
 * Deck creation and slide population are intentionally added in later batches.
 */

const MONTHLY_QUALITY_PROPERTY_KEYS_ = Object.freeze({
  TEMPLATE_ID: 'QUALITY_SLIDES_TEMPLATE_ID',
  OUTPUT_FOLDER_ID: 'QUALITY_REPORTS_FOLDER_ID',
  TIME_ZONE: 'QUALITY_REPORT_TIME_ZONE'
});

const MONTHLY_QUALITY_INITIAL_CONFIGURATION_ = Object.freeze({
  templateId: '1KOmS8kIPn9fGzcIhk5l0Xr-tNLjX1-jc-V0GGnIoJ7c',
  outputFolderId: '1vBoyVpqKYi102sStqLAaskQfKHKcDvO-',
  timeZone: 'America/New_York'
});

/**
 * One-time setup function. Run this after the file is pushed to Apps Script.
 * The stored values are Drive resource identifiers, not credentials.
 */
function configureMonthlyQualityAutomation() {
  const keys = MONTHLY_QUALITY_PROPERTY_KEYS_;

  PropertiesService.getScriptProperties().setProperties({
    [keys.TEMPLATE_ID]: MONTHLY_QUALITY_INITIAL_CONFIGURATION_.templateId,
    [keys.OUTPUT_FOLDER_ID]: MONTHLY_QUALITY_INITIAL_CONFIGURATION_.outputFolderId,
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

  if (templateFile.getMimeType() !== MimeType.GOOGLE_SLIDES) {
    throw new Error(
      'QUALITY_SLIDES_TEMPLATE_ID must identify a native Google Slides presentation.'
    );
  }

  const presentation = SlidesApp.openById(config.templateId);
  const result = {
    templateId: config.templateId,
    templateName: templateFile.getName(),
    templateSlideCount: presentation.getSlides().length,
    outputFolderId: config.outputFolderId,
    outputFolderName: outputFolder.getName(),
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
    timeZone: String(
      properties.getProperty(keys.TIME_ZONE) || Session.getScriptTimeZone() || ''
    ).trim()
  };

  const missing = [];
  if (!config.templateId) missing.push(keys.TEMPLATE_ID);
  if (!config.outputFolderId) missing.push(keys.OUTPUT_FOLDER_ID);
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
