/**
 * Enforce the Monthly Quality package build order:
 *
 *   1. Create/reuse the month folder.
 *   2. Create/reuse the Google Sheet by copying the approved monthly template.
 *   3. Configure and populate the Google Sheet.
 *   4. Only after the data stage succeeds, create/reuse the Slides report.
 *   5. Populate/refresh the report from the completed monthly Google Sheet.
 *
 * This file intentionally loads after the earlier monthly-report overrides so
 * this implementation of updateMonthlyQualityReport() is the active runner.
 */

function updateMonthlyQualityReport(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const context = getMonthlyQualityPackageContext_(asOfDate);

    // DATA STAGE: no new Slides deck is created before this succeeds.
    const dataStage = ensureMonthlyQualityDataStageUnlocked_(context);
    const dataResult = populateMonthlyQualityDataStage_(context, dataStage);

    // REPORT STAGE: only starts after the monthly workbook is fully prepared.
    const packageResult = ensureMonthlyQualityReportStageUnlocked_(
      context,
      dataStage
    );

    const fpyResult = updateMonthlyQualityPackageFPY_(context, packageResult);
    const dppmResult = updateMonthlyQualityPackageDPPM_(packageResult);

    const result = describeMonthlyQualityReportPackage_(packageResult, context);
    result.buildSequence = [
      'COPY_MONTHLY_SHEET_TEMPLATE',
      'POPULATE_MONTHLY_SHEET',
      'CREATE_MONTHLY_REPORT',
      'POPULATE_MONTHLY_REPORT'
    ];
    result.dataStage = dataResult;
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
 * Create/reuse only the month folder and Google Sheet. The Slides template is
 * deliberately not copied here.
 */
function ensureMonthlyQualityDataStageUnlocked_(context) {
  const config = getMonthlyQualityAutomationConfig_();
  const rootFolder = DriveApp.getFolderById(config.outputFolderId);
  const dataTemplate = findMonthlyQualityDPPMTemplate_(rootFolder);
  const folderResult = ensureMonthlyQualityPackageFolder_(rootFolder, context);

  const dataResult = ensureMonthlyQualityPackageFile_(
    folderResult.folder,
    rootFolder,
    context.dataName,
    MimeType.GOOGLE_SHEETS,
    dataTemplate,
    monthlyQualityPackagePropertyKey_('DATA', context.monthKey)
  );

  // Set the report month and extend the DPPM model on the copied workbook.
  configureMonthlyQualityPackageData_(
    dataResult.file,
    context,
    dataResult.created
  );

  PropertiesService.getScriptProperties().setProperty(
    monthlyQualityPackagePropertyKey_('FOLDER', context.monthKey),
    folderResult.folder.getId()
  );

  return {
    rootFolder: rootFolder,
    folder: folderResult.folder,
    dataFile: dataResult.file,
    createdFolder: folderResult.created,
    createdData: dataResult.created
  };
}

/**
 * Populate every currently automated workbook data source before Slides exists.
 *
 * HubSpot is authoritative for service/quality issue counts. If that refresh
 * fails, throw here so the run does not produce a half-built monthly report.
 * The DPPM model is recalculated after the source tables are refreshed.
 */
function populateMonthlyQualityDataStage_(context, dataStage) {
  const spreadsheetId = dataStage.dataFile.getId();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const hubSpotSync = syncHubSpotSupportTicketsToSpreadsheet_(spreadsheetId);
  const issueChartData = refreshMonthlyQualityValidatedIssueChartData_(
    spreadsheet
  );

  ensureMonthlyQualityPackageDPPMModel_(spreadsheet, context);
  SpreadsheetApp.flush();

  const configSheet = spreadsheet.getSheetByName(
    MONTHLY_QUALITY_PACKAGE_DPPM_CONFIG_SHEET_
  );
  if (configSheet) {
    configSheet.getRange('B12').setValue(
      'Monthly sheet populated before report creation'
    );
  }

  return {
    spreadsheetId: spreadsheetId,
    spreadsheetName: dataStage.dataFile.getName(),
    hubSpotSync: hubSpotSync,
    issueChartData: issueChartData,
    reportCreatedAfterDataStage: true
  };
}

/**
 * Create/reuse the Slides report only after populateMonthlyQualityDataStage_()
 * has completed successfully.
 */
function ensureMonthlyQualityReportStageUnlocked_(context, dataStage) {
  const config = getMonthlyQualityAutomationConfig_();
  const deckResult = ensureMonthlyQualityPackageFile_(
    dataStage.folder,
    dataStage.rootFolder,
    context.deckName,
    MimeType.GOOGLE_SLIDES,
    DriveApp.getFileById(config.templateId),
    monthlyQualityPackagePropertyKey_('DECK', context.monthKey)
  );

  return {
    folder: dataStage.folder,
    deckFile: deckResult.file,
    dataFile: dataStage.dataFile,
    createdFolder: dataStage.createdFolder,
    createdDeck: deckResult.created,
    createdData: dataStage.createdData
  };
}
