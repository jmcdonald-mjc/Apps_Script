/**
 * Final safety wrapper for Monthly Quality narrative slides.
 *
 * The narrative context sheet is useful as an editable source, but the report
 * is not complete unless the Google Slides deck gets the context slides too.
 *
 * This file intentionally sorts after the master refresh wrapper and after the
 * narrative helpers. It runs the existing one-button refresh first, then forces
 * the narrative slides into the finished monthly deck as the final step.
 */

var calculateFPYSummary_FINAL_BEFORE_NARRATIVE_SLIDES_ = calculateFPYSummary_FINAL;

calculateFPYSummary_FINAL = function() {
  const result = calculateFPYSummary_FINAL_BEFORE_NARRATIVE_SLIDES_();

  Logger.log(
    'Master Quality Refresh: forcing Monthly Quality narrative context slides into the finished report.'
  );

  const narrativeResult = forceMonthlyQualityNarrativeSlidesAfterReport_();

  if (result && typeof result === 'object') {
    result.narrativeContext = narrativeResult;
  }

  Logger.log(JSON.stringify({
    status: 'COMPLETE_WITH_NARRATIVE_SLIDES',
    narrativeContext: narrativeResult
  }));

  return result;
};

/**
 * Add/rebuild Top 3 context slides after the report is fully generated.
 * This deliberately runs after updateMonthlyQualityReport() has finished so no
 * later report step can overwrite or recreate the deck without the context.
 */
function forceMonthlyQualityNarrativeSlidesAfterReport_(asOfDate) {
  const context = getMonthlyQualityPackageContext_(asOfDate);
  const packageResult = ensureMonthlyQualityReportPackageUnlocked_(context);
  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());

  const issueResult = readMonthlyQualityExistingIssueChartData_(spreadsheet);
  const contextResult = prepareMonthlyQualityNarrativeContext_(
    spreadsheet,
    issueResult
  );

  if (!contextResult || !contextResult.sections || !contextResult.sections.length) {
    throw new Error(
      'Monthly narrative context exists as a sheet, but no context rows were available for ' +
      context.monthKey + '. The Slides report was not updated.'
    );
  }

  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const slideResult = updateMonthlyQualityNarrativeContextSlides_(
    presentation,
    contextResult
  );
  presentation.saveAndClose();

  if (!slideResult || slideResult.slidesUpdated !== MONTHLY_QUALITY_CONTEXT_SECTIONS_.length) {
    throw new Error(
      'Monthly narrative context slide creation did not complete. Expected ' +
      MONTHLY_QUALITY_CONTEXT_SECTIONS_.length + ' slides, created ' +
      (slideResult ? slideResult.slidesUpdated : 'none') + '.'
    );
  }

  return {
    status: 'READY',
    reportMonth: context.monthKey,
    monthLabel: context.monthLabel,
    deckId: packageResult.deckFile.getId(),
    dataId: packageResult.dataFile.getId(),
    slidesUpdated: slideResult.slidesUpdated,
    slideType: 'Top 3 Customer Issue Context'
  };
}
