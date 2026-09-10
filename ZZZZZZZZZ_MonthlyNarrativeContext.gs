/**
 * ChatGPT-authored narrative context for the Monthly Quality report.
 *
 * The numeric report is deterministic and belongs in Apps Script. The plain-English
 * explanation of what those failures mean for the floor needs authored context.
 *
 * This file does NOT stop at creating a data tab. It uses the tab as an editable
 * source, then inserts one "Top 3 Highest-Impact Issues" context slide after the
 * Plant Summary, MSC, CSC, and ARU metric slides.
 */

const MONTHLY_QUALITY_CONTEXT_SHEET_NAME_ = 'Monthly Report Context';

const MONTHLY_QUALITY_CONTEXT_SECTIONS_ = Object.freeze([
  'All Lines',
  'MSC',
  'CSC',
  'ARU'
]);

const MONTHLY_QUALITY_CONTEXT_HEADERS_ = Object.freeze([
  'Report Month',
  'Section',
  'Headline',
  'Issue 1 - Highest Impact',
  'Why Issue 1 Matters',
  'Issue 2',
  'Why Issue 2 Matters',
  'Issue 3',
  'Why Issue 3 Matters',
  'Floor Takeaway',
  'What We Are Doing'
]);

const MONTHLY_QUALITY_CHATGPT_CONTEXT_OVERRIDES_ = Object.freeze({
  '2026-08': Object.freeze({
    'All Lines': Object.freeze({
      headline: '17 MSC / CSC / ARU customer failure tickets were counted in August.',
      issues: Object.freeze([
        Object.freeze({
          title: 'Electrical / controls failures across multiple product lines',
          why: 'Relays/contactors, fan-proving switches, refrigerant monitors, temperature sensors, and flow devices all showed up in August. These failures create alarms, downtime, or failed startup conditions for customers.'
        }),
        Object.freeze({
          title: 'Refrigeration reliability and leak concerns',
          why: 'Compressor failures, drive trips, TXV concerns, and refrigerant leaks were part of the August mix. These are high-impact because they usually stop the unit from operating normally and create warranty cost.'
        }),
        Object.freeze({
          title: 'Shipment completeness and field-readiness problems',
          why: 'Missing fittings, missing sensors, and cabinet/filter-door fit concerns are easier for the floor to understand and prevent. They should be caught before the unit or parts leave MJC.'
        })
      ]),
      takeaway: 'The number is not just a score. It tells us what reached the customer and where factory checks, shipment verification, and corrective actions need to stop repeat problems.',
      actions: Object.freeze([
        'Keep questions and documentation-only requests out of the failure count.',
        'Use ticket descriptions to connect DPPM numbers to real customer problems.',
        'Drive repeat electrical, refrigeration, and shipment-completeness issues into CAPA or focused internal follow-up.'
      ])
    }),
    'MSC': Object.freeze({
      headline: 'MSC had 7 reported failure tickets in August.',
      issues: Object.freeze([
        Object.freeze({
          title: 'Compressor / drive / refrigeration reliability',
          why: 'MSC tickets included compressor failure or over-amping, a compressor drive fault, and TXV concerns. These are severe because they affect whether the unit can run and usually require field repair or warranty parts.'
        }),
        Object.freeze({
          title: 'Missing ship-with material',
          why: 'Startup tickets included missing grooved fittings and missing water-pressure sensors. These issues delay startup and are directly tied to shipment verification before the product leaves MJC.'
        }),
        Object.freeze({
          title: 'Flow-switch / safety-circuit repeat concerns',
          why: 'Repeated flow-switch or safety-circuit concerns create field troubleshooting time and repeat customer frustration. This is a focused follow-up item, not just a one-time service call.'
        })
      ]),
      takeaway: 'Several MSC problems are tied to things the plant can influence: complete shipment checks, wiring / controls verification, and catching abnormal refrigeration or component conditions before release.',
      actions: Object.freeze([
        'Reinforce ship-with verification before shipment.',
        'Use the issue list to target final inspection checks for relays, flow switches, wiring, and refrigeration components.',
        'Treat repeated flow-switch / safety-circuit conditions as a focused follow-up item.'
      ])
    }),
    'CSC': Object.freeze({
      headline: 'CSC had 5 reported failure tickets in August.',
      issues: Object.freeze([
        Object.freeze({
          title: 'Refrigerant leak repairs',
          why: 'Two CSC tickets involved refrigerant leak repairs. These are high-impact because they affect unit operation, require field labor/refrigerant, and should reinforce leak-prevention and leak-check discipline.'
        }),
        Object.freeze({
          title: 'Electrical component failures',
          why: 'CSC tickets included a compressor contactor failure and a failed flow switch. These failures can trip breakers, create alarms, or stop normal operation.'
        }),
        Object.freeze({
          title: 'Fan / airflow device reliability',
          why: 'One CSC issue involved an ECM supply fan not starting consistently. Fan and airflow device verification remains important during test and inspection.'
        })
      ]),
      takeaway: 'CSC issues show why electrical checks, fan / flow device verification, and leak-prevention discipline remain critical before units leave MJC.',
      actions: Object.freeze([
        'Continue verifying contactors, flow devices, and ECM fan operation during test / inspection.',
        'Use leak tickets to reinforce refrigeration workmanship and leak-check expectations.',
        'Watch for repeat components that need supplier or design follow-up.'
      ])
    }),
    'ARU': Object.freeze({
      headline: 'ARU had 5 reported failure tickets in August.',
      issues: Object.freeze([
        Object.freeze({
          title: 'Repeat fan-proving switch failures',
          why: 'Fan-proving switch failures were reported again across Niagara ARUs. Repeat field failures need returned-part analysis and clear ownership so they do not become recurring customer issues.'
        }),
        Object.freeze({
          title: 'Electrical / controls device failures',
          why: 'ARU tickets included a failed refrigerant monitor and a faulty discharge-air temperature sensor. These devices can create safety, alarm, or control problems in the field.'
        }),
        Object.freeze({
          title: 'Coil leak and cabinet / filter-door fit concerns',
          why: 'One ARU evaporator-coil leak remained under investigation, and one cabinet/filter-door condition repeated a prior field issue. Both need evidence-based follow-up rather than one-off fixes.'
        })
      ]),
      takeaway: 'ARU issues are highly custom and often field-specific, but repeat device failures and cabinet-fit concerns still need clear ownership and follow-through.',
      actions: Object.freeze([
        'Track returned fan-proving switches and failed devices for analysis.',
        'Keep coil leak evidence tied to the unit and repair decision.',
        'Use the repeat filter-door issue to drive a standard fix, not a one-off repair.'
      ])
    })
  })
});

var updateMonthlyQualityPackageDPPM_BEFORE_NARRATIVE_ = updateMonthlyQualityPackageDPPM_;

updateMonthlyQualityPackageDPPM_ = function(packageResult) {
  const result = updateMonthlyQualityPackageDPPM_BEFORE_NARRATIVE_(packageResult);

  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());
  const issueResult = result.validatedIssueChartData ||
    readMonthlyQualityExistingIssueChartData_(spreadsheet);

  const contextResult = prepareMonthlyQualityNarrativeContext_(spreadsheet, issueResult);

  if (!contextResult || !contextResult.sections || !contextResult.sections.length) {
    Logger.log('Monthly Quality narrative context skipped: no context rows found.');
    result.narrativeContext = {
      status: 'SKIPPED',
      reason: 'No monthly narrative context rows were available.'
    };
    return result;
  }

  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const slideResult = updateMonthlyQualityNarrativeContextSlides_(presentation, contextResult);
  presentation.saveAndClose();

  result.narrativeContext = {
    status: 'READY',
    reportMonth: contextResult.reportMonth,
    monthLabel: contextResult.monthLabel,
    contextSource: contextResult.source,
    slidesUpdated: slideResult.slidesUpdated
  };

  Logger.log(JSON.stringify(result.narrativeContext));
  return result;
};

function prepareMonthlyQualityNarrativeContext_(spreadsheet, issueResult) {
  const reportMonth = String(issueResult.reportMonth || '').trim();
  if (!reportMonth) {
    throw new Error('Monthly narrative context cannot determine report month.');
  }

  const monthLabel = monthlyQualityNarrativeMonthLabel_(reportMonth);
  const sheet = ensureMonthlyQualityNarrativeContextSheet_(spreadsheet);

  seedMonthlyQualityNarrativeContextRows_(sheet, reportMonth);

  const rows = readMonthlyQualityNarrativeContextRows_(sheet, reportMonth);
  return {
    reportMonth: reportMonth,
    monthLabel: monthLabel,
    source: MONTHLY_QUALITY_CONTEXT_SHEET_NAME_,
    sections: rows
  };
}

function ensureMonthlyQualityNarrativeContextSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(MONTHLY_QUALITY_CONTEXT_SHEET_NAME_);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(MONTHLY_QUALITY_CONTEXT_SHEET_NAME_);
  }

  const lastColumn = Math.max(sheet.getLastColumn(), MONTHLY_QUALITY_CONTEXT_HEADERS_.length);
  const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function(value) { return String(value || '').trim(); });
  const expectedHeaders = MONTHLY_QUALITY_CONTEXT_HEADERS_.slice();
  const headersMatch = expectedHeaders.every(function(header, index) {
    return existingHeaders[index] === header;
  });

  if (!headersMatch) {
    sheet.clear();
  }

  sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), expectedHeaders.length).setWrap(true);
  sheet.autoResizeColumns(1, expectedHeaders.length);
  return sheet;
}

function seedMonthlyQualityNarrativeContextRows_(sheet, reportMonth) {
  const existing = readMonthlyQualityNarrativeContextRows_(sheet, reportMonth);
  if (existing.length) return false;

  const override = MONTHLY_QUALITY_CHATGPT_CONTEXT_OVERRIDES_[reportMonth];
  if (!override) return false;

  const rows = MONTHLY_QUALITY_CONTEXT_SECTIONS_.map(function(section) {
    const item = override[section] || {};
    const issues = item.issues || [];
    return [
      reportMonth,
      section,
      item.headline || '',
      issues[0] ? issues[0].title : '',
      issues[0] ? issues[0].why : '',
      issues[1] ? issues[1].title : '',
      issues[1] ? issues[1].why : '',
      issues[2] ? issues[2].title : '',
      issues[2] ? issues[2].why : '',
      item.takeaway || '',
      Array.isArray(item.actions) ? item.actions.join('\n') : String(item.actions || '')
    ];
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(startRow, 4, rows.length, 8).setWrap(true);
  sheet.autoResizeColumns(1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length);
  return true;
}

function readMonthlyQualityNarrativeContextRows_(sheet, reportMonth) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length).getValues();
  const rows = [];

  values.forEach(function(row) {
    const month = String(row[0] || '').trim();
    const section = String(row[1] || '').trim();
    if (month !== reportMonth || MONTHLY_QUALITY_CONTEXT_SECTIONS_.indexOf(section) < 0) {
      return;
    }

    rows.push({
      section: section,
      headline: String(row[2] || '').trim(),
      issues: [
        { title: String(row[3] || '').trim(), why: String(row[4] || '').trim() },
        { title: String(row[5] || '').trim(), why: String(row[6] || '').trim() },
        { title: String(row[7] || '').trim(), why: String(row[8] || '').trim() }
      ].filter(function(issue) { return issue.title || issue.why; }),
      takeaway: String(row[9] || '').trim(),
      actions: monthlyQualityNarrativeSplitLines_(row[10])
    });
  });

  return rows;
}

function monthlyQualityNarrativeSplitLines_(value) {
  return String(value || '')
    .split(/\r?\n|\s*;\s*/)
    .map(function(item) { return item.trim(); })
    .filter(function(item) { return item; });
}

function updateMonthlyQualityNarrativeContextSlides_(presentation, contextResult) {
  removeMonthlyQualityNarrativeContextSlides_(presentation);

  let slidesUpdated = 0;
  const bySection = {};
  contextResult.sections.forEach(function(row) {
    bySection[row.section] = row;
  });

  MONTHLY_QUALITY_CONTEXT_SECTIONS_.slice().reverse().forEach(function(section) {
    const record = bySection[section];
    if (!record) return;

    const mainIndex = findMonthlyQualitySlideIndex_(presentation, section);
    const slide = presentation.insertSlide(mainIndex + 1, SlidesApp.PredefinedLayout.BLANK);

    drawMonthlyQualityNarrativeSlide_(slide, section, record, contextResult.monthLabel);
    slidesUpdated++;
  });

  return { slidesUpdated: slidesUpdated };
}

function removeMonthlyQualityNarrativeContextSlides_(presentation) {
  const slides = presentation.getSlides();
  for (let index = slides.length - 1; index >= 0; index--) {
    if (isMonthlyQualityNarrativeContextSlide_(slides[index])) {
      slides[index].remove();
    }
  }
}

function isMonthlyQualityNarrativeContextSlide_(slide) {
  const elements = slide.getPageElements();
  for (let index = 0; index < elements.length; index++) {
    const element = elements[index];
    if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
    const text = element.asShape().getText().asString().trim();
    if (text.indexOf('Top 3 Customer Issue Context - ') === 0) return true;
    if (text.indexOf('Customer Issue Context - ') === 0) return true;
  }
  return false;
}

function findMonthlyQualitySlideIndex_(presentation, slideLabel) {
  const expectedTitle = 'Quality Status Updates - ' + slideLabel;
  const slides = presentation.getSlides();

  for (let index = 0; index < slides.length; index++) {
    const elements = slides[index].getPageElements();
    for (let elementIndex = 0; elementIndex < elements.length; elementIndex++) {
      const element = elements[elementIndex];
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;
      const text = element.asShape().getText().asString().trim();
      if (text === expectedTitle) return index;
    }
  }

  throw new Error('Monthly Quality slide not found for narrative context: ' + expectedTitle);
}

function drawMonthlyQualityNarrativeSlide_(slide, section, record, monthLabel) {
  // Standard wide Google Slides deck size in points. Using constants avoids a
  // Slide.getParent() call, which is not available in Apps Script.
  const pageWidth = 720;
  const pageHeight = 405;
  const margin = 28;
  const gap = 14;
  const titleHeight = 38;
  const footerHeight = 16;
  const top = 72;
  const issueColWidth = 424;
  const rightColWidth = pageWidth - (margin * 2) - gap - issueColWidth;
  const issueBoxHeight = 78;
  const issueGap = 10;
  const rightBoxHeight = 248;

  slide.getBackground().setSolidFill('#FFFFFF');

  const displaySection = section === 'All Lines' ? 'Plant Summary' : section;
  const title = slide.insertTextBox(
    'Top 3 Customer Issue Context - ' + displaySection,
    margin,
    20,
    pageWidth - (margin * 2),
    titleHeight
  );
  title.setTitle('Top 3 Customer Issue Context - ' + displaySection);
  title.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(22)
    .setBold(true)
    .setForegroundColor('#444444');

  const subtitle = slide.insertTextBox(
    monthLabel + ' - highest-impact issue groups behind the numbers',
    margin,
    52,
    pageWidth - (margin * 2),
    18
  );
  subtitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor('#666666');

  const headline = slide.insertTextBox(
    record.headline,
    margin,
    top,
    pageWidth - (margin * 2),
    26
  );
  headline.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor('#222222');

  for (let index = 0; index < 3; index++) {
    const issue = record.issues[index] || { title: '', why: '' };
    const y = top + 34 + (index * (issueBoxHeight + issueGap));
    drawMonthlyQualitySeverityIssueBox_(
      slide,
      margin,
      y,
      issueColWidth,
      issueBoxHeight,
      index + 1,
      issue.title,
      issue.why
    );
  }

  const rightText = [
    'Floor takeaway',
    record.takeaway,
    '',
    'What we are doing',
    monthlyQualityNarrativeBullets_(record.actions)
  ].join('\n');

  const rightBox = drawMonthlyQualityNarrativeBox_(
    slide,
    margin + issueColWidth + gap,
    top + 34,
    rightColWidth,
    rightBoxHeight,
    rightText,
    '#FFF7CC'
  );
  styleMonthlyQualityNarrativeText_(rightBox, 9);

  const footer = slide.insertTextBox(
    'Source: ChatGPT summary of audited HubSpot ticket descriptions. Questions and documentation-only requests are excluded from failure counts.',
    margin,
    pageHeight - margin - footerHeight + 3,
    pageWidth - (margin * 2),
    footerHeight
  );
  footer.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(7)
    .setForegroundColor('#777777');
}

function drawMonthlyQualitySeverityIssueBox_(slide, x, y, w, h, rank, title, why) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  shape.getFill().setSolidFill('#EEF3FA');
  shape.getBorder().getLineFill().setSolidFill('#9FBAD7');
  shape.setTitle('Monthly Quality Top Severity Issue ' + rank);

  const text = '#'+ rank + '  ' + title + '\n' + why;
  shape.getText().setText(text);
  shape.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(9)
    .setForegroundColor('#222222');

  try {
    shape.getText().getRange(0, String('#' + rank + '  ' + title).length)
      .getTextStyle()
      .setBold(true)
      .setFontSize(10);
  } catch (error) {
    // If Apps Script cannot style the range, keep the full text readable.
  }

  return shape;
}

function drawMonthlyQualityNarrativeBox_(slide, x, y, w, h, text, fillColor) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  shape.getFill().setSolidFill(fillColor);
  shape.getBorder().getLineFill().setSolidFill('#B7B7B7');
  shape.getText().setText(text);
  shape.setTitle('Monthly Quality Narrative Context');
  return shape;
}

function styleMonthlyQualityNarrativeText_(shape, fontSize) {
  shape.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(fontSize)
    .setForegroundColor('#222222');
}

function monthlyQualityNarrativeBullets_(items) {
  if (!items || !items.length) return '';
  return items.map(function(item) { return '- ' + item; }).join('\n');
}

function monthlyQualityNarrativeMonthLabel_(monthKey) {
  const match = /^(\d{4})-(\d{2})$/.exec(String(monthKey || '').trim());
  if (!match) return String(monthKey || '');

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!year || month < 1 || month > 12) return String(monthKey || '');
  return monthNames[month - 1] + ' ' + year;
}
