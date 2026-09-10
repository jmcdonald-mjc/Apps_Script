/**
 * ChatGPT-authored narrative context for the Monthly Quality report.
 *
 * The numeric report is deterministic and belongs in Apps Script. The plain-English
 * explanation of what those failures mean for the floor needs authored context.
 * This file adds that bridge without manually editing the Slides deck:
 *   1. the standard monthly report finalizer runs first;
 *   2. a "Monthly Report Context" sheet is seeded/read for the report month;
 *   3. one context slide is added after Plant Summary, MSC, CSC, and ARU;
 *   4. reruns remove and recreate the context slides so duplicates are avoided.
 *
 * Future month workflow:
 *   - ChatGPT reviews the audited HubSpot tickets for the month;
 *   - ChatGPT or the user updates the "Monthly Report Context" tab;
 *   - calculateFPYSummary_FINAL() is rerun.
 */

const MONTHLY_QUALITY_CONTEXT_SHEET_NAME_ = 'Monthly Report Context';

const MONTHLY_QUALITY_CONTEXT_SECTIONS_ = Object.freeze([
  'All Lines',
  'MSC',
  'CSC',
  'ARU'
]);

const MONTHLY_QUALITY_CHATGPT_CONTEXT_OVERRIDES_ = Object.freeze({
  '2026-08': Object.freeze({
    'All Lines': Object.freeze({
      headline: '17 MSC / CSC / ARU customer failure tickets were counted in August.',
      experienced: [
        'The issues were not all the same type: 3 startup, 13 warranty, and 1 service ticket.',
        'The largest themes were electrical / controls, refrigeration, missing shipment items, and repeat field conditions.',
        'One separate coatings warranty issue was also logged for cosmetic damage / cover latch fit.'
      ],
      themes: ['Electrical / Controls', 'Refrigeration', 'Ship-With Completeness', 'Repeat Field Conditions'],
      takeaway: 'The number is not just a score. It tells us what reached the customer and where our factory checks, shipment verification, and corrective actions need to prevent repeats.',
      actions: [
        'Use ticket descriptions to connect DPPM numbers to real customer problems.',
        'Keep questions and documentation requests out of the failure count.',
        'Drive repeat issues into CAPA or focused internal follow-up.'
      ]
    }),
    'MSC': Object.freeze({
      headline: 'MSC had 7 reported failure tickets in August.',
      experienced: [
        'Startup issues included missing grooved fittings and missing water-pressure sensors.',
        'Warranty issues included failed compressor-control relays, TXV concerns, compressor failures, and repeated flow-switch / safety-circuit concerns.',
        'One service issue involved a compressor drive fault shortly after startup.'
      ],
      themes: ['Ship-With Completeness', 'Electrical / Controls', 'Refrigeration Components', 'Flow / Safety Circuit'],
      takeaway: 'Several MSC problems are directly tied to things the plant can influence: complete shipment checks, wiring / controls verification, and catching abnormal refrigeration or component issues before release.',
      actions: [
        'Reinforce ship-with verification before shipment.',
        'Use the issue list to target final inspection checks for wiring, relays, flow switches, and refrigeration components.',
        'Treat repeated flow-switch / safety-circuit conditions as a focused follow-up item.'
      ]
    }),
    'CSC': Object.freeze({
      headline: 'CSC had 5 reported failure tickets in August.',
      experienced: [
        'Customers reported a failed compressor contactor, a failed flow switch, and an ECM supply-fan issue.',
        'Two refrigerant leak repair tickets were also counted for CSC units.',
        'The month was driven by electrical component reliability and refrigeration leak concerns.'
      ],
      themes: ['Electrical Components', 'Refrigeration Leaks', 'Fan / Flow Devices'],
      takeaway: 'CSC issues show why electrical checks, fan / flow device verification, and leak-prevention discipline remain critical before units leave MJC.',
      actions: [
        'Continue verifying contactors, flow devices, and ECM fan operation during test / inspection.',
        'Use leak tickets to reinforce refrigeration workmanship and leak-check expectations.',
        'Watch for repeat components that need supplier or design follow-up.'
      ]
    }),
    'ARU': Object.freeze({
      headline: 'ARU had 5 reported failure tickets in August.',
      experienced: [
        'Electrical / controls items included fan-proving switch failures, a bad refrigerant monitor, and a faulty discharge-air temperature sensor.',
        'One evaporator-coil leak remained under investigation.',
        'One cabinet / filter-door condition repeated a prior field issue.'
      ],
      themes: ['Electrical / Controls Devices', 'Refrigerant Leak', 'Cabinet / Door Fit', 'Repeat Conditions'],
      takeaway: 'ARU issues are highly custom and often field-specific, but repeat device failures and cabinet-fit concerns still need clear ownership and follow-through.',
      actions: [
        'Track returned fan-proving switches and failed devices for analysis.',
        'Keep coil leak evidence tied to the unit and repair decision.',
        'Use the repeat filter-door issue to drive a standard fix, not a one-off repair.'
      ]
    })
  })
});

var updateMonthlyQualityPackageDPPM_BEFORE_NARRATIVE_ = updateMonthlyQualityPackageDPPM_;

updateMonthlyQualityPackageDPPM_ = function(packageResult) {
  const result = updateMonthlyQualityPackageDPPM_BEFORE_NARRATIVE_(packageResult);

  const spreadsheet = SpreadsheetApp.openById(packageResult.dataFile.getId());
  const issueResult = result.validatedIssueChartData ||
    readMonthlyQualityExistingIssueChartData_(spreadsheet);

  const contextResult = prepareMonthlyQualityNarrativeContext_(
    spreadsheet,
    issueResult
  );

  if (!contextResult || !contextResult.sections || !contextResult.sections.length) {
    Logger.log('Monthly Quality narrative context skipped: no context rows found.');
    result.narrativeContext = {
      status: 'SKIPPED',
      reason: 'No monthly narrative context rows were available.'
    };
    return result;
  }

  const presentation = SlidesApp.openById(packageResult.deckFile.getId());
  const slideResult = updateMonthlyQualityNarrativeContextSlides_(
    presentation,
    contextResult
  );
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

  const headers = [[
    'Report Month',
    'Section',
    'Headline',
    'What Customers Experienced',
    'Main Themes',
    'Floor Takeaway',
    'What We Are Doing'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers[0].length);
  return sheet;
}

function seedMonthlyQualityNarrativeContextRows_(sheet, reportMonth) {
  const existing = readMonthlyQualityNarrativeContextRows_(sheet, reportMonth);
  if (existing.length) return false;

  const override = MONTHLY_QUALITY_CHATGPT_CONTEXT_OVERRIDES_[reportMonth];
  if (!override) return false;

  const rows = MONTHLY_QUALITY_CONTEXT_SECTIONS_.map(function(section) {
    const item = override[section] || {};
    return [
      reportMonth,
      section,
      item.headline || '',
      Array.isArray(item.experienced) ? item.experienced.join('\n') : String(item.experienced || ''),
      Array.isArray(item.themes) ? item.themes.join('\n') : String(item.themes || ''),
      item.takeaway || '',
      Array.isArray(item.actions) ? item.actions.join('\n') : String(item.actions || '')
    ];
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(startRow, 4, rows.length, 4).setWrap(true);
  sheet.autoResizeColumns(1, 7);
  return true;
}

function readMonthlyQualityNarrativeContextRows_(sheet, reportMonth) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
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
      experienced: monthlyQualityNarrativeSplitLines_(row[3]),
      themes: monthlyQualityNarrativeSplitLines_(row[4]),
      takeaway: String(row[5] || '').trim(),
      actions: monthlyQualityNarrativeSplitLines_(row[6])
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

  // Insert in reverse order so each new slide lands immediately after its
  // matching metric slide without changing the positions of earlier targets.
  MONTHLY_QUALITY_CONTEXT_SECTIONS_.slice().reverse().forEach(function(section) {
    const record = bySection[section];
    if (!record) return;

    const mainIndex = findMonthlyQualitySlideIndex_(presentation, section);
    const slide = presentation.insertSlide(
      mainIndex + 1,
      SlidesApp.PredefinedLayout.BLANK
    );

    drawMonthlyQualityNarrativeSlide_(
      slide,
      section,
      record,
      contextResult.monthLabel
    );
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
  const pageWidth = slide.getParent().getPageWidth();
  const pageHeight = slide.getParent().getPageHeight();
  const margin = 36;
  const gap = 18;
  const titleHeight = 46;
  const footerHeight = 18;
  const top = margin + titleHeight + 12;
  const leftColWidth = (pageWidth - (margin * 2) - gap) * 0.58;
  const rightColWidth = pageWidth - (margin * 2) - gap - leftColWidth;
  const boxHeight = pageHeight - top - margin - footerHeight;

  slide.getBackground().setSolidFill('#FFFFFF');

  const displaySection = section === 'All Lines' ? 'Plant Summary' : section;
  const title = slide.insertTextBox(
    'Customer Issue Context - ' + displaySection,
    margin,
    24,
    pageWidth - (margin * 2),
    titleHeight
  );
  title.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(24)
    .setBold(true)
    .setForegroundColor('#444444');

  const subtitle = slide.insertTextBox(
    monthLabel + ' - what the numbers mean on the floor',
    margin,
    58,
    pageWidth - (margin * 2),
    20
  );
  subtitle.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(10)
    .setForegroundColor('#666666');

  const leftText = [
    record.headline,
    '',
    'What customers experienced',
    monthlyQualityNarrativeBullets_(record.experienced)
  ].join('\n');

  const rightText = [
    'Main themes',
    monthlyQualityNarrativeBullets_(record.themes),
    '',
    'Floor takeaway',
    record.takeaway,
    '',
    'What we are doing',
    monthlyQualityNarrativeBullets_(record.actions)
  ].join('\n');

  const leftBox = drawMonthlyQualityNarrativeBox_(
    slide,
    margin,
    top,
    leftColWidth,
    boxHeight,
    leftText,
    '#EEF3FA'
  );
  styleMonthlyQualityNarrativeText_(leftBox, 12);

  const rightBox = drawMonthlyQualityNarrativeBox_(
    slide,
    margin + leftColWidth + gap,
    top,
    rightColWidth,
    boxHeight,
    rightText,
    '#FFF7CC'
  );
  styleMonthlyQualityNarrativeText_(rightBox, 10);

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
