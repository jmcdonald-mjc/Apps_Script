/**
 * Final override for Monthly Quality narrative slides.
 *
 * The floor-facing slide should not summarize broad issue themes. It should
 * list actual tickets: what happened, what was done / current status, and why
 * the ticket matters. This file sorts last and overrides the narrative sheet
 * preparation and slide drawing functions with the Top 3 Ticket format.
 */

var MONTHLY_QUALITY_TOP_TICKET_CONTEXT_ = Object.freeze({
  '2026-08': Object.freeze({
    'All Lines': Object.freeze({
      headline: 'Top 3 August customer tickets selected by customer impact and repeat risk.',
      tickets: Object.freeze([
        Object.freeze({
          id: '47854734444',
          product: 'MSC',
          title: 'Spalding Ridge SWUD operating issue',
          category: 'Warranty - Electrical / Controls',
          issue: 'Customer reported an ongoing critical SWUD issue. Only two compressors were running, other circuits were locked out, and the customer had lost patience after repeated visits.',
          action: 'MJC was pulled in to coordinate with the field team. The ticket calls out faulty relays/contactors and bad TXV concerns; factory support was requested to close the issue.',
          why: 'High customer impact: unit performance issue, repeated field activity, and customer escalation.'
        }),
        Object.freeze({
          id: '47452240697',
          product: 'ARU',
          title: 'KNC2 fan-proving switch failures',
          category: 'Warranty - Electrical / Controls',
          issue: 'Repeat fan-proving switch failures were reported on Niagara KNC2 ARUs, tied to the same issue previously seen at KNC.',
          action: 'Replacement approach was approved and failed switches were requested back for RMA / analysis.',
          why: 'Repeat failure across related ARUs; needs returned-part analysis so it does not become a recurring field issue.'
        }),
        Object.freeze({
          id: '47502355872',
          product: 'MSC',
          title: 'Failed Schneider relays and TXV concern',
          category: 'Warranty - Electrical / Refrigeration',
          issue: 'Field technician reported two failed Schneider compressor-control relays on one MSC and a bad TXV on another unit, plus intermittent safety-circuit alarms.',
          action: 'Replacement parts / tech support were provided and the ticket was closed. Safety-circuit / flow concerns remain a follow-up theme.',
          why: 'Multiple customer-facing failures in one support event: controls, refrigeration, and repeated alarms.'
        })
      ])
    }),
    'MSC': Object.freeze({
      headline: 'Top 3 August MSC tickets selected by customer impact and repeat risk.',
      tickets: Object.freeze([
        Object.freeze({
          id: '47854734444',
          product: 'MSC',
          title: 'Spalding Ridge SWUD operating issue',
          category: 'Warranty - Electrical / Controls',
          issue: 'Critical SWUD issue with only two compressors running, other circuits locked out, and repeated field visits without resolution.',
          action: 'MJC factory/service support was pulled in to help the field team close the relays/contactors and TXV concerns.',
          why: 'Customer escalation and unit not operating as intended.'
        }),
        Object.freeze({
          id: '47502355872',
          product: 'MSC',
          title: 'Failed Schneider relays and TXV concern',
          category: 'Warranty - Electrical / Refrigeration',
          issue: 'Two compressor-control relays were only outputting about 16V with 24V input; a circuit-2 TXV would not respond to adjustment; safety-circuit alarms were also reported.',
          action: 'Replacement/support path was handled through the ticket and closed; related flow / safety-circuit alarms should stay on the follow-up list.',
          why: 'Combines electrical controls failure, refrigeration component concern, and repeated alarms.'
        }),
        Object.freeze({
          id: '47586964573',
          product: 'MSC',
          title: 'AMC Theater repeat flow-switch failures',
          category: 'Warranty - Electrical / Controls',
          issue: 'Repeated flow-switch failures were reported across six MJC units at two AMC Theater sites; installation and wiring were also being reviewed.',
          action: 'Ticket was closed after follow-up; repeat nature should be routed into focused investigation / CAPA tracking if not already captured.',
          why: 'Repeat condition across multiple units and sites, not an isolated one-off.'
        })
      ])
    }),
    'CSC': Object.freeze({
      headline: 'Top 3 August CSC tickets selected by customer impact and repeat risk.',
      tickets: Object.freeze([
        Object.freeze({
          id: '47416132910',
          product: 'CSC',
          title: 'Compressor contactor failure - 8000023-02',
          category: 'Warranty - Electrical / Controls',
          issue: 'Compressor contactor failed, damaged the line side, and tripped breakers. Field checks indicated compressor load wiring and windings were still acceptable.',
          action: 'Replacement contactor was shipped and the ticket was closed.',
          why: 'Electrical component failure caused a clear customer outage / breaker trip event.'
        }),
        Object.freeze({
          id: '47552495015',
          product: 'CSC',
          title: 'Refrigerant leak repair - 800036-01',
          category: 'Warranty - Refrigeration Circuit',
          issue: 'CSC unit 800036-01 required leak-search and refrigerant repair in the field.',
          action: 'Leak-search / repair labor and refrigerant were approved; ticket closed.',
          why: 'Refrigerant leak repair creates direct warranty cost and should reinforce leak-check discipline.'
        }),
        Object.freeze({
          id: '47559051432',
          product: 'CSC',
          title: 'ECM supply fan intermittent start - 800048-01',
          category: 'Warranty - Airflow / Mechanical',
          issue: 'ECM supply fan 1 intermittently failed to start with the other fans and required multiple alarm resets, despite normal voltage and amperage.',
          action: 'Service follow-up was initiated through the ticket; use as a check point for fan / airflow device verification.',
          why: 'Intermittent fan operation can create repeat alarms and customer confidence issues.'
        })
      ])
    }),
    'ARU': Object.freeze({
      headline: 'Top 3 August ARU tickets selected by customer impact and repeat risk.',
      tickets: Object.freeze([
        Object.freeze({
          id: '47452240697',
          product: 'ARU',
          title: 'KNC2 fan-proving switch failures',
          category: 'Warranty - Electrical / Controls',
          issue: 'Repeat fan-proving switch failures were reported on Niagara KNC2 ARUs, matching a prior KNC issue.',
          action: 'Replacement approach was approved and failed switches were requested back for RMA / analysis.',
          why: 'Repeat device failure across related units needs ownership and returned-part analysis.'
        }),
        Object.freeze({
          id: '47675158417',
          product: 'ARU',
          title: 'Evaporator-coil leak - 1000001103',
          category: 'Warranty - Refrigeration Circuit',
          issue: 'Evaporator-coil leak appeared buried within the fin-and-tube section.',
          action: 'Leak location and repair path remained under investigation; evidence should stay tied to the unit and coil decision.',
          why: 'High-impact refrigeration issue with warranty exposure and evidence requirements.'
        }),
        Object.freeze({
          id: '47833398197',
          product: 'ARU',
          title: 'Faulty A2L sensor - Niagara Chester',
          category: 'Warranty - Electrical / Controls',
          issue: 'Bad A2L sensor reported on the Niagara Chester ARU. Field needed replacement support and interim guidance.',
          action: 'Replacement sensor was requested/sent, and the field was advised on the temporary jumper bypass between 82 and 84 while awaiting repair.',
          why: 'Safety-related device failure on R-454B equipment; needs clear replacement and returned-part follow-up.'
        })
      ])
    })
  })
});

function prepareMonthlyQualityNarrativeContext_(spreadsheet, issueResult) {
  const reportMonth = String(issueResult.reportMonth || '').trim();
  if (!reportMonth) {
    throw new Error('Monthly top-ticket context cannot determine report month.');
  }

  const override = MONTHLY_QUALITY_TOP_TICKET_CONTEXT_[reportMonth];
  if (!override) {
    throw new Error('No Top 3 ticket context is configured for ' + reportMonth + '.');
  }

  const monthLabel = monthlyQualityNarrativeMonthLabel_(reportMonth);
  const sheet = ensureMonthlyQualityNarrativeContextSheet_(spreadsheet);

  sheet.clear();
  sheet.getRange(1, 1, 1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length)
    .setValues([MONTHLY_QUALITY_CONTEXT_HEADERS_]);
  sheet.getRange(1, 1, 1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const rows = MONTHLY_QUALITY_CONTEXT_SECTIONS_.map(function(section) {
    const item = override[section];
    const tickets = item.tickets || [];
    return [
      reportMonth,
      section,
      item.headline || '',
      tickets[0] ? formatMonthlyQualityTopTicketTitle_(tickets[0]) : '',
      tickets[0] ? formatMonthlyQualityTopTicketBody_(tickets[0]) : '',
      tickets[1] ? formatMonthlyQualityTopTicketTitle_(tickets[1]) : '',
      tickets[1] ? formatMonthlyQualityTopTicketBody_(tickets[1]) : '',
      tickets[2] ? formatMonthlyQualityTopTicketTitle_(tickets[2]) : '',
      tickets[2] ? formatMonthlyQualityTopTicketBody_(tickets[2]) : '',
      'Use the ticket examples to explain the metric in plain language. Do not present questions or documentation requests as failures.',
      'Show the ticket, what happened, what action/status exists, and why it matters to the floor.'
    ];
  });

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, rows.length + 1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length).setWrap(true);
  sheet.autoResizeColumns(1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length);
  SpreadsheetApp.flush();

  return {
    reportMonth: reportMonth,
    monthLabel: monthLabel,
    source: MONTHLY_QUALITY_CONTEXT_SHEET_NAME_,
    sections: readMonthlyQualityNarrativeContextRows_(sheet, reportMonth)
  };
}

function formatMonthlyQualityTopTicketTitle_(ticket) {
  return 'Ticket ' + ticket.id + ' - ' + ticket.product + ' - ' + ticket.title;
}

function formatMonthlyQualityTopTicketBody_(ticket) {
  return [
    'Category: ' + ticket.category,
    'What happened: ' + ticket.issue,
    'Fix / status: ' + ticket.action,
    'Why it matters: ' + ticket.why
  ].join('\n');
}

function drawMonthlyQualityNarrativeSlide_(slide, section, record, monthLabel) {
  const pageWidth = 720;
  const pageHeight = 405;
  const margin = 26;
  const titleHeight = 32;
  const top = 70;
  const ticketHeight = 86;
  const ticketGap = 10;
  const footerHeight = 14;
  const displaySection = section === 'All Lines' ? 'Plant Summary' : section;

  slide.getBackground().setSolidFill('#FFFFFF');

  const title = slide.insertTextBox(
    'Top 3 Tickets - ' + displaySection,
    margin,
    20,
    pageWidth - margin * 2,
    titleHeight
  );
  title.setTitle('Top 3 Tickets - ' + displaySection);
  title.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(24)
    .setBold(true)
    .setForegroundColor('#333333');

  const subtitle = slide.insertTextBox(
    monthLabel + ' - actual tickets behind the quality numbers',
    margin,
    50,
    pageWidth - margin * 2,
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
    pageWidth - margin * 2,
    22
  );
  headline.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(11)
    .setBold(true)
    .setForegroundColor('#222222');

  for (let index = 0; index < 3; index++) {
    const issue = record.issues[index] || { title: '', why: '' };
    const y = top + 28 + index * (ticketHeight + ticketGap);
    drawMonthlyQualityTopTicketBox_(
      slide,
      margin,
      y,
      pageWidth - margin * 2,
      ticketHeight,
      index + 1,
      issue.title,
      issue.why
    );
  }

  const footer = slide.insertTextBox(
    'Source: audited HubSpot support tickets. Questions and documentation-only requests are excluded from failure counts.',
    margin,
    pageHeight - margin - footerHeight + 2,
    pageWidth - margin * 2,
    footerHeight
  );
  footer.getText().getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(7)
    .setForegroundColor('#777777');
}

function drawMonthlyQualityTopTicketBox_(slide, x, y, w, h, rank, title, body) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  shape.getFill().setSolidFill(rank === 1 ? '#FFF2CC' : '#EEF3FA');
  shape.getBorder().getLineFill().setSolidFill(rank === 1 ? '#D6B656' : '#9FBAD7');
  shape.setTitle('Monthly Quality Top Ticket ' + rank);

  const text = '#' + rank + '  ' + title + '\n' + body;
  const textRange = shape.getText();
  textRange.setText(text);
  textRange.getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(8)
    .setForegroundColor('#222222');

  try {
    textRange.getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  } catch (error) {}

  try {
    textRange.getRange(0, String('#' + rank + '  ' + title).length)
      .getTextStyle()
      .setBold(true)
      .setFontSize(10);
  } catch (error) {}

  return shape;
}
