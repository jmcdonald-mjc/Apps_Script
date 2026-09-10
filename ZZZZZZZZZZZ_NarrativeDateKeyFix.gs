/**
 * Fix narrative context month matching for Google Sheets date values.
 *
 * The Monthly Report Context sheet stores the report month as 2026-08 visually,
 * but Apps Script may read that cell as a Date object instead of the string
 * "2026-08". The previous reader compared String(row[0]) directly to the
 * reportMonth key, so valid rows could be missed and the Slides context step
 * would stop with "no context rows were available".
 *
 * This file intentionally sorts after the narrative files and replaces only the
 * month-sensitive reader behavior.
 */

function monthlyQualityNarrativeNormalizeMonthKey_(value) {
  if (value === null || value === undefined || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || 'America/New_York',
      'yyyy-MM'
    );
  }

  const text = String(value || '').trim();
  const direct = /^(\d{4})-(\d{2})$/.exec(text);
  if (direct) return direct[1] + '-' + direct[2];

  const slash = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (slash) {
    const month = Number(slash[1]);
    const year = Number(slash[3]);
    if (year && month >= 1 && month <= 12) {
      return year + '-' + String(month).padStart(2, '0');
    }
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(
      parsed,
      Session.getScriptTimeZone() || 'America/New_York',
      'yyyy-MM'
    );
  }

  return text;
}

function readMonthlyQualityNarrativeContextRows_(sheet, reportMonth) {
  const targetMonth = monthlyQualityNarrativeNormalizeMonthKey_(reportMonth);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet
    .getRange(2, 1, lastRow - 1, MONTHLY_QUALITY_CONTEXT_HEADERS_.length)
    .getValues();
  const rows = [];

  values.forEach(function(row) {
    const month = monthlyQualityNarrativeNormalizeMonthKey_(row[0]);
    const section = String(row[1] || '').trim();
    if (month !== targetMonth || MONTHLY_QUALITY_CONTEXT_SECTIONS_.indexOf(section) < 0) {
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

  Logger.log(JSON.stringify({
    status: 'NARRATIVE_CONTEXT_ROWS_READ',
    reportMonth: targetMonth,
    rowCount: rows.length
  }));

  return rows;
}
