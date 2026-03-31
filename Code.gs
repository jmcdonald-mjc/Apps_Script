function calculateFPYSummary_FINAL() {
  const token = 'scapi_Rr6jleu8o5EgnXnDi_UriXsH-xta_zGELOJtc0ObUdFweWlrtvJCm1-DCIPLyNZxGLn1CREeZGLu3IybDIX-VpP3o-QmUOQewXhkL3hq8QBLAGnRK7bnTts1_odUZ0HZELTJZlGA1au36uGQ-85dK_V17Jxpayn6g85aJHCgdgY';
  const BASE = 'https://api.safetyculture.io';

  const TEMPLATE_MAP = {
    'template_9f49d4f7e3924b9fa36bcc249f5ea96a': 'ARU',
    'template_95a16e28e5184839899cf3dfb6dbf286': 'CSC',
    'template_db8cb7b6b670439088dfa3f780d020d4': 'MSC',
    'template_eeefdf55f60440e583f20e91db821b8d': 'HGRH'
  };

  const PRODUCTS = ['ARU', 'CSC', 'HGRH', 'MSC'];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) throw new Error('Sheet1 not found.');

  sheet.clear();

  // Layout
  sheet.getRange('A1').setValue('Month/Year');
  sheet.getRange('A2').setValue('Plant Average');
  sheet.getRange('A4').setValue('Current Year Total');

  sheet.getRange('B2:D2').merge().setValue('ARU');
  sheet.getRange('E2:G2').merge().setValue('CSC');
  sheet.getRange('H2:J2').merge().setValue('HGRH');
  sheet.getRange('K2:M2').merge().setValue('MSC');

  sheet.getRange('B3').setValue('Total Inspected');
  sheet.getRange('C3').setValue('Defects');
  sheet.getRange('D3').setValue('FPY Inspection');

  sheet.getRange('E3').setValue('Total Inspected');
  sheet.getRange('F3').setValue('Defects');
  sheet.getRange('G3').setValue('FPY Inspection');

  sheet.getRange('H3').setValue('Total Inspected');
  sheet.getRange('I3').setValue('Defects');
  sheet.getRange('J3').setValue('FPY Inspection');

  sheet.getRange('K3').setValue('Total Inspected');
  sheet.getRange('L3').setValue('Defects');
  sheet.getRange('M3').setValue('FPY Inspection');

  let url = BASE + '/feed/inspections?modified_after=2026-01-01T00:00:00Z&limit=100';

  const data = {};
  const yearTotals = {};
  const monthKeys = [];

  let totalProcessed = 0;
  let matchedTemplate = 0;
  let missingInspectionId = 0;
  let detailFailures = 0;

  PRODUCTS.forEach(function (product) {
    yearTotals[product] = { total: 0, defects: 0 };
  });

  function extractAnswerFromResponses(responses) {
    if (!responses) return '';

    if (
      responses.selected &&
      Array.isArray(responses.selected) &&
      responses.selected.length > 0
    ) {
      const firstSelected = responses.selected[0];
      if (firstSelected && firstSelected.label != null) {
        return String(firstSelected.label).trim();
      }
      if (firstSelected && firstSelected.value != null) {
        return String(firstSelected.value).trim();
      }
    }

    if (Array.isArray(responses) && responses.length > 0) {
      const first = responses[0];

      if (
        first &&
        first.selected &&
        Array.isArray(first.selected) &&
        first.selected.length > 0
      ) {
        const nestedSelected = first.selected[0];
        if (nestedSelected && nestedSelected.label != null) {
          return String(nestedSelected.label).trim();
        }
        if (nestedSelected && nestedSelected.value != null) {
          return String(nestedSelected.value).trim();
        }
      }

      if (first && first.label != null) return String(first.label).trim();
      if (first && first.value != null) return String(first.value).trim();
    }

    if (responses.label != null) return String(responses.label).trim();
    if (responses.value != null) return String(responses.value).trim();

    return '';
  }

  function findDefectAnswer(items) {
    if (!items || !Array.isArray(items)) return '';

    for (const item of items) {
      if (!item) continue;

      const label = String(item.label || '').trim().toLowerCase();

      if (
        label === 'were there any defects found?' ||
        label.includes('were there any defects found')
      ) {
        const answer = extractAnswerFromResponses(item.responses);
        if (answer) return answer;
      }

      if (item.items) {
        const nestedAnswer = findDefectAnswer(item.items);
        if (nestedAnswer) return nestedAnswer;
      }
    }

    return '';
  }

  function fetchInspectionDetail(id) {
    const candidates = [
      BASE + '/audits/' + encodeURIComponent(id),
      BASE + '/inspections/v1/inspections/' + encodeURIComponent(id)
    ];

    for (const endpoint of candidates) {
      const res = UrlFetchApp.fetch(endpoint, {
        method: 'get',
        headers: {
          Authorization: 'Bearer ' + token,
          Accept: 'application/json'
        },
        muteHttpExceptions: true
      });

      if (res.getResponseCode() === 200) {
        return JSON.parse(res.getContentText());
      }
    }

    return null;
  }

  while (url) {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code !== 200) {
      throw new Error('Feed search failed: ' + code + ' ' + text);
    }

    const json = JSON.parse(text);
    const inspections = json.data || [];

    if (!inspections.length) break;

    for (const inspection of inspections) {
      totalProcessed++;

      const productLine = TEMPLATE_MAP[inspection.template_id];
      if (!productLine) continue;
      matchedTemplate++;

      const inspectionId = inspection.inspection_id || inspection.audit_id || inspection.id;
      if (!inspectionId) {
        missingInspectionId++;
        continue;
      }

      const dateStr =
        inspection.date_completed ||
        inspection.modified_at ||
        inspection.created_at;

      if (!dateStr) continue;

      const d = new Date(dateStr);
      if (isNaN(d.getTime())) continue;
      if (d.getFullYear() < 2026) continue;

      const year = String(d.getFullYear());
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const monthKey = year + '-' + month;

      if (!monthKeys.includes(monthKey)) monthKeys.push(monthKey);

      const detail = fetchInspectionDetail(inspectionId);
      if (!detail) {
        detailFailures++;
        continue;
      }

      const defectAnswer =
        findDefectAnswer(detail.header_items) ||
        findDefectAnswer(detail.items) ||
        findDefectAnswer(detail.audit_data && detail.audit_data.items);

      const defectFound = String(defectAnswer).trim().toLowerCase() === 'yes';

      const key = productLine + '|' + monthKey;

      if (!data[key]) {
        data[key] = { total: 0, defects: 0 };
      }

      data[key].total++;
      if (defectFound) data[key].defects++;

      yearTotals[productLine].total++;
      if (defectFound) yearTotals[productLine].defects++;

      Utilities.sleep(100);
    }

    url = json.metadata && json.metadata.next_page
      ? BASE + json.metadata.next_page
      : null;
  }

  monthKeys.sort();

  // Plant average / yearly values
  PRODUCTS.forEach(function (product) {
    const yearlyTotal = yearTotals[product].total;
    const yearlyDefects = yearTotals[product].defects;
    const yearlyFPY = yearlyTotal > 0 ? (yearlyTotal - yearlyDefects) / yearlyTotal : 0;

    const startCol = {
      ARU: 2,
      CSC: 5,
      HGRH: 8,
      MSC: 11
    }[product];

    // Row 4 = current year totals
    sheet.getRange(4, startCol).setValue(yearlyTotal);
    sheet.getRange(4, startCol + 1).setValue(yearlyDefects);
    sheet.getRange(4, startCol + 2).setValue(yearlyFPY);

    // Row 2 = plant average / yearly FPY
    sheet.getRange(2, startCol).setValue(yearlyFPY);
  });

  // Monthly rows start at row 5
  const monthLabelValues = monthKeys.map(function (mk) {
    const parts = mk.split('-');
    return [parts[1] + '/' + parts[0].slice(2)];
  });

  if (monthLabelValues.length) {
    sheet.getRange(5, 1, monthLabelValues.length, 1).setValues(monthLabelValues);
  }

  const startCols = {
    ARU: 2,
    CSC: 5,
    HGRH: 8,
    MSC: 11
  };

  PRODUCTS.forEach(function (product) {
    const out = monthKeys.map(function (mk) {
      const record = data[product + '|' + mk] || { total: 0, defects: 0 };
      const fpy = record.total > 0 ? (record.total - record.defects) / record.total : 0;
      return [record.total, record.defects, fpy];
    });

    if (out.length) {
      sheet.getRange(5, startCols[product], out.length, 3).setValues(out);
    }
  });

  const monthlyRowCount = monthKeys.length;
  const lastDataRow = Math.max(5, monthlyRowCount + 4);

  // Number formatting
  if (monthlyRowCount > 0) {
    sheet.getRange(5, 2, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(5, 5, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(5, 8, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(5, 11, monthlyRowCount, 2).setNumberFormat('0');

    sheet.getRange(5, 4, monthlyRowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(5, 7, monthlyRowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(5, 10, monthlyRowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(5, 13, monthlyRowCount, 1).setNumberFormat('0.00%');
  }

  sheet.getRange('B4').setNumberFormat('0');
  sheet.getRange('C4').setNumberFormat('0');
  sheet.getRange('D4').setNumberFormat('0.00%');

  sheet.getRange('E4').setNumberFormat('0');
  sheet.getRange('F4').setNumberFormat('0');
  sheet.getRange('G4').setNumberFormat('0.00%');

  sheet.getRange('H4').setNumberFormat('0');
  sheet.getRange('I4').setNumberFormat('0');
  sheet.getRange('J4').setNumberFormat('0.00%');

  sheet.getRange('K4').setNumberFormat('0');
  sheet.getRange('L4').setNumberFormat('0');
  sheet.getRange('M4').setNumberFormat('0.00%');

  sheet.getRange('B2').setNumberFormat('0.00%');
  sheet.getRange('E2').setNumberFormat('0.00%');
  sheet.getRange('H2').setNumberFormat('0.00%');
  sheet.getRange('K2').setNumberFormat('0.00%');

  // Formatting and black borders
  sheet.getRange('A1:M' + lastDataRow).setBorder(
    true, true, true, true, true, true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.getRange('A1:M3').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B2:D2').setHorizontalAlignment('center');
  sheet.getRange('E2:G2').setHorizontalAlignment('center');
  sheet.getRange('H2:J2').setHorizontalAlignment('center');
  sheet.getRange('K2:M2').setHorizontalAlignment('center');
  sheet.getRange('A4:A' + lastDataRow).setHorizontalAlignment('right');

  // Auto resize columns A:M
  for (let col = 1; col <= 13; col++) {
    sheet.autoResizeColumn(col);
  }

  // Keep some columns from getting too narrow
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 120));
  sheet.setColumnWidth(2, Math.max(sheet.getColumnWidth(2), 110));
  sheet.setColumnWidth(5, Math.max(sheet.getColumnWidth(5), 110));
  sheet.setColumnWidth(8, Math.max(sheet.getColumnWidth(8), 110));
  sheet.setColumnWidth(11, Math.max(sheet.getColumnWidth(11), 110));

  // Row heights
  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 28);
  sheet.setFrozenRows(4);

  // Debug block
  const debugStartRow = Math.max(9, lastDataRow + 2);
  const debugOutput = [
    ['DEBUG totalProcessed', totalProcessed],
    ['DEBUG matchedTemplate', matchedTemplate],
    ['DEBUG missingInspectionId', missingInspectionId],
    ['DEBUG detailFailures', detailFailures]
  ];
  sheet.getRange(debugStartRow, 1, debugOutput.length, 2).setValues(debugOutput);
}
