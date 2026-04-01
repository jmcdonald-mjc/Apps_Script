function calculateFPYSummary_FINAL() {
  const token = 'scapi_Rr6jleu8o5EgnXnDi_UriXsH-xta_zGELOJtc0ObUdFweWlrtvJCm1-DCIPLyNZxGLn1CREeZGLu3IybDIX-VpP3o-QmUOQewXhkL3hq8QBLAGnRK7bnTts1_odUZ0HZELTJZlGA1au36uGQ-85dK_V17Jxpayn6g85aJHCgdgY';
  const BASE = 'https://api.safetyculture.io';

  const TEMPLATE_MAP = {
    'template_9f49d4f7e3924b9fa36bcc249f5ea96a': 'ARU',
    'template_95a16e28e5184839899cf3dfb6dbf286': 'CSC',
    'template_4d2ee8c207e64f94aa6c7627980a6eea': 'CSC', // Bard Coatings
    'template_db8cb7b6b670439088dfa3f780d020d4': 'MSC',
    'template_9ac31eb9905248f68ca78a069ca23f79': 'MSC', // Coatings
    'template_eeefdf55f60440e583f20e91db821b8d': 'HGRH',
    'template_2cb7229367e84471a048e5d05a54180a': 'HGRH' // Gas Heat
  };

  const PRODUCTS = ['ARU', 'CSC', 'HGRH', 'MSC'];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) throw new Error('Sheet1 not found.');

  sheet.clear();

  // Layout to match sheet
  sheet.getRange('A1').setValue('Month/Year');
  sheet.getRange('A3').setValue('Plant Average');
  sheet.getRange('A5').setValue('Current Year Total');

  sheet.getRange('B1:D2').merge().setValue('ARU');
  sheet.getRange('E1:G2').merge().setValue('CSC');
  sheet.getRange('H1:J2').merge().setValue('HGRH');
  sheet.getRange('K1:M2').merge().setValue('MSC');

  // Metric headers now on row 4
  sheet.getRange('B4').setValue('Total Inspected');
  sheet.getRange('C4').setValue('Defects');
  sheet.getRange('D4').setValue('FPY Inspection');

  sheet.getRange('E4').setValue('Total Inspected');
  sheet.getRange('F4').setValue('Defects');
  sheet.getRange('G4').setValue('FPY Inspection');

  sheet.getRange('H4').setValue('Total Inspected');
  sheet.getRange('I4').setValue('Defects');
  sheet.getRange('J4').setValue('FPY Inspection');

  sheet.getRange('K4').setValue('Total Inspected');
  sheet.getRange('L4').setValue('Defects');
  sheet.getRange('M4').setValue('FPY Inspection');

  let url = BASE + '/feed/inspections?modified_after=2026-01-01T00:00:00Z&limit=100';

  const monthlyData = {};
  const yearlyTotals = {};
  const monthKeys = [];

  PRODUCTS.forEach(function (product) {
    yearlyTotals[product] = { total: 0, defects: 0 };
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
      const productLine = TEMPLATE_MAP[inspection.template_id];
      if (!productLine) continue;

      const inspectionId = inspection.inspection_id || inspection.audit_id || inspection.id;
      if (!inspectionId) continue;

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
      if (!detail) continue;

      const defectAnswer =
        findDefectAnswer(detail.header_items) ||
        findDefectAnswer(detail.items) ||
        findDefectAnswer(detail.audit_data && detail.audit_data.items);

      const defectFound = String(defectAnswer).trim().toLowerCase() === 'yes';

      const key = productLine + '|' + monthKey;

      if (!monthlyData[key]) {
        monthlyData[key] = { total: 0, defects: 0 };
      }

      monthlyData[key].total++;
      if (defectFound) monthlyData[key].defects++;

      yearlyTotals[productLine].total++;
      if (defectFound) yearlyTotals[productLine].defects++;

      Utilities.sleep(100);
    }

    url = json.metadata && json.metadata.next_page
      ? BASE + json.metadata.next_page
      : null;
  }

  monthKeys.sort();

  const startCols = {
    ARU: 2,
    CSC: 5,
    HGRH: 8,
    MSC: 11
  };

  // Row 3 = Plant Average values
  PRODUCTS.forEach(function (product) {
    const yearlyTotal = yearlyTotals[product].total;
    const yearlyDefects = yearlyTotals[product].defects;
    const yearlyFPY = yearlyTotal > 0 ? (yearlyTotal - yearlyDefects) / yearlyTotal : 0;

    const col = startCols[product];

    // Plant Average row
    sheet.getRange(3, col).setValue(yearlyFPY);

    // Current Year Total row
    sheet.getRange(5, col).setValue(yearlyTotal);
    sheet.getRange(5, col + 1).setValue(yearlyDefects);
    sheet.getRange(5, col + 2).setValue(yearlyFPY);
  });

  // Monthly labels start row 6
  const monthLabelValues = monthKeys.map(function (mk) {
    const parts = mk.split('-');
    return [parts[1] + '/' + parts[0].slice(2)];
  });

  if (monthLabelValues.length) {
    sheet.getRange(6, 1, monthLabelValues.length, 1).setValues(monthLabelValues);
  }

  // Monthly data start row 6
  PRODUCTS.forEach(function (product) {
    const out = monthKeys.map(function (mk) {
      const record = monthlyData[product + '|' + mk] || { total: 0, defects: 0 };
      const fpy = record.total > 0 ? (record.total - record.defects) / record.total : 0;
      return [record.total, record.defects, fpy];
    });

    if (out.length) {
      sheet.getRange(6, startCols[product], out.length, 3).setValues(out);
    }
  });

  const monthlyRowCount = monthKeys.length;
  const lastDataRow = Math.max(6, monthlyRowCount + 5);

  // Formatting
  if (monthlyRowCount > 0) {
    // Monthly counts
    sheet.getRange(6, 2, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(6, 5, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(6, 8, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(6, 11, monthlyRowCount, 2).setNumberFormat('0');

    // Monthly FPY
    sheet.getRange(6, 4, monthlyRowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(6, 7, monthlyRowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(6, 10, monthlyRowCount, 1).setNumberFormat('0.00%');
    sheet.getRange(6, 13, monthlyRowCount, 1).setNumberFormat('0.00%');
  }

  // Plant Average row
  sheet.getRange('B3').setNumberFormat('0.00%');
  sheet.getRange('E3').setNumberFormat('0.00%');
  sheet.getRange('H3').setNumberFormat('0.00%');
  sheet.getRange('K3').setNumberFormat('0.00%');

  // Current Year Total row
  sheet.getRange('B5').setNumberFormat('0');
  sheet.getRange('C5').setNumberFormat('0');
  sheet.getRange('D5').setNumberFormat('0.00%');

  sheet.getRange('E5').setNumberFormat('0');
  sheet.getRange('F5').setNumberFormat('0');
  sheet.getRange('G5').setNumberFormat('0.00%');

  sheet.getRange('H5').setNumberFormat('0');
  sheet.getRange('I5').setNumberFormat('0');
  sheet.getRange('J5').setNumberFormat('0.00%');

  sheet.getRange('K5').setNumberFormat('0');
  sheet.getRange('L5').setNumberFormat('0');
  sheet.getRange('M5').setNumberFormat('0.00%');

  // Borders and styling
  sheet.getRange('A1:M' + lastDataRow).setBorder(
    true, true, true, true, true, true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.getRange('A1:M5').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B1:D2').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('E1:G2').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('H1:J2').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('K1:M2').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('A3:A' + lastDataRow).setHorizontalAlignment('left');

  // Auto resize + explicit widths
  for (let col = 1; col <= 13; col++) {
    sheet.autoResizeColumn(col);
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 95);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 105);
  sheet.setColumnWidth(5, 95);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 105);
  sheet.setColumnWidth(8, 95);
  sheet.setColumnWidth(9, 80);
  sheet.setColumnWidth(10, 105);
  sheet.setColumnWidth(11, 95);
  sheet.setColumnWidth(12, 80);
  sheet.setColumnWidth(13, 105);

  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 28);
  sheet.setRowHeight(4, 28);
  sheet.setRowHeight(5, 28);
}
