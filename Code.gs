function calculateFPYSummary_FINAL() {
  const token = 'scapi_Rr6jleu8o5EgnXnDi_UriXsH-xta_zGELOJtc0ObUdFweWlrtvJCm1-DCIPLyNZxGLn1CREeZGLu3IybDIX-VpP3o-QmUOQewXhkL3hq8QBLAGnRK7bnTts1_odUZ0HZELTJZlGA1au36uGQ-85dK_V17Jxpayn6g85aJHCgdgY';
  const BASE = 'https://api.safetyculture.io';

  const TEMPLATE_MAP = {
    'template_9f49d4f7e3924b9fa36bcc249f5ea96a': 'ARU',
    'template_95a16e28e5184839899cf3dfb6dbf286': 'CSC',
    'template_db8cb7b6b670439088dfa3f780d020d4': 'MSC',
    'template_eeefdf55f60440e583f20e91db821b8d': 'HGRH',
    'template_2cb7229367e84471a048e5d05a54180a': 'Gas Heat',
    'template_4d2ee8c207e64f94aa6c7627980a6eea': 'Bard Coatings',
    'template_9ac31eb9905248f68ca78a069ca23f79': 'Coatings'
  };

  const PRODUCTS = [
    'ARU',
    'CSC',
    'HGRH',
    'MSC',
    'Gas Heat',
    'Bard Coatings',
    'Coatings'
  ];

  const START_COLS = {
    ARU: 2,
    CSC: 5,
    HGRH: 8,
    MSC: 11,
    'Gas Heat': 14,
    'Bard Coatings': 17,
    Coatings: 20
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (!sheet) throw new Error('Sheet1 not found.');

  sheet.clear();

  // Layout
  sheet.getRange('A1').setValue('Month/Year');
  sheet.getRange('A2').setValue('Plant Average');
  sheet.getRange('A4').setValue('Current Year Total');

  PRODUCTS.forEach(function (product) {
    const startCol = START_COLS[product];
    sheet.getRange(1, startCol, 2, 3).merge().setValue(product);
  });
  sheet.getRange('W1:Y2').merge().setValue('All Lines');

  // Metric headers on row 4
  PRODUCTS.forEach(function (product) {
    const c = START_COLS[product];
    sheet.getRange(4, c).setValue('Total Inspected');
    sheet.getRange(4, c + 1).setValue('Defects');
    sheet.getRange(4, c + 2).setValue('FPY Inspection');
  });
  sheet.getRange('W4').setValue('Total Inspected');
  sheet.getRange('X4').setValue('Defects');
  sheet.getRange('Y4').setValue('FPY Inspection');

  let url = BASE + '/feed/inspections?modified_after=2026-01-01T00:00:00Z&limit=100';

  const monthlyData = {};
  const yearlyTotals = {};
  const monthKeys = [];

  let plantTotal = 0;
  let plantDefects = 0;

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

      plantTotal++;
      if (defectFound) plantDefects++;

      Utilities.sleep(100);
    }

    url = json.metadata && json.metadata.next_page
      ? BASE + json.metadata.next_page
      : null;
  }

  monthKeys.sort();

  const allLinesMonthly = {};

  monthKeys.forEach(function (mk) {
    let total = 0;
    let defects = 0;

    PRODUCTS.forEach(function (product) {
      const record = monthlyData[product + '|' + mk];
      if (record) {
        total += record.total;
        defects += record.defects;
      }
    });

    allLinesMonthly[mk] = { total: total, defects: defects };
  });

  // Overall plant average in A3 only
  const plantFPY = plantTotal > 0 ? (plantTotal - plantDefects) / plantTotal : 0;
  sheet.getRange('A3').setValue(plantFPY);
  sheet.getRange('W5').setValue(plantTotal);
  sheet.getRange('X5').setValue(plantDefects);
  sheet.getRange('Y5').setValue(plantFPY);

  // Current Year Total row
  PRODUCTS.forEach(function (product) {
    const yearlyTotal = yearlyTotals[product].total;
    const yearlyDefects = yearlyTotals[product].defects;
    const yearlyFPY = yearlyTotal > 0 ? (yearlyTotal - yearlyDefects) / yearlyTotal : 0;

    const col = START_COLS[product];

    sheet.getRange(5, col).setValue(yearlyTotal);
    sheet.getRange(5, col + 1).setValue(yearlyDefects);
    sheet.getRange(5, col + 2).setValue(yearlyFPY);
  });

  const allLinesOut = monthKeys.map(function (mk) {
    const record = allLinesMonthly[mk] || { total: 0, defects: 0 };
    const fpy = record.total > 0 ? (record.total - record.defects) / record.total : 0;
    return [record.total, record.defects, fpy];
  });

  if (allLinesOut.length) {
    sheet.getRange(6, 23, allLinesOut.length, 3).setValues(allLinesOut); // W:Y
  }

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
      sheet.getRange(6, START_COLS[product], out.length, 3).setValues(out);
    }
  });

  const monthlyRowCount = monthKeys.length;
  const lastDataRow = Math.max(6, monthlyRowCount + 5);
  const lastCol = 25; // Y

  // Formatting
  sheet.getRange('A3').setNumberFormat('0.00%');

  PRODUCTS.forEach(function (product) {
    const c = START_COLS[product];

    // Current year totals
    sheet.getRange(5, c).setNumberFormat('0');
    sheet.getRange(5, c + 1).setNumberFormat('0');
    sheet.getRange(5, c + 2).setNumberFormat('0.00%');

    if (monthlyRowCount > 0) {
      sheet.getRange(6, c, monthlyRowCount, 2).setNumberFormat('0');
      sheet.getRange(6, c + 2, monthlyRowCount, 1).setNumberFormat('0.00%');
    }
  });

  sheet.getRange('W5').setNumberFormat('0');
  sheet.getRange('X5').setNumberFormat('0');
  sheet.getRange('Y5').setNumberFormat('0.00%');

  if (monthlyRowCount > 0) {
    sheet.getRange(6, 23, monthlyRowCount, 2).setNumberFormat('0');
    sheet.getRange(6, 25, monthlyRowCount, 1).setNumberFormat('0.00%');
  }

  // Borders and styling
  sheet.getRange(1, 1, lastDataRow, lastCol).setBorder(
    true, true, true, true, true, true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.getRange(1, 1, 5, lastCol).setFontWeight('bold').setHorizontalAlignment('center');

  PRODUCTS.forEach(function (product) {
    const c = START_COLS[product];
    sheet.getRange(1, c, 2, 3).setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  sheet.getRange('A3:A' + lastDataRow).setHorizontalAlignment('left');

  // Auto resize + explicit widths
  for (let col = 1; col <= lastCol; col++) {
    sheet.autoResizeColumn(col);
  }

  sheet.setColumnWidth(1, 140);

  PRODUCTS.forEach(function (product) {
    const c = START_COLS[product];
    sheet.setColumnWidth(c, 95);
    sheet.setColumnWidth(c + 1, 80);
    sheet.setColumnWidth(c + 2, 105);
  });
  sheet.setColumnWidth(23, 95);
  sheet.setColumnWidth(24, 80);
  sheet.setColumnWidth(25, 105);

  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 28);
  sheet.setRowHeight(4, 28);
  sheet.setRowHeight(5, 28);
}
