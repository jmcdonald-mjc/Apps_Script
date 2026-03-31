function calculateFPYSummary_FINAL() {
  const token = 'scapi_Rr6jleu8o5EgnXnDi_UriXsH-xta_zGELOJtc0ObUdFweWlrtvJCm1-DCIPLyNZxGLn1CREeZGLu3IybDIX-VpP3o-QmUOQewXhkL3hq8QBLAGnRK7bnTts1_odUZ0HZELTJZlGA1au36uGQ-85dK_V17Jxpayn6g85aJHCgdgY';
  const BASE = 'https://api.safetyculture.io';
  const TEMPLATE_MAP = {
    template_9f49d4f7e3924b9fa36bcc249f5ea96a: 'ARU',
    template_95a16e28e5184839899cf3dfb6dbf286: 'CSC',
    template_db8cb7b6b670439088dfa3f780d020d4: 'MSC',
    template_eeefdf55f60440e583f20e91db821b8d: 'HGRH'
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();

  const output = [];
  output.push(['Product Line', 'Year', 'Month', 'Total', 'Defects', 'FPY']);

  let url = BASE + '/feed/inspections?modified_after=2026-01-01T00:00:00Z&limit=100';

  const data = {};
  let totalProcessed = 0;
  let matchedTemplate = 0;
  let missingInspectionId = 0;
  let detailFailures = 0;

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

      const year = String(d.getFullYear());
      const month = String(d.getMonth() + 1);
      const key = productLine + '|' + year + '|' + month;

      if (!data[key]) {
        data[key] = {
          productLine: productLine,
          year: year,
          month: month,
          total: 0,
          defects: 0
        };
      }

      data[key].total++;
      if (defectFound) data[key].defects++;

      Utilities.sleep(100);
    }

    url = json.metadata && json.metadata.next_page
      ? BASE + json.metadata.next_page
      : null;
  }

  let grandTotal = 0;
  let grandDefects = 0;

  Object.keys(data).sort().forEach(function (key) {
    const row = data[key];
    const fpy = row.total > 0 ? (row.total - row.defects) / row.total : 0;

    output.push([
      row.productLine,
      row.year,
      row.month,
      row.total,
      row.defects,
      fpy
    ]);

    grandTotal += row.total;
    grandDefects += row.defects;
  });

  output.push([
    'TOTAL',
    '',
    '',
    grandTotal,
    grandDefects,
    grandTotal > 0 ? (grandTotal - grandDefects) / grandTotal : 0
  ]);

  output.push(['DEBUG totalProcessed', totalProcessed, '', '', '', '']);
  output.push(['DEBUG matchedTemplate', matchedTemplate, '', '', '', '']);
  output.push(['DEBUG missingInspectionId', missingInspectionId, '', '', '', '']);
  output.push(['DEBUG detailFailures', detailFailures, '', '', '', '']);

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  // Format FPY column as percentage
  sheet.getRange(2, 6, output.length - 1, 1).setNumberFormat('0.00%');

  // Write the dashboard matrix layout on the FPY Report tab.
  writeFPYMatrixLayout();
}

function writeFPYMatrixLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Sheet1');
  const targetSheet = ss.getSheetByName('FPY Report');

  if (!sourceSheet) throw new Error('Sheet1 not found.');
  if (!targetSheet) throw new Error('FPY Report not found.');

  const sourceData = sourceSheet.getDataRange().getValues();
  if (sourceData.length < 2) throw new Error('No source data found on Sheet1.');

  const headers = sourceData[0];
  const rows = sourceData.slice(1);

  const colIndex = {
    product: headers.indexOf('Product Line'),
    year: headers.indexOf('Year'),
    month: headers.indexOf('Month'),
    total: headers.indexOf('Total'),
    defects: headers.indexOf('Defects'),
    fpy: headers.indexOf('FPY')
  };

  for (const [key, idx] of Object.entries(colIndex)) {
    if (idx === -1) throw new Error('Missing column in Sheet1: ' + key);
  }

  const products = ['ARU', 'CSC', 'HGRH', 'MSC'];
  const monthKeys = [];
  const dataMap = {};

  rows.forEach(function (r) {
    const product = r[colIndex.product];
    const year = r[colIndex.year];
    const month = r[colIndex.month];

    if (!products.includes(product)) return;
    if (!year || month === '' || month == null) return;

    const monthNum = Number(month);
    const monthKey = year + '-' + String(monthNum).padStart(2, '0');

    if (!monthKeys.includes(monthKey)) monthKeys.push(monthKey);

    dataMap[product + '|' + monthKey] = {
      total: Number(r[colIndex.total]) || 0,
      defects: Number(r[colIndex.defects]) || 0,
      fpy: Number(r[colIndex.fpy]) || 0
    };
  });

  monthKeys.sort();

  // Keep report headers; clear only the body area.
  targetSheet.getRange('A4:M1000').clearContent();

  const monthLabelValues = monthKeys.map(function (monthKey) {
    const parts = monthKey.split('-');
    const year = parts[0];
    const month = parts[1];
    return [month + '/' + year.slice(2)];
  });

  if (monthLabelValues.length) {
    targetSheet.getRange(4, 1, monthLabelValues.length, 1).setValues(monthLabelValues);
  }

  const startCols = {
    ARU: 2,
    CSC: 5,
    HGRH: 8,
    MSC: 11
  };

  products.forEach(function (product) {
    const out = monthKeys.map(function (monthKey) {
      const record = dataMap[product + '|' + monthKey] || { total: 0, defects: 0, fpy: 0 };
      return [record.total, record.defects, record.fpy];
    });

    if (out.length) {
      targetSheet.getRange(4, startCols[product], out.length, 3).setValues(out);
    }
  });

  // Only FPY columns are percentages.
  if (monthKeys.length) {
    targetSheet.getRange(4, 4, monthKeys.length, 1).setNumberFormat('0.00%');
    targetSheet.getRange(4, 7, monthKeys.length, 1).setNumberFormat('0.00%');
    targetSheet.getRange(4, 10, monthKeys.length, 1).setNumberFormat('0.00%');
    targetSheet.getRange(4, 13, monthKeys.length, 1).setNumberFormat('0.00%');
  }
}
