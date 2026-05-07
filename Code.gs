function calculateFPYSummary_FINAL() {
  const token = PropertiesService.getScriptProperties().getProperty('SC_API_TOKEN');
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
  const failureDetailRows = [];

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

  function isDefectGateQuestion(label) {
    const lower = String(label || '').trim().toLowerCase();

    return (
      lower === 'were there any defects found?' ||
      lower === 'are there any defects found?' ||
      lower.includes('were there any defects found') ||
      lower.includes('are there any defects found')
    );
  }

  function findDefectAnswer(items) {
    if (!items || !Array.isArray(items)) return '';

    for (const item of items) {
      if (!item) continue;

      const label = String(item.label || item.title || item.data_label || '').trim();

      if (isDefectGateQuestion(label)) {
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

  function guessFailureCategory(text) {
    const s = String(text || '').toLowerCase();

    if (s.match(/wire|wiring|sensor|flow switch|transducer|vfd|relay|terminal|controller|a2l|electrical|voltage|harness/)) {
      return 'Electrical / Controls';
    }

    if (s.match(/leak|refrigerant|txv|compressor|braze|coil|pressure|superheat|charge|circuit/)) {
      return 'Refrigeration Circuit';
    }

    if (s.match(/fastener|screw|bolt|nut|washer|rivet|hardware|missing|ship.?with|loose item|panel|sensor missing|parts missing/)) {
      return 'Accessories / Ship-With';
    }

    if (s.match(/manual|diagram|drawing|label|nameplate|documentation|iom|instructions/)) {
      return 'Documentation';
    }

    if (s.match(/door|panel|cabinet|filter|fan|wheel|structural|damage|bracket/)) {
      return 'Airflow / Mechanical / Structural';
    }

    return 'Uncategorized';
  }

  function looksFailedOrNegative(item, answer) {
    const answerLower = String(answer || '').trim().toLowerCase();
    const negativeAnswers = [
      'fail', 'failed', 'false', 'no', 'not ok', 'ng', 'nok', 'bad'
    ];

    if (negativeAnswers.includes(answerLower)) return true;

    const score = item && item.score;
    if (score && (String(score.result || '').toLowerCase() === 'fail' || score.failed === true)) {
      return true;
    }

    if (item && (item.flagged === true || item.is_flagged === true)) return true;
    if (item && (item.failed === true || item.result === 'fail')) return true;

    return false;
  }

  function extractFailureDetails(items, context, rows, sectionName) {
    if (!items || !Array.isArray(items)) return;

    for (const item of items) {
      if (!item) continue;

      const label = String(item.label || item.title || item.data_label || '').trim();
      const lower = label.toLowerCase();
      const answer = extractAnswerFromResponses(item.responses);
      const currentSection = String(
        item.group_label || item.section_label || sectionName || ''
      ).trim();

      const informationalQuestion =
        lower.startsWith('is this ') ||
        lower.startsWith('is the ') ||
        lower.startsWith('does this ') ||
        lower.startsWith('does the ') ||
        lower.startsWith('are there ') ||
        lower.startsWith('were there ');

      if (informationalQuestion && !lower.includes('defect') && !lower.includes('issue')) {
        if (item.items) {
          extractFailureDetails(item.items, context, rows, currentSection || label);
        }
        continue;
      }

      if (isDefectGateQuestion(label)) {
        if (item.items) {
          extractFailureDetails(item.items, context, rows, currentSection || label);
        }
        continue;
      }

      const looksLikeFailureField =
        lower.includes('defect') ||
        lower.includes('issue') ||
        lower.includes('fail') ||
        lower.includes('nonconformance') ||
        lower.includes('corrective') ||
        lower.includes('comment') ||
        lower.includes('note');

      const answerLower = String(answer || '').trim().toLowerCase();
      const answerLooksUseful =
        answer &&
        !['no', 'n/a', 'na', 'none', 'ok', 'pass'].includes(answerLower);

      const failedOrFlagged = looksFailedOrNegative(item, answer);

      if ((looksLikeFailureField && answerLooksUseful) || failedOrFlagged) {
        rows.push([
          context.inspectionId,
          context.productLine,
          context.templateId,
          context.templateName,
          context.dateStr,
          context.monthKey,
          context.jobNumber || '',
          context.serialNumber || '',
          context.defectAnswer || '',
          label || '(unlabeled item)',
          answer || '',
          currentSection,
          guessFailureCategory(label + ' ' + answer + ' ' + currentSection)
        ]);
      }

      if (item.items) {
        extractFailureDetails(item.items, context, rows, currentSection || label);
      }
    }
  }

  function fetchInspectionDetail(id) {
    const candidates = [
      BASE + '/audits/' + encodeURIComponent(id),
      BASE + '/inspections/v1/inspections/' + encodeURIComponent(id)
    ];

    for (const endpoint of candidates) {
      try {
        const res = UrlFetchApp.fetch(endpoint, {
          method: 'get',
          headers: {
            Authorization: 'Bearer ' + token,
            Accept: 'application/json'
          },
          muteHttpExceptions: true
        });

        const code = res.getResponseCode();

        if (code === 200) {
          return JSON.parse(res.getContentText());
        }

        Logger.log('Inspection detail endpoint returned ' + code + ': ' + endpoint);
      } catch (err) {
        Logger.log('Inspection detail endpoint failed: ' + endpoint + ' | ' + err);
      }
    }

    Logger.log('No inspection detail found for inspection ID: ' + id);
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
      const dateCompletedStr = inspection.date_completed || dateStr;
      const context = {
        inspectionId: inspectionId,
        productLine: productLine,
        templateId: inspection.template_id || '',
        templateName: inspection.template_name || detail.template_name || '',
        dateStr: dateCompletedStr,
        monthKey: monthKey,
        jobNumber: inspection.reference || detail.reference || '',
        serialNumber: inspection.asset_id || '',
        defectAnswer: defectAnswer
      };

      if (defectFound) {
        extractFailureDetails(detail.header_items, context, failureDetailRows, '');
        extractFailureDetails(detail.items, context, failureDetailRows, '');
        extractFailureDetails(
          detail.audit_data && detail.audit_data.items,
          context,
          failureDetailRows,
          ''
        );
      }

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

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet =
    ss.getSheetByName('FPY_Failure_Details') || ss.insertSheet('FPY_Failure_Details');

  detailSheet.clear();
  detailSheet.getRange(1, 1, 1, 13).setValues([[
    'Inspection ID',
    'Product Line',
    'Template ID',
    'Template Name',
    'Date Completed',
    'Month',
    'Job Number',
    'Serial Number',
    'Defects Found',
    'Defect Question',
    'Defect Detail',
    'Section',
    'Category Guess'
  ]]);

  if (failureDetailRows.length) {
    detailSheet.getRange(2, 1, failureDetailRows.length, 13).setValues(failureDetailRows);
  }

  buildProductLineParetoTab(ss, failureDetailRows);
  buildPreviousMonthTopIssuesTab(ss, failureDetailRows, PRODUCTS);
}

function buildPreviousMonthTopIssuesTab(ss, detailRows, products) {
  const sheetName = 'FPY_Previous_Month_Top_Issues';
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  sheet.clear();
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  const now = new Date();
  const previousMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const previousMonthKey =
    previousMonthDate.getFullYear() +
    '-' +
    String(previousMonthDate.getMonth() + 1).padStart(2, '0');
  const previousMonthLabel = Utilities.formatDate(
    previousMonthDate,
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );

  sheet.getRange('A1:F1').merge().setValue('Top 3 Issues by Product Line - ' + previousMonthLabel);

  const headers = [
    'Product Line',
    'Rank',
    'Issue Category',
    'Count',
    'Issue Description',
    'Example Inspections'
  ];

  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);

  const grouped = {};

  detailRows.forEach(function(row) {
    const productLine = row[1] || 'Unknown';
    const monthKey = row[5] || '';

    if (monthKey !== previousMonthKey) return;

    const category = row[12] || 'Uncategorized';
    const defectQuestion = String(row[9] || '').trim();
    const defectDetail = String(row[10] || '').trim();
    const issueDetail = defectQuestion && defectDetail
      ? defectQuestion + ': ' + defectDetail
      : defectDetail || defectQuestion;
    const inspectionId = row[0] || '';
    const key = productLine + '|' + category;

    if (!grouped[key]) {
      grouped[key] = {
        productLine: productLine,
        category: category,
        count: 0,
        descriptions: [],
        inspections: []
      };
    }

    grouped[key].count++;

    if (issueDetail && grouped[key].descriptions.indexOf(issueDetail) === -1) {
      grouped[key].descriptions.push(issueDetail);
    }

    if (inspectionId && grouped[key].inspections.indexOf(inspectionId) === -1) {
      grouped[key].inspections.push(inspectionId);
    }
  });

  const rows = [];
  const productList = products && products.length ? products.slice() : [];

  Object.values(grouped).forEach(function(record) {
    if (productList.indexOf(record.productLine) === -1) {
      productList.push(record.productLine);
    }
  });

  productList.forEach(function(productLine) {
    const topIssues = Object.values(grouped)
      .filter(function(record) {
        return record.productLine === productLine;
      })
      .sort(function(a, b) {
        if (b.count !== a.count) return b.count - a.count;
        return a.category.localeCompare(b.category);
      })
      .slice(0, 3);

    topIssues.forEach(function(record, index) {
      rows.push([
        record.productLine,
        index + 1,
        record.category,
        record.count,
        record.descriptions.length
          ? record.descriptions.slice(0, 3).join(' | ')
          : record.category,
        record.inspections.slice(0, 5).join(', ')
      ]);
    });
  });

  if (rows.length) {
    sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
  } else {
    sheet.getRange('A4').setValue('No issue details found for ' + previousMonthLabel + '.');
  }

  const lastRow = Math.max(rows.length + 3, 4);
  sheet.getRange(1, 1, lastRow, headers.length).setBorder(
    true, true, true, true, true, true,
    'black',
    SpreadsheetApp.BorderStyle.SOLID
  );

  sheet.getRange('A1:F1')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  sheet.getRange(3, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (rows.length) {
    sheet.getRange(4, 4, rows.length, 1).setNumberFormat('0');
    sheet.getRange(4, 5, rows.length, 2).setWrap(true);
  }

  for (let col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(3, 190);
  sheet.setColumnWidth(5, 420);
  sheet.setColumnWidth(6, 220);
}

function buildProductLineParetoTab(ss, detailRows) {
  const sheet =
    ss.getSheetByName('FPY_Product_Line_Pareto') ||
    ss.insertSheet('FPY_Product_Line_Pareto');

  sheet.clear();
  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });

  const headers = [
    'Product Line',
    'Issue Category',
    'Count',
    'Cumulative',
    'Cumulative %'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (!detailRows.length) return;

  const grouped = {};

  detailRows.forEach(row => {
    const product = row[1] || 'Unknown';
    const category = row[12] || 'Uncategorized';
    const key = product + '|' + category;

    if (!grouped[key]) {
      grouped[key] = { product, category, count: 0 };
    }

    grouped[key].count++;
  });

  const sorted = Object.values(grouped).sort((a, b) => b.count - a.count);

  let cumulative = 0;
  const total = sorted.reduce((s, r) => s + r.count, 0);

  const output = sorted.map(r => {
    cumulative += r.count;
    return [
      r.product,
      r.category,
      r.count,
      cumulative,
      total ? cumulative / total : 0
    ];
  });

  sheet.getRange(2, 1, output.length, headers.length).setValues(output);
  sheet.getRange(2, 5, output.length).setNumberFormat('0.00%');

  const chart = sheet.newChart()
    .asComboChart()
    .addRange(sheet.getRange(1, 2, output.length + 1, 4))
    .setPosition(1, 7, 0, 0)
    .setOption('title', 'Pareto - Top Issues by Product Line')
    .setOption('seriesType', 'bars')
    .setOption('series', {
      2: { type: 'line', targetAxisIndex: 1 }
    })
    .setOption('vAxes', {
      0: { title: 'Count' },
      1: { title: 'Cumulative %', format: 'percent' }
    })
    .build();

  sheet.insertChart(chart);
}
