/**
 * HubSpot ticket sync for the Monthly Quality reporting package.
 *
 * Credentials are read from Apps Script Script Properties. Never store a
 * HubSpot access key in source control.
 *
 * Required Script Property:
 *   HUBSPOT_ACCESS_TOKEN = HubSpot Personal Access Key
 */

const HUBSPOT_QUALITY_CONFIG_ = Object.freeze({
  baseUrl: 'https://api.hubapi.com',
  portalId: '44705249',
  tokenProperty: 'HUBSPOT_ACCESS_TOKEN',
  supportPipelineId: '0',
  rawSheetName: 'HubSpot Support Tickets',
  pageSize: 100,
  maxAttempts: 5
});

const HUBSPOT_QUALITY_TICKET_PROPERTIES_ = Object.freeze([
  'createdate',
  'hs_lastmodifieddate',
  'closed_date',
  'subject',
  'content',
  'hs_pipeline',
  'hs_pipeline_stage',
  'hs_ticket_category',
  'hs_ticket_priority',
  'hubspot_owner_id',
  'first_response_sla_met',
  'resolution_sla_met',
  'hs_time_to_first_response_sla_status',
  'hs_time_to_close_sla_status',
  'first_agent_reply_date',
  'last_reply_date',
  'product',
  'job_serial_',
  'job_number',
  'serial__',
  'problem___issue_description',
  'issue___warranty',
  'mjc_no_fault',
  'group',
  'category_tier_2',
  'category_tier_3',
  'n2_category_tier_1',
  'n2_category_tier_2',
  'n2_category_tier_3',
  'n3_category_tier_1',
  'n3_category_tier_2',
  'n3_category_tier_3',
  'repeat_issue',
  'number_of_service_issues'
]);

function validateHubSpotQualityConnection() {
  const data = hubSpotQualityGetJson_(
    '/crm/v3/objects/tickets?limit=1&archived=false'
  );
  const result = {
    connected: true,
    ticketReadAvailable: true,
    sampleTicketId: data.results && data.results.length
      ? data.results[0].id
      : null
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function syncHubSpotSupportTickets(asOfDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const context = getMonthlyQualityPackageContext_(asOfDate);
    const packageResult = ensureMonthlyQualityReportPackageUnlocked_(context);
    return syncHubSpotSupportTicketsToSpreadsheet_(
      packageResult.dataFile.getId()
    );
  } finally {
    lock.releaseLock();
  }
}

function syncHubSpotSupportTicketsToSpreadsheet_(spreadsheetId) {
  if (!getHubSpotQualityToken_()) {
    const skipped = {
      status: 'SKIPPED',
      reason: 'Missing Script Property ' + HUBSPOT_QUALITY_CONFIG_.tokenProperty,
      ticketCount: 0,
      scannedTicketCount: 0
    };
    Logger.log(JSON.stringify(skipped));
    return skipped;
  }

  const allTickets = getAllHubSpotTickets_();
  const supportTickets = allTickets.filter(function(ticket) {
    const properties = ticket.properties || {};
    return String(properties.hs_pipeline || '') ===
      String(HUBSPOT_QUALITY_CONFIG_.supportPipelineId);
  });

  supportTickets.sort(function(a, b) {
    return hubSpotQualityDateMillis_(b.properties && b.properties.createdate) -
      hubSpotQualityDateMillis_(a.properties && a.properties.createdate);
  });

  const ownerMap = getHubSpotOwnerMap_();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ensureHubSpotQualitySheet_(spreadsheet);
  writeHubSpotSupportTickets_(sheet, supportTickets, ownerMap);

  const result = {
    status: 'READY',
    ticketCount: supportTickets.length,
    scannedTicketCount: allTickets.length,
    spreadsheetId: spreadsheetId,
    sheetName: HUBSPOT_QUALITY_CONFIG_.rawSheetName,
    syncedAt: new Date().toISOString()
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function getAllHubSpotTickets_() {
  const encodedProperties = encodeURIComponent(
    HUBSPOT_QUALITY_TICKET_PROPERTIES_.join(',')
  );
  let path = '/crm/v3/objects/tickets?limit=' +
    HUBSPOT_QUALITY_CONFIG_.pageSize +
    '&archived=false&properties=' + encodedProperties;
  const tickets = [];
  let pageNumber = 0;

  while (path) {
    pageNumber++;
    const data = hubSpotQualityGetJson_(path);
    const results = data.results || [];
    Array.prototype.push.apply(tickets, results);
    Logger.log(
      'HubSpot tickets page ' + pageNumber + ': ' +
      results.length + ' rows; ' + tickets.length + ' total scanned.'
    );

    if (data.paging && data.paging.next && data.paging.next.after) {
      path = '/crm/v3/objects/tickets?limit=' +
        HUBSPOT_QUALITY_CONFIG_.pageSize +
        '&archived=false&properties=' + encodedProperties +
        '&after=' + encodeURIComponent(data.paging.next.after);
    } else {
      path = null;
    }
  }

  return tickets;
}

function getHubSpotOwnerMap_() {
  const owners = {};
  let path = '/crm/v3/owners?limit=500&archived=false';

  while (path) {
    const data = hubSpotQualityGetJson_(path);
    (data.results || []).forEach(function(owner) {
      const name = [owner.firstName, owner.lastName]
        .filter(function(value) { return Boolean(value); })
        .join(' ')
        .trim();
      owners[String(owner.id)] = name || owner.email || String(owner.id);
    });

    if (data.paging && data.paging.next && data.paging.next.after) {
      path = '/crm/v3/owners?limit=500&archived=false&after=' +
        encodeURIComponent(data.paging.next.after);
    } else {
      path = null;
    }
  }

  return owners;
}

function ensureHubSpotQualitySheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(HUBSPOT_QUALITY_CONFIG_.rawSheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(HUBSPOT_QUALITY_CONFIG_.rawSheetName);
  }
  return sheet;
}

function writeHubSpotSupportTickets_(sheet, tickets, ownerMap) {
  const headers = [
    'Ticket ID',
    'Created Date',
    'Last Modified',
    'Closed Date',
    'Subject',
    'Pipeline',
    'Pipeline Stage',
    'Ticket Category',
    'Priority',
    'Owner ID',
    'Owner Name',
    'First Response SLA Met',
    'Resolution SLA Met',
    'First Response SLA Status',
    'Resolution SLA Status',
    'First Agent Reply Date',
    'Last Customer Reply Date',
    'Product',
    'Job / Serial',
    'Job Number',
    'Serial',
    'MJC No Fault',
    'Group',
    'Category Tier 2',
    'Category Tier 3',
    'N2 Tier 1',
    'N2 Tier 2',
    'N2 Tier 3',
    'N3 Tier 1',
    'N3 Tier 2',
    'N3 Tier 3',
    'Repeat Issue',
    'Number of Service Issues',
    'Problem / Issue Description',
    'Warranty',
    'HubSpot Link',
    'Ticket Content'
  ];

  const rows = tickets.map(function(ticket) {
    const p = ticket.properties || {};
    const ownerId = String(p.hubspot_owner_id || '');
    const hubSpotLink = 'https://app.hubspot.com/contacts/' +
      HUBSPOT_QUALITY_CONFIG_.portalId + '/record/0-5/' + ticket.id;

    return [
      String(ticket.id || ''),
      hubSpotQualityDateValue_(p.createdate),
      hubSpotQualityDateValue_(p.hs_lastmodifieddate),
      hubSpotQualityDateValue_(p.closed_date),
      p.subject || '',
      p.hs_pipeline || '',
      p.hs_pipeline_stage || '',
      p.hs_ticket_category || '',
      p.hs_ticket_priority || '',
      ownerId,
      ownerMap[ownerId] || '',
      p.first_response_sla_met || '',
      p.resolution_sla_met || '',
      p.hs_time_to_first_response_sla_status || '',
      p.hs_time_to_close_sla_status || '',
      hubSpotQualityDateValue_(p.first_agent_reply_date),
      hubSpotQualityDateValue_(p.last_reply_date),
      p.product || '',
      p.job_serial_ || '',
      p.job_number || '',
      p.serial__ || '',
      p.mjc_no_fault || '',
      p.group || '',
      p.category_tier_2 || '',
      p.category_tier_3 || '',
      p.n2_category_tier_1 || '',
      p.n2_category_tier_2 || '',
      p.n2_category_tier_3 || '',
      p.n3_category_tier_1 || '',
      p.n3_category_tier_2 || '',
      p.n3_category_tier_3 || '',
      p.repeat_issue || '',
      p.number_of_service_issues || '',
      p.problem___issue_description || '',
      p.issue___warranty || '',
      hubSpotLink,
      p.content || ''
    ];
  });

  const requiredColumns = headers.length;
  const missingColumns = requiredColumns - sheet.getMaxColumns();
  if (missingColumns > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), missingColumns);
  }

  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.clear();

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    [2, 3, 4, 16, 17].forEach(function(columnNumber) {
      sheet.getRange(2, columnNumber, rows.length, 1)
        .setNumberFormat('yyyy-mm-dd hh:mm');
    });
    sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(5, 320);
  sheet.setColumnWidth(34, 360);
  sheet.setColumnWidth(36, 300);
  sheet.setColumnWidth(37, 420);
}

function hubSpotQualityGetJson_(path) {
  return hubSpotQualityRequestJson_('get', path, null);
}

function hubSpotQualityRequestJson_(method, path, payload) {
  const token = requireHubSpotQualityToken_();
  const options = {
    method: method,
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  };
  if (payload !== null && payload !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  for (let attempt = 1; attempt <= HUBSPOT_QUALITY_CONFIG_.maxAttempts; attempt++) {
    const response = UrlFetchApp.fetch(
      HUBSPOT_QUALITY_CONFIG_.baseUrl + path,
      options
    );
    const status = response.getResponseCode();
    const body = response.getContentText();

    if (status >= 200 && status < 300) {
      return body ? JSON.parse(body) : {};
    }

    if (status === 429 && attempt < HUBSPOT_QUALITY_CONFIG_.maxAttempts) {
      const headers = response.getAllHeaders();
      const retryAfter = Number(
        headers['Retry-After'] || headers['retry-after'] || 2
      );
      Utilities.sleep(Math.max(1, retryAfter) * 1000);
      continue;
    }

    if (status >= 500 && status <= 599 &&
        attempt < HUBSPOT_QUALITY_CONFIG_.maxAttempts) {
      Utilities.sleep(attempt * 2000);
      continue;
    }

    throw new Error(
      'HubSpot API request failed (' + status + '): ' + body
    );
  }

  throw new Error('HubSpot API request failed after all retry attempts.');
}

function getHubSpotQualityToken_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(
      HUBSPOT_QUALITY_CONFIG_.tokenProperty
    ) || ''
  ).trim();
}

function requireHubSpotQualityToken_() {
  const token = getHubSpotQualityToken_();
  if (!token) {
    throw new Error(
      'Missing Script Property ' + HUBSPOT_QUALITY_CONFIG_.tokenProperty +
      '. Store the HubSpot Personal Access Key in Apps Script Script Properties.'
    );
  }
  return token;
}

function hubSpotQualityDateValue_(value) {
  if (!value) return '';
  const date = new Date(value);
  return isNaN(date.getTime()) ? value : date;
}

function hubSpotQualityDateMillis_(value) {
  if (!value) return 0;
  const date = new Date(value);
  return isNaN(date.getTime()) ? 0 : date.getTime();
}
