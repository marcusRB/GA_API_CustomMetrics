var sheetName = 'Source_CustomMetric'; // rename the sheetName - globalVar 

function buildSourceSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var i, defaultMet, scopeRange, scopeRule, typeRange, typeRule, activeRange, activeRule ;
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  defaultMet = sheet.getRange(2, 2, 1, 4).getValues();
  sheet.getRange(1, 2, 1, 4).setValues([['Name', 'Scope', 'Type','Active']]);
  sheet.getRange(2, 1, 1, 1).setValue('DEFAULT/EMPTY');
  sheet.getRange(1, 1, 203, 5).setNumberFormat('@');
  if (isEmpty(defaultMet[0])) {
    sheet.getRange(2, 2, 1, 4).setNumberFormat('@').setValues([['(n/a)', 'HIT', 'INTEGER', 'false']]);
  }
  for (i = 1; i <= 200; i++) {
    sheet.getRange(2 + i, 1, 1, 1).setValue('ga:metric' + i);
  }
  
  // Set validation for SCOPE
  scopeRange = sheet.getRange(2, 3, 201, 1);
  scopeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['HIT','PRODUCT'])
    .setAllowInvalid(false)
    .setHelpText('Scope must be one of HIT or PRODUCT')
    .build();
  scopeRange.setDataValidation(scopeRule);
  
    // Set validation for TYPE
  typeRange = sheet.getRange(2, 4, 201, 1);
  typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['INTEGER', 'CURRENCY', 'TIME'])
    .setAllowInvalid(false)
    .setHelpText('Active must be one of INTEGER, CURRENCY or TIME')
    .build();
  typeRange.setDataValidation(typeRule);
  
  // Set validation for ACTIVE
  activeRange = sheet.getRange(2, 5, 201, 1);
  activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['true', 'false'])
    .setAllowInvalid(false)
    .setHelpText('Active must be one of true or false')
    .build();
  activeRange.setDataValidation(activeRule);
}


function fetchAccounts() {
  return Analytics.Management.AccountSummaries.list({
    fields: 'items(id,name,webProperties(id,name))'
  });
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function isEmpty(row) {
  return /^$/.test(row[0]) && /^$/.test(row[1]) && /^$/.test(row[2]) && /^$/.test(row[3]);
}

function isValid(row) {
  var name = row[0];
  var scope = row[1];
  var type = row[2];
  var active = row[3];
  return !/^$/.test(name) && /HIT|PRODUCT/.test(scope) && /INTEGER|CURRENCY|TIME/.test(type) && /true|false/.test(active) ;
}

function buildSourceData(sheet) {
  var range = sheet.getRange(3, 2, 200, 4).getValues();
  var defaultMet = sheet.getRange(3, 2, 1, 4).getValues();
  if (!isValid(defaultMet[0])) {
    throw new Error('Invalid source value found in DEFAULT/EMPTY row');
  }
  var sourceMetrics = [];
  var i;
  for (i = 0; i < range.length; i++) {
    if (!isEmpty(range[i]) && !isValid(range[i])) {
      throw new Error('Invalid source value found in metric ga:metric' + (i + 1));
    }
    if (!isEmpty(range[i])) {
      sourceMetrics.push({
        id: 'ga:metric' + (i + 1),
        name: range[i][0],
        scope: range[i][1],
        type: range[i][2],
        active: range[i][3]
        
      });
    } else {
      sourceMetrics.push({
        name: defaultMet[0][0] || '(n/a)',
        scope: defaultMet[0][1] || 'HIT',
        type: defaultMet[0][2] || 'INTEGER',
        active: defaultMet[0][3] || 'false'
      });
    }
  }
  return sourceMetrics;
}

function updateMetric(action, aid, pid, index, newMetric) {
  if (action === 'update') {
    return Analytics.Management.CustomMetrics.update(newMetric, aid, pid, 'ga:metric' + index, {ignoreCustomDataSourceLinks: true});
  }
  if (action === 'create') {
    return Analytics.Management.CustomMetrics.insert(newMetric, aid, pid);
  }
}

function startProcess(aid, pid, limit) {
  var metrics = Analytics.Management.CustomMetrics.list(aid, pid, {fields: 'items(id, name, scope, type, active)'});
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var sourceData = buildSourceData(sheet);
  var template = HtmlService.createTemplateFromFile('ProcessMetrics');
  template.data = {
    limit: limit,
    metrics: metrics,
    sourceData: sourceData,
    accountId: aid,
    propertyId: pid
  };
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(400).setHeight(400), 'Manage Custom Metrics for ' + pid);
} 

function isValidSheet(sheet) {
  var defaultMet = sheet.getRange(2, 2, 1, 4).getValues();
  var mets = sheet.getRange(3, 2, 200, 4).getValues();
  var i;
  if (!isValid(defaultMet[0])) {
    throw new Error('You must populate the DEFAULT/EMPTY row with proper values');
  }
  for (i = 0; i < mets.length; i++) {
    if (!isEmpty(mets[i]) && !isValid(mets[i])) {
      throw new Error('Invalid values for metric ga:metric' + (i + 1));
    }
  }
  return true;
}

function openMetricModal() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var html = HtmlService.createTemplateFromFile('PropertySelector').evaluate().setWidth(400).setHeight(280);
  if (!sheet) { 
    throw new Error('You need to create the Source Data sheet first');
  }
  if (!isValidSheet(sheet)) {
    throw new Error('You must populate the Source Data fields correctly');
  }
  SpreadsheetApp.getUi().showModalDialog(html, 'Select account and property for management');
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Google Analytics Custom Metric Manager')
      .addItem('1. Build/reformat Source Data sheet', 'buildSourceSheet')
      .addItem('2. Manage Custom Metrics', 'openMetricModal')
      .addToUi();
}