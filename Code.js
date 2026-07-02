const CONFIG = {
  SHEET_ID: '1Z627WaYnj0ZdWsjcho6sw0Jr3c5gjBWLoYHIxq0vL-4',

  TIME_CARD_FOLDER_ID: '1JSqdniofDVoFZWoxbCkIy46aYrSOJ9q3',
  EXPENSE_FOLDER_ID: '19GgJABuPeHAG8pmv1KzZs7QD_Rh-yFth',
  PROCESSED_TIME_CARD_FOLDER_ID: '1X2B88LvvyVsYJBJCEO4d1oy-heBXxjVl',
  PROCESSED_EXPENSE_FOLDER_ID: '1yNL1hMxAWPcO9NNsRCo04VAi2MyjGjZw',
  REVIEW_FOLDER_ID: '10X6OghLBJQdHjvifAX4mfcVqmK15GgX2',
  INVOICE_FOLDER_ID: '1Lf3uKKyCstKGXV_wU136kGMoNAdndftP',

  EMPLOYER: 'Innotech',
  HOURLY_RATE: 65,
  KM_PER_DAY: 132,
  KM_RATE: 0.73,
  REVIEW_DIFFERENCE_LIMIT: 0.50
};

const INVOICE_CONFIG = {
  MY_NAME:    'Simon Tremblay',
  MY_EMAIL:   'sims_tremblay@hotmail.com',
  MY_PHONE:   '+1(514) 817-9103',
  MY_ADDRESS: '843 rue Nadar, Saint-Jean-sur-Richelieu',
  MY_CITY:    'QC, Canada, J3B OE7',
  MY_PAYMENT: 'by direct deposit',
  TPS_NUMBER: '744633553 RT0001',
  TVQ_NUMBER: '4002578340 TQ0001',
  SIGN_OFF:   'Execaire sign-off : ___________________________',
  TEMPLATE_FILE_ID: '19yOcaYM2xk5DyMeGmYfDfkVORn71ykRD',
  EMPLOYERS: {
    'Innotech': ['Innotech Aviation', '10225 RYAN AVENUE', 'DORVAL-QUEBEC H9P-1A2']
  }
};

// Cached sheet reference — reused across all functions within one execution
let _sheet = null;
function getSheet_() {
  if (!_sheet) {
    _sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getActiveSheet();
  }
  return _sheet;
}

const COL = {
  EMPLOYER: 1,
  DATE: 2,
  PLANE: 3,
  HOURS: 4,
  SALARY: 5,
  KM: 6,
  EXPENSE: 7,
  TPS: 8,
  TVQ: 9,
  TIP: 10,
  DESCRIPTION: 11
};

function processContractorDocuments() {
  removeSummaryRows();

  processFolder(CONFIG.TIME_CARD_FOLDER_ID, 'TIME_CARD');
  processFolder(CONFIG.EXPENSE_FOLDER_ID, 'EXPENSE');

  rebuildSheet();
}

function rebuildSheet() {
  const sheet = getSheet_();
  removeSummaryRows();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numCols = COL.DESCRIPTION;
  const rawData = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const dataRows = rawData.filter(row => {
    const d = row[COL.DATE - 1];
    return d && !isSummaryLabel(d);
  });

  if (dataRows.length === 0) return;

  dataRows.sort((a, b) => {
    const da = new Date(normalizeDateKey(a[COL.DATE - 1]));
    const db = new Date(normalizeDateKey(b[COL.DATE - 1]));
    return da - db;
  });

  // Grouper par mois + employeur pour gérer plusieurs employeurs dans le même mois
  const groupMap = {};
  const groupOrder = [];
  let currentEmployer = CONFIG.EMPLOYER;
  dataRows.forEach(row => {
    const dateStr = normalizeDateKey(row[COL.DATE - 1]);
    if (!dateStr || dateStr.length < 7) return;
    const month = dateStr.substring(0, 7);
    const employerCell = (row[COL.EMPLOYER - 1] || '').toString().trim();
    if (employerCell && !employerCell.startsWith('Facture')) currentEmployer = employerCell;
    const key = month + '|' + currentEmployer;
    if (!groupMap[key]) {
      groupMap[key] = { month, employer: currentEmployer, rows: [] };
      groupOrder.push(key);
    }
    groupMap[key].rows.push(row);
  });

  let nextRow = 2;
  const groupMeta = groupOrder.map(key => {
    const { month, employer, rows } = groupMap[key];
    const startRow = nextRow;
    nextRow += rows.length;
    const brutRowNum = nextRow;
    nextRow += 2;
    nextRow++;
    return { key, month, employer, rows, startRow, endRow: startRow + rows.length - 1, brutRowNum };
  });

  const output = [];
  const boldRows = [];

  groupMeta.forEach(({ key, month, employer, rows, startRow, endRow, brutRowNum }) => {
    const facture = getFactureForMonth_(key);

    rows.forEach((row, i) => {
      const newRow = row.slice(0, numCols);
      if (i === 0) {
        newRow[COL.EMPLOYER - 1] = employer;
      } else if (i === rows.length - 1 && rows.length > 1) {
        newRow[COL.EMPLOYER - 1] = facture;
      } else {
        newRow[COL.EMPLOYER - 1] = '';
      }
      output.push(newRow);
    });

    const bRow = new Array(numCols).fill('');
    bRow[COL.DATE - 1] = 'Brut';
    [COL.HOURS, COL.SALARY, COL.KM, COL.EXPENSE, COL.TPS, COL.TVQ, COL.TIP].forEach(col => {
      bRow[col - 1] = `=SUM(${columnToLetter(col)}${startRow}:${columnToLetter(col)}${endRow})`;
    });
    output.push(bRow);
    boldRows.push(output.length);

    const nRow = new Array(numCols).fill('');
    nRow[COL.DATE - 1] = 'Net';
    nRow[COL.SALARY - 1] = `=${columnToLetter(COL.SALARY)}${brutRowNum}-${columnToLetter(COL.KM)}${brutRowNum}-${columnToLetter(COL.EXPENSE)}${brutRowNum}`;
    output.push(nRow);
    boldRows.push(output.length);

    output.push(new Array(numCols).fill(''));
  });

  sheet.getRange(2, 1, Math.max(lastRow - 1, output.length), numCols).clearContent();
  sheet.getRange(2, 1, output.length, numCols).setValues(output);

  boldRows.forEach(relRow => {
    sheet.getRange(1 + relRow, 1, 1, numCols).setFontWeight('bold');
  });
}

function getFactureForMonth_(key) {
  const props = PropertiesService.getScriptProperties();
  const stored = props.getProperty('FACTURE_MAP');
  const map = stored ? JSON.parse(stored) : {};

  if (!map[key]) {
    const counter = parseInt(props.getProperty('FACTURE_COUNTER') || '0') + 1;
    map[key] = 'Facture ' + String(counter).padStart(4, '0');
    props.setProperty('FACTURE_MAP', JSON.stringify(map));
    props.setProperty('FACTURE_COUNTER', String(counter));
  }

  return map[key];
}

function processFolder(folderId, expectedType) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    try {
      const extraction = analyzePdfWithOpenAI(file, expectedType);

      if (!extraction || extraction.needs_review) {
        moveFile(file, CONFIG.REVIEW_FOLDER_ID);
        continue;
      }

      normalizeExtraction(extraction);
      if (expectedType === 'EXPENSE' && extraction.expense) {
        extraction.expense.date = Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      validateExpenseMath(extraction);

      if (extraction.needs_review) {
        Logger.log('REVIEW: ' + file.getName() + ' - ' + extraction.review_reason);
        moveFile(file, CONFIG.REVIEW_FOLDER_ID);
        continue;
      }

      writeExtractionToSheet(extraction);

      if (expectedType === 'TIME_CARD') {
        moveFile(file, CONFIG.PROCESSED_TIME_CARD_FOLDER_ID);
      } else {
        moveFile(file, CONFIG.PROCESSED_EXPENSE_FOLDER_ID);
      }

    } catch (error) {
      Logger.log('ERROR with file ' + file.getName() + ': ' + error);
      moveFile(file, CONFIG.REVIEW_FOLDER_ID);
    }
  }
}

function analyzePdfWithOpenAI(file, expectedType) {
  let lastError = '';

  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      return callOpenAI(file, expectedType);
    } catch (error) {
      lastError = error;
      Logger.log('OpenAI attempt ' + attempt + ' failed for ' + file.getName() + ': ' + error);
      Utilities.sleep(3000);
    }
  }

  throw new Error('OpenAI failed after retries: ' + lastError);
}

function callOpenAI(file, expectedType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY missing in Script Properties.');

  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  const payload = {
    model: 'gpt-4.1-mini',
    input: [
      {
        role: 'user',
        content: [
          { type: 'input_text', text: buildPrompt(expectedType, file.getName()) },
          {
            type: 'input_file',
            filename: file.getName(),
            file_data: `data:${mimeType};base64,${base64}`
          }
        ]
      }
    ]
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseText = response.getContentText();

  if (!responseText.trim().startsWith('{')) {
    throw new Error('OpenAI returned non-JSON response: ' + responseText);
  }

  const json = JSON.parse(responseText);

  let outputText = '';

  if (
    json.output &&
    json.output.length > 0 &&
    json.output[0].content &&
    json.output[0].content.length > 0 &&
    json.output[0].content[0].text
  ) {
    outputText = json.output[0].content[0].text;
  }

  if (!outputText) {
    throw new Error('No text returned from OpenAI: ' + responseText);
  }

  return JSON.parse(cleanJson(outputText));
}

function buildPrompt(expectedType, fileName) {
  return `
You are an extraction engine for contractor accounting documents.

Expected document type: ${expectedType}
File name: ${fileName}

Return ONLY valid JSON. No markdown. No explanation.

GENERAL RULES:
- Employer is "${CONFIG.EMPLOYER}" unless another employer/client is clearly written.
- Hourly rate is ${CONFIG.HOURLY_RATE}.
- For Innotech time cards, mileage is ${CONFIG.KM_PER_DAY} km per worked DATE.
- Mileage rate is ${CONFIG.KM_RATE}.
- Mileage amount per worked DATE is ${(CONFIG.KM_PER_DAY * CONFIG.KM_RATE).toFixed(2)}.
- Each date on a time card must be its own final entry.
- If multiple work blocks exist on the same date, MERGE them into one entry.
- Use ONLY the numbers written in the R.T. / HOURS column as worked hours.
- Do NOT calculate hours from handwritten start/end times unless R.T. is missing.
- Do NOT add W.O. numbers.
- Ignore job descriptions on time cards.
- Extract aircraft registration if visible.
- Be very careful distinguishing "9H" from "C" or "G".
- Malta aircraft registrations start with "9H-", for example "9H-VTD".
- Canadian aircraft registrations start with "C-", for example "C-GJDR".
- USA aircraft registrations start with "N", for example "N123AB".
- Never output impossible aircraft format like "C14-VTD".
- If you see something like "9H VTD", output "9H-VTD".
- If no aircraft registration is visible, plane must be "".

TIME CARD output:
{
  "document_type": "TIME_CARD",
  "confidence": 0.95,
  "time_card": {
    "employer": "",
    "raw_time_blocks": [
      {
        "date": "YYYY-MM-DD",
        "plane": "",
        "rt_hours": 0
      }
    ]
  },
  "expense": null,
  "needs_review": false,
  "review_reason": ""
}

TIME CARD RULES:
- Return raw_time_blocks only.
- If a field labeled "TOTAL HOURS" or "TOTAL" is clearly written at the bottom of the card for a specific date, use that value as rt_hours for that date and return it as a SINGLE raw_time_blocks entry for that date. Do NOT return other RT blocks for that date.
- If no TOTAL HOURS field is visible, extract each visible R.T. / HOURS value as a separate raw_time_blocks item and let the script sum them.
- Do NOT add hours yourself when TOTAL HOURS is not present.
- Do NOT merge dates yourself.
- Do NOT calculate final daily totals when TOTAL HOURS is not present.
- The script will calculate totals later when TOTAL HOURS is absent.
- Ignore handwritten start/end times like 8:15 -> 14:00.
- Ignore W.O. numbers.
- rt_hours must only be the number written inside the R.T. / HOURS column, or the TOTAL HOURS field if present.

EXPENSE output:
{
  "document_type": "EXPENSE",
  "confidence": 0.95,
  "time_card": null,
  "expense": {
    "employer": "${CONFIG.EMPLOYER}",
    "date": "YYYY-MM-DD",
    "expense": 0,
    "tps": 0,
    "tvq": 0,
    "tip": 0,
    "total": 0,
    "description": ""
  },
  "needs_review": false,
  "review_reason": ""
}

EXPENSE RULES:
- expense = amount before taxes.
- TPS/GST goes in tps.
- TVQ/QST goes in tvq.
- tip goes in tip.
- total = final amount paid on the receipt.
- Copy money values exactly as written.
- Preserve 2 decimal places when visible.
- Never round money values yourself.
- description must be SHORT, 2 to 5 words maximum.
- Examples of good descriptions:
  - "Jumbo"
  - "Amir"
  - "Delibee"
  - "garderie bumbu"
  - "Tim Hortons"
  - "training food"
- Do NOT include long legal names, explanations, or file names.
- If a value is not visible, use 0.
- If date, total, or main amount is unclear, set needs_review true.
`;
}

function normalizeExtraction(data) {
  if (data.document_type === 'TIME_CARD' && data.time_card) {
    const rawBlocks = data.time_card.raw_time_blocks || data.time_card.entries || [];
    const grouped = {};

    rawBlocks.forEach(block => {
      const dateKey = normalizeDateKey(block.date);
      const plane = normalizePlane(block.plane || '');
      const hours = round2(block.rt_hours || block.hours || 0);

      if (!dateKey || !hours) return;

      if (!grouped[dateKey]) {
        grouped[dateKey] = {
          date: dateKey,
          plane: '',
          hours: 0,
          salary: 0,
          kilometrage: 0
        };
      }

      if (hours <= 0) {
        data.needs_review = true;
        data.review_reason = 'Heures à 0 ou manquantes pour la date ' + dateKey + '. Vérifier la time card.';
        return;
      }

      grouped[dateKey].hours = round2(grouped[dateKey].hours + hours);
      grouped[dateKey].plane = mergeText(grouped[dateKey].plane, plane);
    });

    data.time_card.entries = Object.keys(grouped).map(dateKey => {
      const entry = grouped[dateKey];
      entry.salary = round2(entry.hours * CONFIG.HOURLY_RATE);
      entry.kilometrage = round2(CONFIG.KM_PER_DAY * CONFIG.KM_RATE);
      return entry;
    });
  }

  if (data.document_type === 'EXPENSE' && data.expense) {
    data.expense.date = normalizeDateKey(data.expense.date);
    data.expense.expense = round2(data.expense.expense || 0);
    data.expense.tps = round2(data.expense.tps || 0);
    data.expense.tvq = round2(data.expense.tvq || 0);
    data.expense.tip = round2(data.expense.tip || 0);
    data.expense.total = round2(data.expense.total || 0);
    data.expense.description = shortenDescription(data.expense.description || '');

    // Si TPS et TVQ sont à 0 mais qu'il y a un total > subtotal, calculer les taxes par reverse
    if (data.expense.tps === 0 && data.expense.tvq === 0 && data.expense.total > 0 && data.expense.expense > 0) {
      const diff = round2(data.expense.total - data.expense.tip - data.expense.expense);
      if (diff > 0.05) {
        // Reverse TPS (5%) + TVQ (9.975%) sur le subtotal
        const subtotal = data.expense.expense;
        data.expense.tps = round2(subtotal * 0.05);
        data.expense.tvq = round2(subtotal * 0.09975);
      }
    }
  }
}

function validateExpenseMath(data) {
  if (data.document_type !== 'EXPENSE' || !data.expense) return data;

  const expense = data.expense;

  const subtotal = Number(expense.expense) || 0;
  const tps = Number(expense.tps) || 0;
  const tvq = Number(expense.tvq) || 0;
  const tip = Number(expense.tip) || 0;
  const total = Number(expense.total) || 0;

  if (!total) return data;

  const calculatedWithTip = round2(subtotal + tps + tvq + tip);
  const calculatedWithoutTip = round2(subtotal + tps + tvq);

  const diffWithTip = Math.abs(round2(calculatedWithTip - total));
  const diffWithoutTip = Math.abs(round2(calculatedWithoutTip - total));

  // If the receipt total seems to be before tip, accept it.
  // Example: total read = 24.01, but subtotal+tax+tip = 26.83.
  // This means total probably excludes tip or tip was handled separately.
  if (diffWithoutTip < CONFIG.REVIEW_DIFFERENCE_LIMIT) {
    return data;
  }

  // If total including tip balances, accept it.
  if (diffWithTip < CONFIG.REVIEW_DIFFERENCE_LIMIT) {
    return data;
  }

  data.needs_review = true;
  data.review_reason =
    'Expense math mismatch. Without tip: ' +
    calculatedWithoutTip +
    ', with tip: ' +
    calculatedWithTip +
    ', receipt total: ' +
    total +
    '. Difference with tip: ' +
    diffWithTip +
    ', difference without tip: ' +
    diffWithoutTip;

  return data;
}

function writeExtractionToSheet(data) {
  if (data.document_type === 'TIME_CARD') {
    data.time_card.entries.forEach(entry => {
      upsertRowByDate({
        employer: data.time_card.employer || CONFIG.EMPLOYER,
        date: entry.date,
        plane: entry.plane,
        hours: entry.hours,
        salary: entry.salary,
        km: entry.kilometrage
      });
    });
  }

  if (data.document_type === 'EXPENSE') {
    const expense = data.expense;

    upsertRowByDate({
      employer: expense.employer || CONFIG.EMPLOYER,
      date: expense.date,
      expense: expense.expense,
      tps: expense.tps,
      tvq: expense.tvq,
      tip: expense.tip,
      description: expense.description
    });
  }
}

function upsertRowByDate(data) {
  const sheet = getSheet_();
  const rowNumber = findRowByDate(data.date);

  if (rowNumber) {
    mergeIntoExistingRow(sheet, rowNumber, data);
  } else {
    appendNewRow(sheet, data);
  }
}

function findRowByDate(dateValue) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const target = normalizeDateKey(dateValue);
  const values = sheet.getRange(2, COL.DATE, lastRow - 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const existing = values[i][0];

    if (isSummaryLabel(existing)) continue;

    if (normalizeDateKey(existing) === target) {
      return i + 2;
    }
  }

  return null;
}

function appendNewRow(_, data) {
  const sheet = getSheet_();
  sheet.appendRow([
    data.employer || CONFIG.EMPLOYER,
    data.date || '',
    data.plane || '',
    data.hours || '',
    data.salary || '',
    data.km || '',
    data.expense || '',
    data.tps || '',
    data.tvq || '',
    data.tip || '',
    data.description || ''
  ]);
}

function mergeIntoExistingRow(_, row, data) {
  const sheet = getSheet_();
  if (data.employer) sheet.getRange(row, COL.EMPLOYER).setValue(data.employer);

  if (data.plane) {
    const oldPlane = sheet.getRange(row, COL.PLANE).getValue();
    sheet.getRange(row, COL.PLANE).setValue(mergeText(oldPlane, data.plane));
  }

  if (data.hours !== undefined && data.hours !== null && data.hours !== '') {
    sheet.getRange(row, COL.HOURS).setValue(round2(Number(data.hours)));
  }

  if (data.salary !== undefined && data.salary !== null && data.salary !== '') {
    sheet.getRange(row, COL.SALARY).setValue(round2(Number(data.salary)));
  }

  if (data.km !== undefined && data.km !== null && data.km !== '') {
    sheet.getRange(row, COL.KM).setValue(round2(Number(data.km)));
  }

  addNumber(sheet, row, COL.EXPENSE, data.expense);
  addNumber(sheet, row, COL.TPS, data.tps);
  addNumber(sheet, row, COL.TVQ, data.tvq);
  addNumber(sheet, row, COL.TIP, data.tip);

  if (data.description) {
    const oldDescription = sheet.getRange(row, COL.DESCRIPTION).getValue();
    sheet.getRange(row, COL.DESCRIPTION).setValue(mergeText(oldDescription, data.description));
  }
}

function addNumber(_, row, col, value) {
  const sheet = getSheet_();
  if (value === undefined || value === null || value === '') return;

  const cell = sheet.getRange(row, col);
  const newVal = round2(Number(value));
  const existing = cell.getFormula();
  if (existing && existing.startsWith('=')) {
    cell.setFormula(existing + '+' + newVal);
  } else {
    const oldValue = Number(cell.getValue()) || 0;
    if (oldValue !== 0) {
      cell.setFormula('=' + oldValue + '+' + newVal);
    } else {
      cell.setValue(newVal);
    }
  }
}

function sortSheetByDate() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;

  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort({ column: COL.DATE, ascending: true });
}

function addSummaryRows() {
  const sheet = getSheet_();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const brutRow = lastRow + 2;
  const netRow = lastRow + 3;

  sheet.getRange(brutRow, COL.DATE).setValue('Brut');
  sheet.getRange(netRow, COL.DATE).setValue('Net');

  const totalCols = [
    COL.HOURS,
    COL.SALARY,
    COL.KM,
    COL.EXPENSE,
    COL.TPS,
    COL.TVQ,
    COL.TIP
  ];

  totalCols.forEach(col => {
    const letter = columnToLetter(col);
    sheet.getRange(brutRow, col).setFormula(`=SUM(${letter}2:${letter}${lastRow})`);
  });

  sheet.getRange(netRow, COL.SALARY).setFormula(
    `=${columnToLetter(COL.SALARY)}${brutRow}-${columnToLetter(COL.KM)}${brutRow}-${columnToLetter(COL.EXPENSE)}${brutRow}`
  );

  sheet.getRange(brutRow, 1, 2, sheet.getLastColumn()).setFontWeight('bold');
}

function removeSummaryRows() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const values = sheet.getRange(2, COL.DATE, lastRow - 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const val = values[i][0];
    if (!val || isSummaryLabel(val)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function isSummaryLabel(value) {
  const v = String(value || '').trim().toUpperCase();
  return v === 'TOTAL' || v === 'BRUT' || v === 'NET';
}

function normalizeDateKey(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const s = String(value).trim();

  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return s;

  // Format slash M/J/AAAA ou J/M/AAAA — détecter intelligemment
  const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    const n1 = parseInt(slash[1]);
    const n2 = parseInt(slash[2]);
    const year = slash[3];
    let month, day;
    if (n2 > 12) {
      // Ex: 6/14/2026 → mois=6, jour=14 (format américain, jour > 12)
      month = String(n1).padStart(2, '0');
      day = String(n2).padStart(2, '0');
    } else if (n1 > 12) {
      // Ex: 14/6/2026 → jour=14, mois=6 (format québécois, jour > 12)
      day = String(n1).padStart(2, '0');
      month = String(n2).padStart(2, '0');
    } else {
      // Ambigu (ex: 7/6/2026) → assume québécois J/M/AAAA → jour=7, mois=6
      day = String(n1).padStart(2, '0');
      month = String(n2).padStart(2, '0');
    }
    return `${year}-${month}-${day}`;
  }

  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return s;
}

function normalizePlane(plane) {
  let p = String(plane).trim().toUpperCase();

  if (!p) return '';

  p = p.replace(/\s+/g, '-');

  if (p === 'C14-VTD' || p === 'C1H-VTD' || p === 'C9H-VTD') {
    return '9H-VTD';
  }

  const canadian = p.match(/^C-[A-Z0-9]{4}$/);
  if (canadian) return p;

  const malta = p.match(/^9H-[A-Z0-9]{3,4}$/);
  if (malta) return p;

  const usa = p.match(/^N[0-9A-Z]{3,5}$/);
  if (usa) return p;

  return p;
}

function shortenDescription(text) {
  if (!text) return '';

  let t = String(text).trim();

  t = t.replace(/restaurant/ig, '');
  t = t.replace(/food expense/ig, '');
  t = t.replace(/meal expense/ig, '');
  t = t.replace(/purchase/ig, '');
  t = t.replace(/services/ig, '');
  t = t.replace(/reservation.*$/ig, '');
  t = t.replace(/\s+/g, ' ').trim();

  if (t.length > 35) {
    t = t.substring(0, 35).trim();
  }

  return t;
}

function mergeText(oldText, newText) {
  if (!oldText) return newText;
  if (!newText) return oldText;

  const oldStr = String(oldText);
  const newStr = String(newText);

  if (oldStr.includes(newStr)) return oldStr;

  return oldStr + ' / ' + newStr;
}

function moveFile(file, destinationFolderId) {
  const destination = DriveApp.getFolderById(destinationFolderId);
  file.moveTo(destination);
}

function cleanJson(text) {
  return text
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();
}

function round2(num) {
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}

function columnToLetter(column) {
  let temp;
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function generateAllInvoices() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numCols = COL.DESCRIPTION;
  const rawData = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const groupMap = {};
  const groupOrder = [];
  let currentEmployer = CONFIG.EMPLOYER;

  rawData.forEach(row => {
    const dateVal = row[COL.DATE - 1];
    if (!dateVal || isSummaryLabel(dateVal)) return;
    const dateStr = normalizeDateKey(dateVal);
    if (!dateStr || dateStr.length < 7) return;

    const month = dateStr.substring(0, 7);
    const employerCell = (row[COL.EMPLOYER - 1] || '').toString().trim();
    if (employerCell && !employerCell.startsWith('Facture')) {
      currentEmployer = employerCell;
    }

    const key = month + '|' + currentEmployer;
    if (!groupMap[key]) {
      groupMap[key] = { month, employer: currentEmployer, rows: [] };
      groupOrder.push(key);
    }
    groupMap[key].rows.push(row);
  });

  groupOrder.forEach(key => {
    const { month, employer, rows } = groupMap[key];
    const salaryRows = rows.filter(r => (r[COL.HOURS - 1] || 0) > 0);
    if (salaryRows.length === 0) return;
    if (hasInvoiceBeenGenerated_(key)) {
      Logger.log('Facture déjà générée pour ' + key);
      return;
    }
    generateInvoice_(key, month, employer, rows);
  });
}

function generateInvoice_(key, month, employer, allRows) {
  const factureLabel = getFactureForMonth_(key);
  const factureNum = factureLabel.replace('Facture ', '');
  const year = month.substring(0, 4);
  const monthNum = parseInt(month.substring(5, 7));
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const monthName = monthNames[monthNum - 1];
  const lastDay = new Date(parseInt(year), monthNum, 0).getDate();
  const invoiceNumber = year + '-' + factureNum;
  const invoiceDate = monthName + ' ' + lastDay + ', ' + year;
  const fileName = 'Invoice ' + factureNum + ' - ' + monthName + ' ' + year + ' - ' + INVOICE_CONFIG.MY_NAME;

  const salaryRows = allRows.filter(r => (r[COL.HOURS - 1] || 0) > 0);
  if (salaryRows.length === 0) return;

  const subtotal = round2(salaryRows.reduce((s, r) => s + (r[COL.SALARY - 1] || 0), 0));
  const tps     = round2(subtotal * 0.05);
  const tvq     = round2(subtotal * 0.09975);
  const total   = round2(subtotal + tps + tvq);
  const employerInfo = INVOICE_CONFIG.EMPLOYERS[employer] || [employer, '', ''];
  const N = salaryRows.length;

  const templateFile = DriveApp.getFileById(INVOICE_CONFIG.TEMPLATE_FILE_ID);
  const templateBlob = templateFile.getBlob().setContentType('application/zip');
  const zipEntries = Utilities.unzip(templateBlob);

  const newEntries = zipEntries.map(function(entry) {
    const name = entry.getName();
    if (name === 'xl/workbook.xml') {
      let wbXml = entry.getDataAsString('UTF-8');
      const lastRow = 23 + N;
      const printArea = "'Invoice'!$A$1:$E$" + lastRow;
      if (wbXml.indexOf('_xlnm.Print_Area') !== -1) {
        wbXml = wbXml.replace(/'Invoice'!\$A\$1:\$E\$\d+/, printArea);
      } else {
        wbXml = wbXml.replace('<definedNames />', '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">' + printArea + '</definedName></definedNames>');
      }
      return Utilities.newBlob(wbXml, 'application/xml', name);
    }
    if (name !== 'xl/worksheets/sheet1.xml') return entry;

    let xml = entry.getDataAsString('UTF-8');

    // Simple placeholder replacements for static cells
    xml = xml.replace('{{INVOICE_NUMBER}}', xmlEscape_(invoiceNumber));
    xml = xml.replace('{{INVOICE_DATE}}',   xmlEscape_(invoiceDate));
    xml = xml.replace('{{EMPLOYER_NAME}}',  xmlEscape_(employerInfo[0]));
    xml = xml.replace('{{EMPLOYER_ADDR1}}', xmlEscape_(employerInfo[1]));
    xml = xml.replace('{{EMPLOYER_ADDR2}}', xmlEscape_(employerInfo[2]));

    // Replace rows 11+ (data + summary + footer) with dynamic content
    const row11Start  = xml.indexOf('<row r="11"');
    const sheetDataEnd = xml.indexOf('</sheetData>');
    const dynamicXml  = buildNewInvoiceRowsXml_(salaryRows, subtotal, tps, tvq, total, N);
    xml = xml.substring(0, row11Start) + dynamicXml + '</sheetData>' + xml.substring(sheetDataEnd + 12);

    // Update mergeCells row numbers (shift by N-2 since template has 2 placeholder rows)
    const shift = N - 2;
    xml = xml.replace('ref="B22:E22"', 'ref="B' + (22 + shift) + ':E' + (22 + shift) + '"');
    xml = xml.replace('ref="B23:E23"', 'ref="B' + (23 + shift) + ':E' + (23 + shift) + '"');
    xml = xml.replace('ref="C25:E25"', 'ref="C' + (25 + shift) + ':E' + (25 + shift) + '"');

    return Utilities.newBlob(xml, 'application/xml', name);
  });

  const newXlsx = Utilities.zip(newEntries, fileName + '.xlsx');
  const xlsxBlob = Utilities.newBlob(newXlsx.getBytes(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName + '.xlsx');
  const folder = DriveApp.getFolderById(CONFIG.INVOICE_FOLDER_ID);
  const file = folder.createFile(xlsxBlob);

  markInvoiceGenerated_(key, fileName);
  Logger.log('Facture generee : ' + fileName + ' -> ' + file.getUrl());
}

function formatDisplayDate_(val) {
  const s = normalizeDateKey(val);
  if (!s || s.length < 10) return String(val);
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const m = parseInt(s.substring(5, 7));
  const d = parseInt(s.substring(8, 10));
  const y = s.substring(0, 4);
  return months[m - 1] + ' ' + d + ', ' + y;
}

function fmtMoney_(n) {
  return '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function buildNewInvoiceRowsXml_(salaryRows, subtotal, tps, tvq, total, N) {
  // Style indices from new template:
  // s1=left margin, s13/s14/s15=odd row (B/C, D, E), s16/s17/s18=even row
  // s19=separator stripe, s20=summary label, s18=summary value
  // s21=TOTAL label (navy), s22=TOTAL value (navy)
  // s7=blue stripe, s23=footer small text, s24=thank you italic
  // s25=sign-off label, s10=sign-off line
  var xml = '';

  // Data rows
  salaryRows.forEach(function(row, i) {
    var r = 11 + i;
    var isOdd = (i % 2 === 0);
    var sBtC = isOdd ? '13' : '16';
    var sD   = isOdd ? '14' : '17';
    var sE   = isOdd ? '15' : '18';
    var dateStr = xmlEscape_(formatDisplayDate_(row[COL.DATE - 1]));
    var plane = row[COL.PLANE - 1];
    var planeStr = (!plane || String(plane) === 'None') ? '' : xmlEscape_(String(plane));
    var hours = row[COL.HOURS - 1] || 0;
    var amount = row[COL.SALARY - 1] || 0;
    xml += '<row r="' + r + '" ht="26" customHeight="1">';
    xml += '<c r="A' + r + '" s="1"/>';
    xml += '<c r="B' + r + '" s="' + sBtC + '" t="inlineStr"><is><t>' + dateStr + '</t></is></c>';
    xml += '<c r="C' + r + '" s="' + sBtC + '" t="inlineStr"><is><t>' + planeStr + '</t></is></c>';
    xml += '<c r="D' + r + '" s="' + sD + '" t="inlineStr"><is><t>' + hours + '</t></is></c>';
    xml += '<c r="E' + r + '" s="' + sE + '" t="inlineStr"><is><t>' + fmtMoney_(amount) + '</t></is></c>';
    xml += '</row>';
  });

  // Two empty spacer rows
  for (var e = 0; e < 2; e++) {
    var r2 = 11 + N + e;
    xml += '<row r="' + r2 + '"><c r="A' + r2 + '" s="1"/></row>';
  }

  // Separator stripe
  var sep = 13 + N;
  xml += '<row r="' + sep + '" ht="6" customHeight="1">';
  xml += '<c r="A' + sep + '" s="1"/><c r="B' + sep + '" s="19"/><c r="C' + sep + '" s="19"/><c r="D' + sep + '" s="19"/><c r="E' + sep + '" s="19"/>';
  xml += '</row>';

  // Summary rows
  function sumRow(offset, label, value, labelSty, valueSty) {
    var r3 = 14 + N + offset;
    xml += '<row r="' + r3 + '" ht="22" customHeight="1">';
    xml += '<c r="A' + r3 + '" s="1"/>';
    xml += '<c r="D' + r3 + '" s="' + labelSty + '" t="inlineStr"><is><t>' + label + '</t></is></c>';
    xml += '<c r="E' + r3 + '" s="' + valueSty + '" t="inlineStr"><is><t>' + value + '</t></is></c>';
    xml += '</row>';
  }
  sumRow(0, 'Subtotal',     fmtMoney_(subtotal), '20', '18');
  sumRow(1, 'TPS (5%)',     fmtMoney_(tps),      '20', '18');
  sumRow(2, 'TVQ (9.975%)', fmtMoney_(tvq),      '20', '18');
  sumRow(3, 'TOTAL',        fmtMoney_(total),    '21', '22');

  // Empty row before footer
  var emptyR = 18 + N;
  xml += '<row r="' + emptyR + '"><c r="A' + emptyR + '" s="1"/></row>';

  // Blue stripe
  var stripeR = 19 + N;
  xml += '<row r="' + stripeR + '" ht="8" customHeight="1">';
  xml += '<c r="A' + stripeR + '" s="1"/><c r="B' + stripeR + '" s="7"/><c r="C' + stripeR + '" s="7"/><c r="D' + stripeR + '" s="7"/><c r="E' + stripeR + '" s="7"/>';
  xml += '</row>';

  // Footer text (merged B:E in mergeCells)
  var footR = 20 + N;
  xml += '<row r="' + footR + '" ht="14" customHeight="1">';
  xml += '<c r="A' + footR + '" s="1"/>';
  xml += '<c r="B' + footR + '" s="23" t="inlineStr"><is><t>TPS # 744633553 RT0001     |     TVQ # 4002578340 TQ0001     |     Payment: by direct deposit</t></is></c>';
  xml += '</row>';

  // Thank you (merged B:E in mergeCells)
  var tyR = 21 + N;
  xml += '<row r="' + tyR + '" ht="28" customHeight="1">';
  xml += '<c r="A' + tyR + '" s="1"/>';
  xml += '<c r="B' + tyR + '" s="24" t="inlineStr"><is><t>Thank you for your trust!</t></is></c>';
  xml += '</row>';

  // Empty
  var emptyR2 = 22 + N;
  xml += '<row r="' + emptyR2 + '"><c r="A' + emptyR2 + '" s="1"/></row>';

  // Sign-off (C:E merged)
  var signR = 23 + N;
  xml += '<row r="' + signR + '" ht="18" customHeight="1">';
  xml += '<c r="A' + signR + '" s="1"/>';
  xml += '<c r="B' + signR + '" s="25" t="inlineStr"><is><t>Execaire sign-off :</t></is></c>';
  xml += '<c r="C' + signR + '" s="10" t="inlineStr"><is><t>_______________________________</t></is></c>';
  xml += '</row>';

  return xml;
}

function xmlEscape_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function dateToExcelSerial_(dateVal) {
  if (!dateVal) return 0;
  var d = dateVal instanceof Date ? dateVal : new Date(String(dateVal));
  var utc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  return Math.round(utc.getTime() / 86400000) + 25569;
}



function hasInvoiceBeenGenerated_(key) {
  const props = PropertiesService.getScriptProperties();
  const generated = JSON.parse(props.getProperty('INVOICES_GENERATED') || '{}');
  return !!generated[key];
}

function markInvoiceGenerated_(key, fileName) {
  const props = PropertiesService.getScriptProperties();
  const generated = JSON.parse(props.getProperty('INVOICES_GENERATED') || '{}');
  generated[key] = fileName;
  props.setProperty('INVOICES_GENERATED', JSON.stringify(generated));
}

function initFactureCounter() {
  const props = PropertiesService.getScriptProperties();
  // Clé = "YYYY-MM|Employeur"
  const map = {
    '2025-08|Sky Service': 'Facture 0001',
    '2025-09|Sky Service': 'Facture 0001',
    '2025-10|OG aviation':  'Facture 0002',
    '2025-11|PAL airlines': 'Facture 0003',
    '2025-12|PAL airlines': 'Facture 0004',
    '2025-12|SkyService':   'Facture 0005',
    '2026-02|Innotech':     'Facture 0006',
    '2026-03|Innotech':     'Facture 0007',
    '2026-04|Innotech':     'Facture 0008',
    '2026-05|Innotech':     'Facture 0009'
  };
  props.setProperty('FACTURE_MAP', JSON.stringify(map));
  props.setProperty('FACTURE_COUNTER', '9');
  Logger.log('Facture counter initialisé à 9. Prochain: Facture 0010.');
}

function debugOneTimeCard() {
  const folder = DriveApp.getFolderById(CONFIG.TIME_CARD_FOLDER_ID);
  const files = folder.getFiles();

  if (!files.hasNext()) {
    Logger.log('Aucun fichier dans Time Cards à traiter.');
    return;
  }

  const file = files.next();
  Logger.log('Testing file: ' + file.getName());

  const extraction = analyzePdfWithOpenAI(file, 'TIME_CARD');
  Logger.log('BEFORE NORMALIZE:\n' + JSON.stringify(extraction, null, 2));

  normalizeExtraction(extraction);
  Logger.log('AFTER NORMALIZE:\n' + JSON.stringify(extraction, null, 2));

  // File is NOT moved — debug only
}

function debugOneExpense() {
  const folder = DriveApp.getFolderById(CONFIG.EXPENSE_FOLDER_ID);
  const files = folder.getFiles();

  if (!files.hasNext()) {
    Logger.log('Aucun fichier dans Expenses à traiter.');
    return;
  }

  const file = files.next();
  Logger.log('Testing file: ' + file.getName());

  const extraction = analyzePdfWithOpenAI(file, 'EXPENSE');
  Logger.log('BEFORE NORMALIZE:\n' + JSON.stringify(extraction, null, 2));

  normalizeExtraction(extraction);
  validateExpenseMath(extraction);
  Logger.log('AFTER NORMALIZE + VALIDATE:\n' + JSON.stringify(extraction, null, 2));

  // File is NOT moved — debug only
}

function inspectAndFixFactureMap() {
  const props = PropertiesService.getScriptProperties();
  const map = JSON.parse(props.getProperty('FACTURE_MAP') || '{}');
  const counter = props.getProperty('FACTURE_COUNTER');
  Logger.log('FACTURE_COUNTER: ' + counter);
  Logger.log('FACTURE_MAP: ' + JSON.stringify(map, null, 2));

  // Fix: reassign 2026-06|Innotech to Facture 0010
  map['2026-06|Innotech'] = 'Facture 0010';
  // Fix counter to 10 if currently higher due to orphaned test runs
  if (parseInt(counter) > 10) {
    props.setProperty('FACTURE_COUNTER', '10');
  }
  props.setProperty('FACTURE_MAP', JSON.stringify(map));
  Logger.log('Fixed: 2026-06|Innotech -> Facture 0010');
}

function normalizeDatesInSheet() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const dateCol = COL.DATE;
  const range = sheet.getRange(2, dateCol, lastRow - 1, 1);
  const values = range.getValues();
  const updates = values.map(row => {
    const val = row[0];
    if (!val) return [val];
    const normalized = normalizeDateKey(val);
    return [normalized || val];
  });
  range.setNumberFormat('yyyy-MM-dd');
  range.setValues(updates);
  Logger.log('Dates normalisées en yyyy-MM-dd: ' + (lastRow - 1) + ' lignes.');
}

function resetJuneGenerated() {
  const props = PropertiesService.getScriptProperties();
  const gen = JSON.parse(props.getProperty('INVOICES_GENERATED') || '{}');
  delete gen['2026-06|Innotech'];
  props.setProperty('INVOICES_GENERATED', JSON.stringify(gen));
  Logger.log('June 2026 cleared from INVOICES_GENERATED.');
}

function resetMayGenerated() {
  const props = PropertiesService.getScriptProperties();
  const gen = JSON.parse(props.getProperty('INVOICES_GENERATED') || '{}');
  delete gen['2026-05|Innotech'];
  props.setProperty('INVOICES_GENERATED', JSON.stringify(gen));
  Logger.log('May 2026 cleared from INVOICES_GENERATED.');
}

function listInvoiceFiles() {
  const folder = DriveApp.getFolderById(CONFIG.INVOICE_FOLDER_ID);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    Logger.log(f.getName() + ' | ' + f.getDateCreated());
  }
}

function resetFactureMap() {
  const props = PropertiesService.getScriptProperties();
  const map = {
    '2025-08|Sky Service':  'Facture 0001',
    '2025-09|Sky Service':  'Facture 0001',
    '2025-10|OG aviation':  'Facture 0002',
    '2025-11|PAL airlines': 'Facture 0003',
    '2025-12|PAL airlines': 'Facture 0004',
    '2025-12|SkyService':   'Facture 0005',
    '2026-02|Innotech':     'Facture 0006',
    '2026-03|Innotech':     'Facture 0007',
    '2026-04|Innotech':     'Facture 0008',
    '2026-05|Innotech':     'Facture 0009',
    '2026-06|Innotech':     'Facture 0010'
  };
  props.setProperty('FACTURE_MAP', JSON.stringify(map));
  props.setProperty('FACTURE_COUNTER', '10');
  // Mark May and June as already generated
  const generated = JSON.parse(props.getProperty('INVOICES_GENERATED') || '{}');
  generated['2026-05|Innotech'] = 'Invoice 0009 - May 2026 - Simon Tremblay';
  generated['2026-06|Innotech'] = 'Invoice 0010 - June 2026 - Simon Tremblay';
  props.setProperty('INVOICES_GENERATED', JSON.stringify(generated));
  Logger.log('FACTURE_MAP reset. Counter=10. May+June marked as generated.');
}

function fixAmbiguousDates() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const fixes = {
    'Home Depot': '2026-06-07',
    'Claude Pro': '2026-06-05'
  };
  for (let i = 0; i < data.length; i++) {
    const desc = (data[i][COL.DESCRIPTION - 1] || '').toString();
    for (const keyword in fixes) {
      if (desc.indexOf(keyword) !== -1) {
        const cell = sheet.getRange(i + 1, COL.DATE);
        cell.setNumberFormat('yyyy-MM-dd');
        cell.setValue(fixes[keyword]);
        Logger.log('Fixed ' + keyword + ' date at row ' + (i + 1) + ' -> ' + fixes[keyword]);
      }
    }
  }
}

function fixHomeDepotDate() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const dateCell = data[i][1];
    const desc = (data[i][10] || '').toString();
    if (desc.indexOf('Home Depot') !== -1) {
      sheet.getRange(i + 1, 2).setValue('2026-06-07');
      Logger.log('Fixed Home Depot date at row ' + (i + 1) + ' -> 2026-06-07');
    }
  }
}

function forceRegenerateMayInvoice() {
  const folder = DriveApp.getFolderById(CONFIG.INVOICE_FOLDER_ID);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf('May 2026') !== -1) { f.setTrashed(true); Logger.log('Trashed: ' + f.getName()); }
  }
  const props = PropertiesService.getScriptProperties();
  const generated = JSON.parse(props.getProperty('INVOICES_GENERATED') || '{}');
  delete generated['2026-05|Innotech'];
  props.setProperty('INVOICES_GENERATED', JSON.stringify(generated));
  generateAllInvoices();
  Logger.log('done');
}



