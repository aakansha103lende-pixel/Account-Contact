const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
const cors = require('cors');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

function toStringSafe(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return '';
  }
  return String(value);
}

function normalizeText(value) {
  return toStringSafe(value).trim().replace(/\s+/g, ' ');
}

function parseWorkbook(buffer, fileName) {
  const ext = path.extname(fileName).toLowerCase();
  if (ext === '.csv') {
    return XLSX.read(buffer.toString('utf8'), { type: 'string', raw: false });
  }
  return XLSX.read(buffer, { type: 'buffer', raw: false });
}

function workbookToJsonArray(workbook) {
  const sheetName = workbook.SheetNames[0];
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null, raw: false });
}

function dateSafe(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return value;
}

function appendIssue(existing, ruleName) {
  const trimmed = toStringSafe(existing).trim();
  if (!trimmed) {
    return ruleName;
  }
  const parts = trimmed.split(',').map(p => p.trim()).filter(Boolean);
  if (parts.includes(ruleName)) {
    return trimmed;
  }
  return `${trimmed}, ${ruleName}`;
}

function titleCase(text) {
  const value = toStringSafe(text).toLowerCase();
  return value.replace(/\b\w/g, (match) => match.toUpperCase());
}

function readDataFromBuffer(buffer, fileName) {
  const workbook = parseWorkbook(buffer, fileName);
  return workbookToJsonArray(workbook);
}

function createWorkbookFromRecords(sheetName, records, headers) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);
  if (headers.length) {
    worksheet.columns = headers.map((key) => ({ header: key, key }));
    records.forEach((row) => {
      const output = {};
      headers.forEach((key) => {
        output[key] = row[key] === undefined ? null : dateSafe(row[key]);
      });
      worksheet.addRow(output);
    });
  }
  return workbook;
}

function addWorksheetFromRecords(workbook, sheetName, records, headers) {
  const worksheet = workbook.addWorksheet(sheetName);
  if (headers.length) {
    worksheet.columns = headers.map((key) => ({ header: key, key }));
    records.forEach((row) => {
      const output = {};
      headers.forEach((key) => {
        output[key] = row[key] === undefined ? null : dateSafe(row[key]);
      });
      worksheet.addRow(output);
    });
  }
}

function cleanName(name) {
  if (name === null || name === undefined) {
    return '';
  }
  return normalizeText(String(name).toLowerCase());
}

function tokenSet(text) {
  const words = normalizeText(text).split(/\s+/).filter(Boolean);
  return new Set(words);
}

function calculateMatchScore(a, b) {
  if (!a || !b) {
    return 0;
  }
  const intersection = new Set([...a].filter((value) => b.has(value)));
  const maxLen = Math.max(a.size, b.size);
  if (maxLen === 0) {
    return 0;
  }
  return (intersection.size / maxLen) * 100;
}

function buildCleanedRecords(records, columnMapping) {
  return records.map((record) => {
    const cleaned = { ...record, Issues: '' };
    const actualColumns = Object.values(columnMapping).filter((columnName) => Object.prototype.hasOwnProperty.call(record, columnName));
    actualColumns.forEach((actualCol) => {
      const logical = Object.keys(columnMapping).find((key) => columnMapping[key] === actualCol);
      if (logical) {
        cleaned[`${logical}_corrected`] = toStringSafe(record[actualCol]);
      }
    });
    return cleaned;
  });
}

function buildActualColumns(records, columnMapping) {
  const actualCols = {};
  if (!records.length) {
    return actualCols;
  }
  const sample = records[0];
  for (const [logical, originalName] of Object.entries(columnMapping)) {
    if (Object.prototype.hasOwnProperty.call(sample, originalName)) {
      actualCols[logical] = originalName;
    }
  }
  return actualCols;
}

function cleanData(records, ruleFlags, columnMapping) {
  const df = records.map((record) => ({ ...record, Issues: '' }));
  const issueSummary = {};
  const highlightRows = new Set();
  const actualCols = buildActualColumns(df, columnMapping);

  const ruleDescriptions = {
    'Account Name - Short Form Correction: Private': "Standardizes 'pvt.' and 'pvt' in Account Name.",
    'Account Name - Short Form Correction: Private Limited': "Standardizes 'Pvt. Ltd.', 'pvt.limites', etc. → 'Private Limited' (no dot).",
    'Account Name - Short Form Correction: Limited': "Standardizes 'Ltd.' in Account Name.",
    'Account Name - Short Form Correction: LLP': "Standardizes 'LLP' in Account Name.",
    'Account Name - Short Form Correction: Brothers': "Corrects 'bros.' and 'bros' in Account Name.",
    'Account Name - Short Form Correction: and Sons': "Corrects '& sons' and 'and sons' in Account Name.",
    'Account Name - Short Form Correction: and Company': "Corrects '& co.', 'co', and 'comp' in Account Name.",
    'Account Name - Short Form Correction: Company (co)': "Corrects 'co' at the end of a line in Account Name.",
    'Account Name - Short Form Correction: Corporation': "Corrects 'corp.' in Account Name.",
    'Account Name - Short Form Correction: Incorporated': "Corrects 'inc.' in Account Name.",
    'Account Name - Short Form Correction: Manufacturing': "Corrects 'mfg.' in Account Name.",
    'Account Name - Short Form Correction: Traders': "Corrects 'trdrs.' in Account Name.",
    'Account Name - Short Form Correction: Associates': "Corrects 'assoc.' in Account Name.",
    'Account Name - Short Form Correction: Consultants': "Corrects 'cons.' in Account Name.",
    'Account Name - Short Form Correction: Services': "Corrects 'svc.' in Account Name.",
    'Account Name - Short Form Correction: Electrical': "Corrects 'Elec.' and 'Ele' to 'Electrical' in Account Name.",
    'Account Name - Remove Brackets': "Removes brackets and their contents from Account Name.",
    'Account Name - Contains Location': "Checks if Account Name contains common Indian city names and removes after ',' or '-'.",
    'Account Name - Invalid Prefixes': "Checks for invalid prefixes like 'M.s', 'M/s', or 'Messrs.'.",
    'Account Name - Remove Punctuation (. and ,)': "Removes punctuation marks like dots and commas from Account Name.",
    'Address - Abbreviation Correction: Apartment': "Corrects 'apt.' and 'aptt.' in Address.",
    'Address - Abbreviation Correction: Avenue': "Corrects 'ave.' and 'avn.' in Address.",
    'Address - Abbreviation Correction: Building': "Corrects 'bldg.' and 'bld.' in Address.",
    'Address - Abbreviation Correction: District': "Corrects 'dist.' and 'dis.' in Address.",
    'Address - Abbreviation Correction: Extension': "Corrects 'ext.' and 'extn.' in Address.",
    'Address - Abbreviation Correction: Floor': "Corrects 'flr.' and 'fl.' in Address.",
    'Address - Abbreviation Correction: Industrial': "Corrects 'ind.' and 'indust.' in Address.",
    'Address - Abbreviation Correction: Lane': "Corrects 'ln.' in Address.",
    'Address - Abbreviation Correction: Market': "Corrects 'mkt.' and 'mrkt.' in Address.",
    'Address - Abbreviation Correction: Near': "Corrects 'nr.' in Address.",
    'Address - Abbreviation Correction: Opposite': "Corrects 'opp.' and 'oppo.' in Address.",
    'Address - Abbreviation Correction: Post Office': "Corrects 'p.o.' and 'po' in Address.",
    'Address - Abbreviation Correction: Road': "Corrects 'rd.' and 'r.d.' in Address.",
    'Address - Abbreviation Correction: Street': "Corrects 'st.' and 'str.' in Address.",
    'Address - Abbreviation Correction: Town': "Corrects 'twn.' and 'tn.' in Address.",
    'Address - Abbreviation Correction: Village': "Corrects 'vill.' and 'vlg.' in Address.",
    'Address - Abbreviation Correction: House Number': "Corrects 'H. No.' and 'H/No.' in Address.",
    'Address - Abbreviation Correction: Colony': "Corrects 'col.' in Address.",
    'Pin Code - Missing': "Checks for blank or missing entries in the Pin Code column.",
    'Pin Code - 6 Digits': "Checks if Pin Code is exactly 6 digits.",
    'Job Title - Missing': "Checks for blank or missing entries in the Job Title column.",
    'Department - Missing': "Checks for blank or missing entries in the Department column.",
    'Mobile Number - Missing': "Checks for blank or missing entries in the Mobile Number column.",
    'Mobile Number - Format/Length': "Checks if Mobile Number has 10 digits and removes prefixes like +91/0.",
    'Phone Number - Missing': "Checks for blank or missing entries in the Phone Number column.",
    'Phone Number - Format/Length': "Checks if Phone Number starts with '0', has no spaces, and is 11 digits long.",
    'Fax Number - Missing': "Checks for blank or missing entries in the Fax Number column.",
    'Fax Number - Format/Length': "Checks if Fax Number starts with '0', has no spaces, and is 11 digits long.",
    'Account Name - Elect/elec/Elec → Electrical': "Replaces 'elect', 'elec', 'Elect', 'Elec' with 'Electrical'.",
    'Account Name - Ent/ent → Enterprise': "Replaces 'ent', 'Ent' with 'Enterprise'.",
    'Address - Clean Leading/Trailing Punctuation': "Removes commas and dots from start and end of Address.",
    'Address - Title Case': "Capitalizes first letter of each word in Address.",
    'Account Name - Engg/engg → Engineering': "Replaces 'engg', 'Engg', 'ENGG' with 'Engineering'",
    'Location - Remove Brackets & Content': "Removes brackets and their contents from Location column.",
    'Location - Clean Leading/Trailing Dots & Commas': "Removes dots, commas, spaces from start/end of Location.",
    'Location - Title Case': "Capitalizes first letter of every word in Location column.",
    'Location - Keep Only Last Value After Final Comma': "Keeps only the text after the last comma",
    'Location - Remove Digits': "Removes all numbers from Location"
  };

  function ensureCorrectedColumns() {
    for (const logical of Object.keys(actualCols)) {
      const corrected = `${logical}_corrected`;
      df.forEach((row) => {
        if (!(corrected in row)) {
          row[corrected] = toStringSafe(row[actualCols[logical]]);
        }
      });
    }
  }

  ensureCorrectedColumns();

  const abbreviationRules = {
    'Account Name - Short Form Correction: Private': { column: 'Account Name', patterns: { '\\bpvt\\.?\\b': 'Private' } },
    'Account Name - Short Form Correction: Private Limited': { column: 'Account Name', patterns: { '\\bpvt\\.?\\s*l(?:imites|ltd?)?.?\\b': 'Private Limited' } },
    'Account Name - Short Form Correction: Limited': { column: 'Account Name', patterns: { '\\bltd\\.?\\b': 'Limited' } },
    'Account Name - Short Form Correction: LLP': { column: 'Account Name', patterns: { '\\bllp\\b': 'Limited Liability Partnership' } },
    'Account Name - Short Form Correction: Brothers': { column: 'Account Name', patterns: { '\\bbros\\.?\\b': 'Brothers' } },
    'Account Name - Short Form Correction: and Sons': { column: 'Account Name', patterns: { '\\b& sons\\b|and sons\\b': 'and Sons' } },
    'Account Name - Short Form Correction: and Company': { column: 'Account Name', patterns: { '\\b& co\\.?\\b|\\bcomp\\.?\\b': 'Company' } },
    'Account Name - Short Form Correction: Company (co)': { column: 'Account Name', patterns: { '\\bco\\.?$': 'Company' } },
    'Account Name - Short Form Correction: Corporation': { column: 'Account Name', patterns: { '\\bcorp\\.?\\b': 'Corporation' } },
    'Account Name - Short Form Correction: Incorporated': { column: 'Account Name', patterns: { '\\binc\\.?\\b': 'Incorporated' } },
    'Account Name - Short Form Correction: Manufacturing': { column: 'Account Name', patterns: { '\\bmfg\\.?\\b': 'Manufacturing' } },
    'Account Name - Short Form Correction: Traders': { column: 'Account Name', patterns: { '\\btrdrs\\.?\\b': 'Traders' } },
    'Account Name - Short Form Correction: Associates': { column: 'Account Name', patterns: { '\\bassoc\\.?\\b': 'Associates' } },
    'Account Name - Short Form Correction: Consultants': { column: 'Account Name', patterns: { '\\bcons\\.?\\b': 'Consultants' } },
    'Account Name - Short Form Correction: Services': { column: 'Account Name', patterns: { '\\bsvc\\.?\\b': 'Services' } },
    'Account Name - Short Form Correction: Electrical': { column: 'Account Name', patterns: { '\\bele\\.?\\b': 'Electrical' } },
    'Address - Abbreviation Correction: Apartment': { column: 'Address', patterns: { '\\b(apt|aptt)\\.?\\b': 'Apartment' } },
    'Address - Abbreviation Correction: Avenue': { column: 'Address', patterns: { '\\b(ave|avn)\\.?\\b': 'Avenue' } },
    'Address - Abbreviation Correction: Building': { column: 'Address', patterns: { '\\b(bldg|bld)\\.?\\b': 'Building' } },
    'Address - Abbreviation Correction: District': { column: 'Address', patterns: { '\\b(dist|dis)\\.?\\b': 'District' } },
    'Address - Abbreviation Correction: Extension': { column: 'Address', patterns: { '\\b(ext|extn)\\.?\\b': 'Extension' } },
    'Address - Abbreviation Correction: Floor': { column: 'Address', patterns: { '\\b(flr|fl)\\.?\\b': 'Floor' } },
    'Address - Abbreviation Correction: Industrial': { column: 'Address', patterns: { '\\b(ind|indust)\\.?\\b': 'Industrial' } },
    'Address - Abbreviation Correction: Lane': { column: 'Address', patterns: { '\\bln\\.?\\b': 'Lane' } },
    'Address - Abbreviation Correction: Market': { column: 'Address', patterns: { '\\b(mkt|mrkt)\\.?\\b': 'Market' } },
    'Address - Abbreviation Correction: Near': { column: 'Address', patterns: { '\\bnr\\.?\\b': 'Near' } },
    'Address - Abbreviation Correction: Opposite': { column: 'Address', patterns: { '\\b(opp|oppo)\\.?\\b': 'Opposite' } },
    'Address - Abbreviation Correction: Post Office': { column: 'Address', patterns: { '\\bp\\.?o\\.?\\b': 'Post Office' } },
    'Address - Abbreviation Correction: Road': { column: 'Address', patterns: { '\\b(rd|r\\.d)\\.?\\b': 'Road' } },
    'Address - Abbreviation Correction: Street': { column: 'Address', patterns: { '\\b(st|str)\\.?\\b(?!\\s[A-Z])': 'Street' } },
    'Address - Abbreviation Correction: Town': { column: 'Address', patterns: { '\\b(twn|tn)\\.?\\b': 'Town' } },
    'Address - Abbreviation Correction: Village': { column: 'Address', patterns: { '\\b(vill|vlg)\\.?\\b': 'Village' } },
    'Address - Abbreviation Correction: House Number': { column: 'Address', patterns: { '\\b(h\\.? ?no\\.?|h/no\\.?)\\b': 'House No.' } },
    'Address - Abbreviation Correction: Colony': { column: 'Address', patterns: { '\\bcol\\.?\\b': 'Colony' } }
  };

  for (const [ruleName, rule] of Object.entries(abbreviationRules)) {
    const logical = rule.column;
    const correctedCol = `${logical}_corrected`;
    if (ruleFlags[ruleName] !== 'Y' || !(logical in actualCols)) {
      issueSummary[ruleName] = { count: 'Not Checked', description: ruleDescriptions[ruleName] || 'N/A', corrections: 'Not Checked' };
      continue;
    }

    let issues = 0;
    let corrections = 0;

    df.forEach((row) => {
      let text = toStringSafe(row[correctedCol]);
      const original = text;
      for (const [pattern, replacement] of Object.entries(rule.patterns)) {
        const regex = new RegExp(pattern, 'gi');
        if (regex.test(text)) {
          const replClean = replacement.replace(/\.$/, '');
          text = text.replace(regex, replClean);
        }
      }
      if (text !== original) {
        issues += 1;
        corrections += 1;
        text = text.replace(/\.+$/, '');
        row[correctedCol] = text;
        row.Issues = appendIssue(row.Issues, ruleName);
      }
    });

    issueSummary[ruleName] = { count: issues, description: ruleDescriptions[ruleName], corrections };
  }

  if ('Location' in actualCols) {
    const locCol = 'Location_corrected';
    let issuesAdded = false;

    const ruleBrackets = 'Location - Remove Brackets & Content';
    if (ruleFlags[ruleBrackets] === 'Y') {
      const regex = /\s*[\(\[\{\<].*?[\)\]\}\>]/g;
      const mask = df.map((row) => regex.test(toStringSafe(row[locCol])));
      if (mask.some(Boolean)) {
        df.forEach((row, idx) => {
          if (mask[idx]) {
            row[locCol] = toStringSafe(row[locCol]).replace(regex, '');
            row.Issues = appendIssue(row.Issues, ruleBrackets);
          }
        });
        issueSummary[ruleBrackets] = { count: mask.filter(Boolean).length, description: ruleDescriptions[ruleBrackets], corrections: mask.filter(Boolean).length };
        issuesAdded = true;
      } else {
        issueSummary[ruleBrackets] = { count: 0, description: ruleDescriptions[ruleBrackets], corrections: 0 };
      }
    } else {
      issueSummary[ruleBrackets] = { count: 'Not Checked', description: ruleDescriptions[ruleBrackets], corrections: 'Not Checked' };
    }

    const rulePunct = 'Location - Clean Leading/Trailing Dots & Commas';
    if (ruleFlags[rulePunct] === 'Y') {
      const regex = /^[.,\s]+|[.,\s]+$/g;
      const mask = df.map((row) => regex.test(toStringSafe(row[locCol])));
      if (mask.some(Boolean)) {
        df.forEach((row, idx) => {
          if (mask[idx]) {
            row[locCol] = toStringSafe(row[locCol]).replace(regex, '').trim();
            row.Issues = appendIssue(row.Issues, rulePunct);
          }
        });
        issueSummary[rulePunct] = { count: mask.filter(Boolean).length, description: ruleDescriptions[rulePunct], corrections: mask.filter(Boolean).length };
        issuesAdded = true;
      } else {
        issueSummary[rulePunct] = { count: 0, description: ruleDescriptions[rulePunct], corrections: 0 };
      }
    } else {
      issueSummary[rulePunct] = { count: 'Not Checked', description: ruleDescriptions[rulePunct], corrections: 'Not Checked' };
    }

    const ruleTitle = 'Location - Title Case';
    if (ruleFlags[ruleTitle] === 'Y') {
      df.forEach((row) => {
        const original = toStringSafe(row[locCol]);
        const titled = titleCase(original);
        if (original !== titled) {
          row[locCol] = titled;
          row.Issues = appendIssue(row.Issues, ruleTitle);
        }
      });
      issueSummary[ruleTitle] = { count: df.filter((row) => toStringSafe(row[locCol]) !== titleCase(toStringSafe(row[locCol]))).length, description: ruleDescriptions[ruleTitle], corrections: 0 };
    } else {
      issueSummary[ruleTitle] = { count: 'Not Checked', description: ruleDescriptions[ruleTitle], corrections: 'Not Checked' };
    }
  }

  const ruleEngg = 'Account Name - Engg/engg → Engineering';
  if (ruleFlags[ruleEngg] === 'Y' && 'Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    const regex = /\bengg\.?\b/gi;
    const mask = df.map((row) => regex.test(toStringSafe(row[correctedCol])));
    if (mask.some(Boolean)) {
      df.forEach((row, idx) => {
        if (mask[idx]) {
          row[correctedCol] = toStringSafe(row[correctedCol]).replace(regex, 'Engineering');
          row.Issues = appendIssue(row.Issues, ruleEngg);
          highlightRows.add(idx);
        }
      });
      issueSummary[ruleEngg] = { count: mask.filter(Boolean).length, description: ruleDescriptions[ruleEngg], corrections: mask.filter(Boolean).length };
    } else {
      issueSummary[ruleEngg] = { count: 0, description: ruleDescriptions[ruleEngg], corrections: 0 };
    }
  } else {
    issueSummary[ruleEngg] = { count: 'Not Checked', description: ruleDescriptions[ruleEngg], corrections: 'Not Checked' };
  }

  if ('Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    const regex = /\bPrivate\.\s*Limited\b/g;
    const mask = df.map((row) => regex.test(toStringSafe(row[correctedCol])));
    if (mask.some(Boolean)) {
      df.forEach((row, idx) => {
        if (mask[idx]) {
          row[correctedCol] = toStringSafe(row[correctedCol]).replace(regex, 'Private Limited');
          row.Issues = appendIssue(row.Issues, 'Account Name - Remove Dot: Private. Limited');
        }
      });
      issueSummary['Account Name - Remove Dot: Private. Limited'] = { count: mask.filter(Boolean).length, description: "Removes stray dot between 'Private' and 'Limited' (e.g., 'Private. Limited' → 'Private Limited')", corrections: mask.filter(Boolean).length };
    } else {
      issueSummary['Account Name - Remove Dot: Private. Limited'] = { count: 0, description: "Removes stray dot between 'Private' and 'Limited' (e.g., 'Private. Limited' → 'Private Limited')", corrections: 0 };
    }
  }

  const ruleKeyBrackets = 'Account Name - Remove Brackets';
  if (ruleFlags[ruleKeyBrackets] === 'Y' && 'Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    const regex = /\s*\([^)]*\)/g;
    const mask = df.map((row) => regex.test(toStringSafe(row[correctedCol])));
    if (mask.some(Boolean)) {
      df.forEach((row, idx) => {
        if (mask[idx]) {
          row[correctedCol] = toStringSafe(row[correctedCol]).replace(regex, '').trim();
          row.Issues = appendIssue(row.Issues, ruleKeyBrackets);
        }
      });
      issueSummary[ruleKeyBrackets] = { count: mask.filter(Boolean).length, description: ruleDescriptions[ruleKeyBrackets], corrections: mask.filter(Boolean).length };
    } else {
      issueSummary[ruleKeyBrackets] = { count: 0, description: ruleDescriptions[ruleKeyBrackets], corrections: 0 };
    }
  } else {
    issueSummary[ruleKeyBrackets] = { count: 'Not Checked', description: ruleDescriptions[ruleKeyBrackets], corrections: 'Not Checked' };
  }

  const ruleKeyLoc = 'Account Name - Contains Location';
  if (ruleFlags[ruleKeyLoc] === 'Y' && 'Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    const locations = [
      'andaman and nicobar islands', 'chandigarh', 'dadra and nagar haveli', 'daman and diu',
      'delhi', 'jammu and kashmir', 'ladakh', 'lakshadweep', 'puducherry',
      'andhra pradesh', 'arunachal pradesh', 'assam', 'bihar', 'chhattisgarh', 'goa', 'gujarat',
      'haryana', 'himachal pradesh', 'jharkhand', 'karnataka', 'kerala', 'madhya pradesh',
      'maharashtra', 'manipur', 'meghalaya', 'mizoram', 'nagaland', 'odisha', 'punjab',
      'rajasthan', 'sikkim', 'tamil nadu', 'telangana', 'tripura', 'uttar pradesh',
      'uttarakhand', 'west bengal',
      'mumbai', 'delhi', 'new delhi', 'bengaluru', 'bangalore', 'chennai', 'kolkata',
      'hyderabad', 'ahmedabad', 'pune', 'surat', 'jaipur', 'lucknow', 'nagpur', 'indore',
      'thane', 'bhopal', 'visakhapatnam', 'patna', 'vadodara', 'ludhiana', 'agra', 'nashik',
      'rajkot', 'madurai', 'kanpur', 'coimbatore', 'varanasi', 'meerut', 'faridabad', 'allahabad',
      'amritsar', 'aurangabad', 'ranchi', 'howrah', 'gwalior', 'jabalpur', 'vijayawada', 'mysore',
      'noida', 'ghaziabad', 'salem', 'trichy', 'tiruchirappalli', 'warangal', 'nellore', 'kalyan',
      'vasai', 'mira bhayandar', 'bhubaneswar', 'raipur', 'kota', 'moradabad', 'gurugram',
      'madgaon', 'panaji', 'bilaspur', 'ujjain', 'solapur', 'tirunelveli', 'malegaon', 'guntur',
      'nellore', 'jodhpur', 'chandigarh', 'ambala', 'rohtak', 'hisar', 'yamunanagar',
      'sonipat', 'panipat', 'kurukshetra', 'karnal', 'shimla', 'dehradun', 'haridwar',
      'rishikesh', 'nainital', 'bareilly', 'morena', 'satna', 'rewa', 'katni', 'betul',
      'ratlam', 'ujjain', 'dhar', 'bhila', 'korba', 'bilaspur', 'raigarh', 'durg', 'jammu',
      'srinagar', 'udhampur', 'anantnag', 'kargil', 'leh', 'pondicherry', 'karur', 'erode',
      'tiruppur', 'vellore', 'hosur', 'kumbakonam', 'nagercoil', 'thoothukudi', 'madurai',
      'salem', 'tirunelveli', 'thanjavur', 'trivandrum', 'kochi', 'kollam', 'thrissur',
      'alappuzha', 'kottayam', 'calicut', 'kozhikode', 'palakkad', 'malappuram', 'wayanad',
      'idukki', 'kasaragod', 'mangalore', 'hubli', 'dharwad', 'belgaum', 'bijapur',
      'gulbarga', 'bellary', 'tumkur', 'davangere', 'shimoga', 'bidar', 'chikmagalur',
      'hassan', 'udupi', 'mandya', 'chitradurga', 'gadag', 'raichur', 'bagalkot', 'koppal',
      'karwar', 'haveri', 'ranebennur', 'sirsi', 'gangavati', 'yadgir', 'vijayapura',
      'gulbarga', 'solan', 'hamirpur', 'kangra', 'una', 'mohali', 'bathinda', 'moga', 'patiala',
      'jalandhar', 'hoshiarpur', 'firozpur', 'ludhiana', 'bhatinda', 'barnala', 'pathankot',
      'ajmer', 'udaipur', 'bikaner', 'alwar', 'bharatpur', 'bhilwara', 'sikar', 'beawar',
      'churu', 'pali', 'jaisalmer', 'barmer', 'tonk', 'bundi', 'hanumangarh', 'sawai madhopur',
      'sirohi', 'nagaur', 'karauli', 'dholpur', 'banswara', 'dungarpur', 'jhalawar', 'rajsamand',
      'jaipur', 'udaipur', 'ahmednagar', 'akola', 'amravati', 'bhandara', 'beed', 'buldhana',
      'chandrapur', 'dhule', 'gadchiroli', 'gondia', 'hingoli', 'jalgaon', 'jalna', 'kolhapur',
      'latur', 'nanded', 'nandurbar', 'osmanabad', 'parbhani', 'ratnagiri', 'sangli', 'satara',
      'sindhudurg', 'solapur', 'wardha', 'washim', 'yavatmal', 'nashik', 'pune', 'thane',
      'mumbai', 'palghar', 'raigad', 'aurangabad', 'nagpur', 'chandrapur', 'gadchiroli', 'goa',
      'panaji', 'margao', 'mapusa', 'porvorim', 'vasco da gama','Domlur', 'Koramangala', 'Vanagaram', 'Dharmanagar',
      'DIPHU', 'Padubidari', 'Kudankulam', 'Megalift', 'Jhamarkotra Mines', 'Kulti', 'Silvassa', 'Arrow House', 'KHARGHAR', 'Katraj',
      'Warisaliganj', 'Khanapur', 'Atpadi', 'Dewas', 'Gopanari', 'RO Hyd','Sayan', 'Adajan', 'Maroli', 'Ankleshwar', 'Kamrej',
      'Patan', 'Girdhar Patelmarg', 'Deesa', 'Deesa', 'Kadodara', 'Manjeri', 'Thalappara', 'Kumbidi', 'Kunnumpuram', 'Kumbidi', 'Kavanoor', 'Tirur', 'Jahagirpura',
    ];

    const locationsLower = locations.map((loc) => loc.toLowerCase().trim());
    const textLowerList = df.map((row) => toStringSafe(row[correctedCol]).toLowerCase().trim());
    const escaped = locationsLower.map((loc) => loc.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const pattern = new RegExp(`[,|-]\\s*(?:${escaped.join('|')})\\b.*$`, 'gi');
    const mask = textLowerList.map((text) => pattern.test(text));
    if (mask.some(Boolean)) {
      df.forEach((row, idx) => {
        if (mask[idx]) {
          row[correctedCol] = toStringSafe(row[correctedCol]).replace(pattern, '').trim();
          row.Issues = appendIssue(row.Issues, ruleKeyLoc);
        }
      });
      issueSummary[ruleKeyLoc] = { count: mask.filter(Boolean).length, description: ruleDescriptions[ruleKeyLoc], corrections: mask.filter(Boolean).length };
    } else {
      issueSummary[ruleKeyLoc] = { count: 0, description: ruleDescriptions[ruleKeyLoc], corrections: 0 };
    }
  } else {
    issueSummary[ruleKeyLoc] = { count: 'Not Checked', description: ruleDescriptions[ruleKeyLoc], corrections: 'Not Checked' };
  }

  const ruleKeyPrefix = 'Account Name - Invalid Prefixes';
  if (ruleFlags[ruleKeyPrefix] === 'Y' && 'Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    const regex = /^(?:m\.s\.|m\/s|messrs\.?|m\/s\s)/i;
    const mask = df.map((row) => regex.test(toStringSafe(row[correctedCol])));
    if (mask.some(Boolean)) {
      df.forEach((row, idx) => {
        if (mask[idx]) {
          row[correctedCol] = toStringSafe(row[correctedCol]).replace(regex, '').trim();
          row.Issues = appendIssue(row.Issues, ruleKeyPrefix);
        }
      });
      issueSummary[ruleKeyPrefix] = { count: mask.filter(Boolean).length, description: ruleDescriptions[ruleKeyPrefix], corrections: mask.filter(Boolean).length };
    } else {
      issueSummary[ruleKeyPrefix] = { count: 0, description: ruleDescriptions[ruleKeyPrefix], corrections: 0 };
    }
  } else {
    issueSummary[ruleKeyPrefix] = { count: 'Not Checked', description: ruleDescriptions[ruleKeyPrefix], corrections: 'Not Checked' };
  }

  const ruleKeyPunct = 'Account Name - Remove Punctuation (. and ,)';
  if (ruleFlags[ruleKeyPunct] === 'Y' && 'Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    const regex = /[.,]+/g;
    const mask = df.map((row) => regex.test(toStringSafe(row[correctedCol])));
    if (mask.some(Boolean)) {
      df.forEach((row, idx) => {
        if (mask[idx]) {
          let cleaned = toStringSafe(row[correctedCol]).replace(regex, ' ');
          cleaned = cleaned.replace(/\s+/g, ' ').trim();
          row[correctedCol] = cleaned;
          row.Issues = appendIssue(row.Issues, ruleKeyPunct);
        }
      });
      issueSummary[ruleKeyPunct] = { count: mask.filter(Boolean).length, description: ruleDescriptions[ruleKeyPunct], corrections: mask.filter(Boolean).length };
    } else {
      issueSummary[ruleKeyPunct] = { count: 0, description: ruleDescriptions[ruleKeyPunct], corrections: 0 };
    }
  } else {
    issueSummary[ruleKeyPunct] = { count: 'Not Checked', description: ruleDescriptions[ruleKeyPunct], corrections: 'Not Checked' };
  }

  if ('Account Name' in actualCols) {
    const correctedCol = 'Account Name_corrected';
    df.forEach((row) => {
      row[correctedCol] = titleCase(toStringSafe(row[correctedCol]));
    });
  }

  const missingRules = {
    'Pin Code': 'Pin Code - Missing',
    'Job Title': 'Job Title - Missing',
    'Department': 'Department - Missing',
    'Mobile Number': 'Mobile Number - Missing',
    'Phone Number': 'Phone Number - Missing',
    'Fax Number': 'Fax Number - Missing'
  };

  for (const [logical, rule] of Object.entries(missingRules)) {
    if (ruleFlags[rule] === 'Y' && logical in actualCols) {
      const col = actualCols[logical];
      const mask = df.map((row) => {
        const value = row[col];
        const text = toStringSafe(value).trim();
        return text === '';
      });
      df.forEach((row, idx) => {
        if (mask[idx]) {
          row.Issues = appendIssue(row.Issues, rule);
        }
      });
      issueSummary[rule] = { count: mask.filter(Boolean).length, description: ruleDescriptions[rule], corrections: 0 };
    } else {
      issueSummary[rule] = { count: 'Not Checked', description: ruleDescriptions[rule], corrections: 'Not Checked' };
    }
  }

  const formatRules = {
    'Pin Code': {
      rule: 'Pin Code - 6 Digits',
      clean: (x) => {
        const raw = toStringSafe(x).trim();
        if (raw === '') {
          return '';
        }
        const parsed = Number(raw);
        if (!Number.isNaN(parsed) && /^\s*[+-]?(?:\d+|\d*\.\d+)\s*$/.test(raw)) {
          return String(Math.trunc(parsed));
        }
        return raw.replace(/\D/g, '');
      },
      valid: (x) => x.length === 6 && /^\d{6}$/.test(x)
    },
    'Mobile Number': {
      rule: 'Mobile Number - Format/Length',
      clean: (x) => {
        const cleaned = toStringSafe(x).replace(/\D/g, '');
        return cleaned.replace(/^\+91|^0/, '').trim();
      },
      valid: (x) => x.length === 10 && /^\d{10}$/.test(x)
    },
    'Phone Number': {
      rule: 'Phone Number - Format/Length',
      clean: (x) => toStringSafe(x).replace(/\D/g, ''),
      valid: (x) => x.length === 11 && /^0\d{10}$/.test(x)
    },
    'Fax Number': {
      rule: 'Fax Number - Format/Length',
      clean: (x) => toStringSafe(x).replace(/\D/g, ''),
      valid: (x) => x.length === 11 && /^0\d{10}$/.test(x)
    }
  };

  for (const [logical, spec] of Object.entries(formatRules)) {
    const rule = spec.rule;
    if (ruleFlags[rule] !== 'Y' || !(logical in actualCols)) {
      issueSummary[rule] = { count: 'Not Checked', description: ruleDescriptions[rule], corrections: 'Not Checked' };
      continue;
    }
    const col = actualCols[logical];
    const corrected = `${logical}_corrected`;
    const mask = df.map((row) => {
      const val = row[col];
      const text = toStringSafe(val).trim();
      return text !== '';
    });
    df.forEach((row, idx) => {
      if (mask[idx]) {
        const cleaned = spec.clean(row[col]);
        row[corrected] = cleaned;
      }
    });
    const invalid = df.map((row) => {
      const cleaned = toStringSafe(row[corrected]);
      return mask[df.indexOf(row)] && !spec.valid(cleaned);
    });
    df.forEach((row, idx) => {
      if (invalid[idx]) {
        row.Issues = appendIssue(row.Issues, rule);
      }
    });
    issueSummary[rule] = { count: invalid.filter(Boolean).length, description: ruleDescriptions[rule], corrections: mask.filter((val, idx) => val && !invalid[idx]).length };
  }

  df.forEach((row) => {
    row.Issues = toStringSafe(row.Issues).replace(/^,\s*|,\s*$/g, '').trim();
    if (row.Issues === '') {
      row.Issues = null;
    }
  });

  return { df, issueSummary, highlightRows };
}

function safeGet(data, key) {
  return key in data ? data[key] : null;
}

app.post('/api/match-duplicates', upload.fields([{ name: 'retailor', maxCount: 1 }, { name: 'sfdc', maxCount: 1 }]), async (req, res) => {
  try {
    if (!req.files || !req.files.retailor || !req.files.sfdc) {
      return res.status(400).json({ error: 'Both retailor and sfdc files are required.' });
    }

    const retailorWorkbook = parseWorkbook(req.files.retailor[0].buffer, req.files.retailor[0].originalname);
    const sfdcWorkbook = parseWorkbook(req.files.sfdc[0].buffer, req.files.sfdc[0].originalname);
    const retailorRows = workbookToJsonArray(retailorWorkbook);
    const sfdcRows = workbookToJsonArray(sfdcWorkbook);

    const retailorProcessed = retailorRows.map((row) => ({
      ...row,
      clean: cleanName(row['Account Name']),
      tokens: tokenSet(cleanName(row['Account Name']))
    }));
    const sfdcProcessed = sfdcRows.map((row) => ({
      ...row,
      clean: cleanName(row['Account Name']),
      tokens: tokenSet(cleanName(row['Account Name']))
    }));

    const merged = [];
    retailorProcessed.forEach((rowR) => {
      const pinR = normalizeText(rowR['Pin Code']);
      if (!pinR) {
        return;
      }
      sfdcProcessed.forEach((rowS) => {
        const pinS = normalizeText(rowS['Pin Code']);
        if (!pinS) {
          return;
        }
        if (pinR === pinS) {
          merged.push({
            ...Object.fromEntries(Object.entries(rowR).map(([k, v]) => [`${k}_r`, v])),
            ...Object.fromEntries(Object.entries(rowS).map(([k, v]) => [`${k}_s`, v])),
            'Pin Code': pinR,
            'Pin Code_r': pinR,
            'Pin Code_s': pinS,
            tokens_r: rowR.tokens,
            tokens_s: rowS.tokens
          });
        }
      });
    });

    merged.forEach((row) => {
      row.match_score = calculateMatchScore(row.tokens_r, row.tokens_s);
    });

    const bestMatches = new Map();
    merged.forEach((row) => {
      const accountName = toStringSafe(row['Account Name_r']);
      const existing = bestMatches.get(accountName);
      if (!existing || row.match_score > existing.match_score) {
        bestMatches.set(accountName, row);
      }
    });

    const best = Array.from(bestMatches.values());
    best.forEach((row) => {
      const rWords = normalizeText(row['clean_r']).split(/\s+/).filter(Boolean);
      const sWords = normalizeText(row['clean_s']).split(/\s+/).filter(Boolean);
      if (rWords.length < 3 || sWords.length < 3) {
        row.three_word_match = false;
      } else {
        row.three_word_match = rWords[0] === sWords[0] && rWords[1] === sWords[1] && rWords[2] === sWords[2];
      }
    });

    const finalData = best.filter((row) => {
      if (row.match_score > 67) {
        return true;
      }
      return row.match_score >= 60 && row.match_score <= 67 && row.three_word_match === true;
    });

    const outputFields = [
      'Account Name_r',
      'Account Name_s',
      'match_score',
      'three_word_match',
      'Account Record ID_r',
      'Account Record ID_s',
      'Pin Code',
      'Pin Code_r',
      'Pin Code_s'
    ];

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    worksheet.columns = outputFields.map((key) => ({ header: key, key }));
    finalData.forEach((row) => {
      const output = {};
      outputFields.forEach((field) => {
        output[field] = row[field] === undefined ? null : row[field];
      });
      worksheet.addRow(output);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=matched_duplicates.xlsx');
    res.send(Buffer.from(buffer));
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || 'Internal server error' });
  }
});

app.post('/api/clean-data', upload.single('dataFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Data file is required.' });
    }

    const records = readDataFromBuffer(req.file.buffer, req.file.originalname);
    const columnMapping = {
      'Account Record ID': 'Account RecoRoadID',
      'Account ID': 'Account ID',
      'Location': 'Location',
      'Account Name': 'Account Name',
      'Address': 'Address',
      'Pin Code': 'Pin Code',
      'Job Title': 'Job Title',
      'Department': 'Department',
      'Mobile Number': 'Mobile Number',
      'Phone Number': 'Phone',
      'Fax Number': 'Fax Number'
    };

    const ruleFlags = {
      'Account Name - Short Form Correction: Private': 'Y',
      'Account Name - Short Form Correction: Private Limited': 'Y',
      'Account Name - Short Form Correction: Limited': 'Y',
      'Account Name - Short Form Correction: LLP': 'Y',
      'Account Name - Short Form Correction: Brothers': 'Y',
      'Account Name - Short Form Correction: and Sons': 'Y',
      'Account Name - Short Form Correction: and Company': 'Y',
      'Account Name - Short Form Correction: Company (co)': 'Y',
      'Account Name - Short Form Correction: Corporation': 'Y',
      'Account Name - Short Form Correction: Incorporated': 'Y',
      'Account Name - Short Form Correction: Manufacturing': 'Y',
      'Account Name - Short Form Correction: Traders': 'Y',
      'Account Name - Short Form Correction: Associates': 'Y',
      'Account Name - Short Form Correction: Consultants': 'Y',
      'Account Name - Short Form Correction: Services': 'Y',
      'Account Name - Short Form Correction: Electrical': 'Y',
      'Account Name - Remove Punctuation (. and ,)': 'Y',
      'Account Name - Remove Brackets': 'Y',
      'Account Name - Contains Location': 'Y',
      'Account Name - Invalid Prefixes': 'Y',
      'Address - Abbreviation Correction: Apartment': 'Y',
      'Address - Abbreviation Correction: Avenue': 'Y',
      'Address - Abbreviation Correction: Building': 'Y',
      'Address - Abbreviation Correction: District': 'Y',
      'Address - Abbreviation Correction: Extension': 'Y',
      'Address - Abbreviation Correction: Floor': 'Y',
      'Address - Abbreviation Correction: Industrial': 'Y',
      'Address - Abbreviation Correction: Lane': 'Y',
      'Address - Abbreviation Correction: Market': 'Y',
      'Address - Abbreviation Correction: Near': 'Y',
      'Address - Abbreviation Correction: Opposite': 'Y',
      'Address - Abbreviation Correction: Post Office': 'Y',
      'Address - Abbreviation Correction: Road': 'Y',
      'Address - Abbreviation Correction: Street': 'Y',
      'Address - Abbreviation Correction: Town': 'Y',
      'Address - Abbreviation Correction: Village': 'Y',
      'Address - Abbreviation Correction: House Number': 'Y',
      'Address - Abbreviation Correction: Colony': 'Y',
      'Pin Code - Missing': 'Y',
      'Pin Code - 6 Digits': 'Y',
      'Job Title - Missing': 'Y',
      'Department - Missing': 'Y',
      'Mobile Number - Missing': 'Y',
      'Mobile Number - Format/Length': 'Y',
      'Phone Number - Missing': 'Y',
      'Phone Number - Format/Length': 'Y',
      'Fax Number - Missing': 'Y',
      'Fax Number - Format/Length': 'Y',
      'Account Name - Elect/elec/Elec → Electrical': 'Y',
      'Account Name - Ent/ent → Enterprise': 'Y',
      'Address - Clean Leading/Trailing Punctuation': 'Y',
      'Address - Title Case': 'Y',
      'Account Name - Engg/engg → Engineering': 'Y',
      'Location - Remove Brackets & Content': 'Y',
      'Location - Clean Leading/Trailing Dots & Commas': 'Y',
      'Location - Title Case': 'Y',
      'Location - Keep Only Last Value After Final Comma': 'Y',
      'Location - Remove Digits': 'Y'
    };

    const { df, issueSummary, highlightRows } = cleanData(records, ruleFlags, columnMapping);

    const summaryRecords = Object.entries(issueSummary).map(([ruleName, data]) => ({
      'Rule Name': ruleName,
      'Number of Issues Found': data.count,
      'Description': data.description,
      'Corrections': data.corrections
    }));

    const inputHeaders = Object.keys(records[0] || {});
    const baseCols = Object.values(columnMapping).filter((colName) => inputHeaders.includes(colName));
    const corrCols = Object.keys(columnMapping).map((logical) => `${logical}_corrected`).filter((colName) => df.some((row) => Object.prototype.hasOwnProperty.call(row, colName)));
    const finalCols = [...baseCols, ...corrCols, 'Issues'];

    const issueRecords = df.filter((row) => row.Issues !== null);

    const workbook = new ExcelJS.Workbook();
    const summarySheet = workbook.addWorksheet('Issue Summary');
    summarySheet.columns = ['Rule Name', 'Number of Issues Found', 'Description', 'Corrections'].map((key) => ({ header: key, key }));
    summaryRecords.forEach((row) => summarySheet.addRow(row));

    const issueSheet = workbook.addWorksheet('Issue Records');
    issueSheet.columns = finalCols.map((key) => ({ header: key, key }));
    issueRecords.forEach((row) => {
      const output = {};
      finalCols.forEach((col) => {
        output[col] = row[col] === undefined ? null : row[col];
      });
      issueSheet.addRow(output);
    });

    const cleanedSheet = workbook.addWorksheet('Full Cleaned Data');
    const cleanedHeaders = Array.from(new Set([...Object.keys(df[0] || {})]));
    cleanedSheet.columns = cleanedHeaders.map((key) => ({ header: key, key }));
    df.forEach((row) => {
      const output = {};
      cleanedHeaders.forEach((key) => {
        output[key] = row[key] === undefined ? null : row[key];
      });
      cleanedSheet.addRow(output);
    });

    if (highlightRows.size > 0) {
      const highlightColIndex = cleanedHeaders.indexOf('Account Name_corrected');
      if (highlightColIndex >= 0) {
        highlightRows.forEach((rowIndex) => {
          const excelRow = cleanedSheet.getRow(rowIndex + 2);
          const cell = excelRow.getCell(highlightColIndex + 1);
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFC7CE' },
            bgColor: { argb: 'FFFFC7CE' }
          };
        });
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=cleaned_output.xlsx');
    res.send(Buffer.from(buffer));
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || 'Internal server error' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
