const compareForm = document.getElementById('compareForm');
const compareStatus = document.getElementById('compareStatus');
const analyticsCard = document.getElementById('analyticsCard');
const summaryGrid = document.getElementById('summaryGrid');
const issueTableBody = document.getElementById('issueTableBody');
const downloadReportButton = document.getElementById('downloadReport');
const accountUpload = document.getElementById('accountUpload');
const contactUpload = document.getElementById('contactUpload');
const combinedUpload = document.getElementById('combinedUpload');

const progressOverlay = document.getElementById('progressOverlay');
const progressTitle = document.getElementById('progressTitle');
const progressFill = document.getElementById('progressFill');
const progressStepList = document.getElementById('progressStepList');
const progressFootnote = document.getElementById('progressFootnote');

const detailOverlay = document.getElementById('detailOverlay');
const detailEyebrow = document.getElementById('detailEyebrow');
const detailTitle = document.getElementById('detailTitle');
const detailBody = document.getElementById('detailBody');
const detailClose = document.getElementById('detailClose');

let lastRunData = null; // { records, issueSummary, mode }

const LOCATION_NAMES = [
  'Maharashtra', 'Delhi', 'Karnataka', 'Tamil Nadu', 'Uttar Pradesh', 'Gujarat', 'Rajasthan',
  'West Bengal', 'Bihar', 'Telangana', 'Andhra Pradesh', 'Kerala', 'Punjab', 'Haryana',
  'Chennai', 'Mumbai', 'Bengaluru', 'Hyderabad', 'Pune', 'Kolkata', 'Jaipur', 'Lucknow',
  'Ahmedabad', 'Surat', 'Nagpur', 'Noida', 'Gurgaon', 'Faridabad', 'Kanpur', 'Agra',
  'Chandigarh', 'Vadodara', 'Coimbatore', 'Mysore', 'Pondicherry', 'Visakhapatnam', 'Indore'
];

let lastExportWorkbook = null;
let lastExportFilename = 'data-clean-report.xlsx';

function setStatus(message) {
  compareStatus.textContent = message;
}

function toStringSafe(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return '';
  }
  return String(value);
}

function normalizeText(value) {
  return toStringSafe(value).trim().replace(/\s+/g, ' ');
}

function properCase(value) {
  const text = normalizeText(value).toLowerCase();
  return text.replace(/\b[a-z]/g, (match) => match.toUpperCase());
}

function stripSpecial(value) {
  return normalizeText(value).replace(/[^a-zA-Z0-9\s&\/\-\.]/g, ' ').replace(/\s+/g, ' ').trim();
}

function removeTextInBrackets(value) {
  return normalizeText(value).replace(/[\(\[\{][^\)\]\}]*[\)\]\}]/g, ' ').replace(/\s+/g, ' ').trim();
}

function removeInvalidAccountPrefix(value) {
  return normalizeText(value).replace(/^(m\/s\.?|m\.s\.?|messrs\.?\s*)/i, '').trim();
}

function replaceEntToEnterprise(value) {
  return normalizeText(value).replace(/\b(ent)\b/gi, 'Enterprise');
}

function removePrivateDotLimited(value) {
  return normalizeText(value).replace(/private\.\s*limited/gi, 'Private Limited');
}

function stripTerminalDotsAndCommas(value) {
  return normalizeText(value).replace(/^[\.,\s]+|[\.,\s]+$/g, '').trim();
}

function cleanAccountNameLocation(value) {
  const text = normalizeText(value);
  const lower = text.toLowerCase();
  for (const location of LOCATION_NAMES) {
    const pattern = new RegExp(`[,-]\\s*${location.replace(/[-\\/\\^$*+?.()|[\]{}]/g, '\\$&')}(?:\b|$)`, 'i');
    const match = lower.match(pattern);
    if (match) {
      const cutoff = text.slice(0, match.index).trim();
      return stripSpecial(cutoff);
    }
  }
  return text;
}

function cleanPhone(value) {
  const digits = toStringSafe(value).replace(/\D/g, '');
  if (digits.length === 11 && digits.startsWith('0')) {
    return digits;
  }
  if (digits.length === 10) {
    return `0${digits}`;
  }
  if (digits.length === 12 && digits.startsWith('91')) {
    return `0${digits.slice(2)}`;
  }
  return digits;
}

function cleanMobile(value) {
  const digits = toStringSafe(value).replace(/\D/g, '');
  if (digits.length === 10) {
    return digits;
  }
  if (digits.length === 11 && digits.startsWith('0')) {
    return digits.slice(1);
  }
  if (digits.length === 12 && digits.startsWith('91')) {
    return digits.slice(2);
  }
  return digits;
}

function cleanPinCode(value) {
  const digits = toStringSafe(value).replace(/\D/g, '');
  return digits;
}

function cleanEmail(value) {
  return normalizeText(value).toLowerCase();
}

function readFileAsWorkbook(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Unable to read file'));
    reader.onload = (event) => {
      try {
        const data = event.target.result;
        const workbook = ext === '.csv'
          ? XLSX.read(data, { type: 'string', raw: false })
          : XLSX.read(data, { type: 'array', raw: false });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };
    if (ext === '.csv') {
      reader.readAsText(file, 'utf8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

function workbookToJsonArray(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });
}

function normalizeHeaderValue(value) {
  return toStringSafe(value).trim().toLowerCase();
}

function buildActualColumns(records, columnMapping) {
  const actual = {};
  if (!records.length) return actual;
  const sample = records[0];
  const headers = Object.keys(sample).reduce((acc, key) => {
    acc[normalizeHeaderValue(key)] = key;
    return acc;
  }, {});

  Object.entries(columnMapping).forEach(([logical, candidates]) => {
    const found = candidates.find((candidate) => headers[normalizeHeaderValue(candidate)]);
    if (found) {
      actual[logical] = headers[normalizeHeaderValue(found)];
    }
  });

  return actual;
}

function findHeaderKey(sampleRow, candidates) {
  const headers = Object.keys(sampleRow).reduce((acc, key) => {
    acc[normalizeHeaderValue(key)] = key;
    return acc;
  }, {});
  const found = candidates.find((candidate) => headers[normalizeHeaderValue(candidate)]);
  return found ? headers[normalizeHeaderValue(found)] : null;
}

// Some files store the contact's name as separate First Name / Last Name
// columns instead of a single "Contact Name" column. If no Contact Name
// column exists, synthesize one by combining First + Last Name so that
// contact cleaning and contact dedupe both have a name to work with.
function ensureContactNameColumn(records, columnMapping) {
  if (!records.length) return false;
  const sample = records[0];
  const contactNameCandidates = columnMapping['Contact Name'] || ['Contact Name'];
  if (findHeaderKey(sample, contactNameCandidates)) return false; // already present

  const firstNameCandidates = ['First Name', 'FirstName', 'Fname', 'First', 'Given Name'];
  const lastNameCandidates = ['Last Name', 'LastName', 'Lname', 'Last', 'Surname', 'Family Name'];
  const firstKey = findHeaderKey(sample, firstNameCandidates);
  const lastKey = findHeaderKey(sample, lastNameCandidates);
  if (!firstKey && !lastKey) return false; // nothing to combine

  records.forEach((row) => {
    const first = firstKey ? toStringSafe(row[firstKey]) : '';
    const last = lastKey ? toStringSafe(row[lastKey]) : '';
    row['Contact Name'] = normalizeText(`${first} ${last}`);
  });
  return true;
}

function syncCombinedContactName(records, actualCols) {
  if (!records.length) return;
  const firstCorrectedKey = actualCols['First Name'] ? 'First Name_corrected' : null;
  const lastCorrectedKey = actualCols['Last Name'] ? 'Last Name_corrected' : null;
  if (!firstCorrectedKey && !lastCorrectedKey) return;

  records.forEach((row) => {
    const first = firstCorrectedKey ? normalizeText(row[firstCorrectedKey]) : '';
    const last = lastCorrectedKey ? normalizeText(row[lastCorrectedKey]) : '';
    const combined = normalizeText([first, last].filter(Boolean).join(' '));
    row['Contact Name'] = combined;
    row['Contact Name_corrected'] = combined;
  });
}

function ensureCorrectedColumns(records, actualCols) {
  records.forEach((row) => {
    Object.keys(actualCols).forEach((logical) => {
      const correctedKey = `${logical}_corrected`;
      if (!(correctedKey in row)) {
        row[correctedKey] = toStringSafe(row[actualCols[logical]]);
      }
    });
  });
}

function appendIssue(existing, issueName) {
  const current = normalizeText(existing);
  if (!current) return issueName;
  const parts = current.split(',').map((part) => part.trim()).filter(Boolean);
  if (parts.includes(issueName)) return current;
  return `${current}, ${issueName}`;
}

function autoSizeColumns(worksheet, minWidth = 10, maxWidth = 50) {
  worksheet.columns.forEach((column) => {
    let maxLength = minWidth;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const value = cell.value;
      const text = value == null ? '' : String(value).replace(/\r\n/g, ' ');
      maxLength = Math.max(maxLength, Math.min(text.length + 2, maxWidth));
    });
    column.width = maxLength;
  });
}

function styleHeaderRow(worksheet) {
  worksheet.getRow(1).eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD4E4F7' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
}

function saveWorkbookToFile(workbook, filename) {
  return workbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  });
}

function buildRuleFlags() {
  return {
    'Account Name - Remove Brackets': 'Y',
    'Account Name - Contains Location': 'Y',
    'Account Name - Invalid Prefixes': 'Y',
    'Account Name - Remove Punctuation (. and ,)': 'Y',
    'Account Name - Ent/ent to Enterprise': 'Y',
    'Account Name - Remove Dot: Private. Limited': 'Y',
    'Account Name - Short form correction': 'Y',
    'Account Name - Proper case': 'Y',
    'Address - Clean Leading/Trailing Punctuation': 'Y',
    'Address - Apartment': 'Y',
    'Address - Avenue': 'Y',
    'Address - Building': 'Y',
    'Address - District': 'Y',
    'Address - Extension': 'Y',
    'Address - Floor': 'Y',
    'Address - Industrial': 'Y',
    'Address - Lane': 'Y',
    'Address - Market': 'Y',
    'Address - Near': 'Y',
    'Address - Opposite': 'Y',
    'Address - Post Office': 'Y',
    'Address - Road': 'Y',
    'Address - Street': 'Y',
    'Address - Town': 'Y',
    'Address - Village': 'Y',
    'Address - House Number': 'Y',
    'Address - Colony': 'Y',
    'Address - Proper case': 'Y',
    'Location - Keep Only Last Value After Final Comma': 'Y',
    'Location - Remove Brackets & Content': 'Y',
    'Location - Remove Digits': 'Y',
    'Location - Clean Leading/Trailing Dots & Commas': 'Y',
    'Location - Proper case': 'Y',
    'Pin Code - Missing': 'Y',
    'Job Title - Missing': 'Y',
    'Department - Missing': 'Y',
    'Mobile Number - Missing': 'Y',
    'Phone Number - Missing': 'Y',
    'Fax Number - Missing': 'Y',
    'Pin Code - 6 digits': 'Y',
    'Phone Number - 11 digits': 'Y',
    'Mobile Number - 10 digits': 'Y',
    'Fax Number - Format/Length': 'Y',
    'Email - lowercase': 'Y',
    'Job Title - Proper case': 'Y',
    'Department - Proper case': 'Y',
    'First Name - Proper case': 'Y',
    'Last Name - Proper case': 'Y',
    'Contact Name - Proper case': 'Y'
  };
}

function buildColumnMapping() {
  return {
    'Account Name': ['Account Name', 'AccountName', 'Company Name', 'Legal Name'],
    'Address': ['Address', 'Street Address', 'Address Line 1', 'Address Line'],
    'Location': ['Location', 'City', 'Town', 'Region'],
    'Pin Code': ['Pin Code', 'Pincode', 'Postal Code', 'Zip', 'Zip Code'],
    'Account Record ID': ['Account RecoRoadID', 'Account Record ID', 'Account ID', 'Account Id'],
    'First Name': ['First Name', 'FirstName', 'Fname', 'First', 'Given Name'],
    'Last Name': ['Last Name', 'LastName', 'Lname', 'Last', 'Surname', 'Family Name'],
    'Contact Name': ['Contact Name', 'Name', 'Contact Full Name'],
    'Phone Number': ['Phone', 'Phone Number', 'Telephone', 'Landline'],
    'Mobile Number': ['Mobile Number', 'Mobile', 'Mobile No', 'Mobile No.'],
    'Fax Number': ['Fax Number', 'Fax', 'Fax No', 'Fax No.'],
    'Email ID': ['Email ID', 'Email', 'Email Address', 'Email Addresss'],
    'Job Title': ['Job Title', 'Designation', 'Role'],
    'Department': ['Department', 'Dept'],
    'Contact Record ID': ['Contact Record ID', 'Contact ID', 'ContactID']
  };
}

function cleanData(records, ruleFlags, columnMapping, mode) {
  const synthesizedContactName = ensureContactNameColumn(records, columnMapping);
  const df = records.map((record) => ({ ...record, Issues: '' }));
  const issueSummary = {};
  const actualCols = buildActualColumns(df, columnMapping);
  ensureCorrectedColumns(df, actualCols);

  const accountCorrections = [
    {
      name: 'Account Name - Remove Brackets',
      columns: ['Account Name'],
      action: (value) => removeTextInBrackets(value)
    },
    {
      name: 'Account Name - Contains Location',
      columns: ['Account Name'],
      action: (value) => cleanAccountNameLocation(value)
    },
    {
      name: 'Account Name - Invalid Prefixes',
      columns: ['Account Name'],
      action: (value) => removeInvalidAccountPrefix(value)
    },
    {
      name: 'Account Name - Remove Punctuation (. and ,)',
      columns: ['Account Name'],
      action: (value) => normalizeText(toStringSafe(value).replace(/[.,]/g, ' '))
    },
    {
      name: 'Account Name - Ent/ent to Enterprise',
      columns: ['Account Name'],
      action: (value) => replaceEntToEnterprise(value)
    },
    {
      name: 'Account Name - Remove Dot: Private. Limited',
      columns: ['Account Name'],
      action: (value) => removePrivateDotLimited(value)
    },
    {
      name: 'Account Name - Short form correction',
      columns: ['Account Name'],
      action: (value) => {
        let next = normalizeText(value);
        const patterns = [
          { regex: /\bpvt\.?\s*ltd\.?\b/gi, replace: 'Private Limited' },
          { regex: /\bpvt\.?\b/gi, replace: 'Private' },
          { regex: /\bltd\.?\b/gi, replace: 'Limited' },
          { regex: /\bllp\b/gi, replace: 'LLP' },
          { regex: /\bbros\.?\b/gi, replace: 'Brothers' },
          { regex: /\b(?:&|and)\s+sons\b/gi, replace: 'and Sons' },
          { regex: /\b(?:&|and)\s+co\.?\b/gi, replace: 'Company' },
          { regex: /(^|\s)co\.?($|\s)/gi, replace: '$1Company$2' },
          { regex: /\bcorp\.?\b/gi, replace: 'Corporation' },
          { regex: /\binc\.?\b/gi, replace: 'Incorporated' },
          { regex: /\bmfg\.?\b/gi, replace: 'Manufacturing' },
          { regex: /\btrdrs\.?\b/gi, replace: 'Traders' },
          { regex: /\bassoc\.?\b/gi, replace: 'Associates' },
          { regex: /\bcons\.?\b/gi, replace: 'Consultants' },
          { regex: /\bsvc\.?\b/gi, replace: 'Services' },
          { regex: /\belec\.?\b/gi, replace: 'Electrical' },
          { regex: /\bengg\.?\b/gi, replace: 'Engineering' }
        ];
        patterns.forEach(({ regex, replace }) => {
          next = next.replace(regex, replace);
        });
        return normalizeText(next);
      }
    },
    {
      name: 'Account Name - Proper case',
      columns: ['Account Name'],
      action: (value) => properCase(stripSpecial(value))
    }
  ];

  const contactCorrections = [
    {
      name: 'Contact Name - Proper case',
      columns: ['Contact Name'],
      action: (value) => properCase(stripSpecial(value))
    },
    {
      name: 'Job Title - Proper case',
      columns: ['Job Title'],
      action: (value) => properCase(stripSpecial(value))
    },
    {
      name: 'Department - Proper case',
      columns: ['Department'],
      action: (value) => properCase(stripSpecial(value))
    }
  ];

  const addressCorrections = [
    {
      name: 'Address - Clean Leading/Trailing Punctuation',
      columns: ['Address'],
      action: (value) => stripTerminalDotsAndCommas(value)
    },
    {
      name: 'Address - Apartment',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\bapt\.?\b/gi, 'Apartment'))
    },
    {
      name: 'Address - Avenue',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(av[e]?|avn?)\.?\b/gi, 'Avenue'))
    },
    {
      name: 'Address - Building',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(bldg|bld)\.?\b/gi, 'Building'))
    },
    {
      name: 'Address - District',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(dist|dis)\.?\b/gi, 'District'))
    },
    {
      name: 'Address - Extension',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\bext\.?\b/gi, 'Extension'))
    },
    {
      name: 'Address - Floor',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(fl|flr)\.?\b/gi, 'Floor'))
    },
    {
      name: 'Address - Industrial',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(ind|indust)\.?\b/gi, 'Industrial'))
    },
    {
      name: 'Address - Lane',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(lane|ln)\.?\b/gi, 'Lane'))
    },
    {
      name: 'Address - Market',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(mkt|mrkt)\.?\b/gi, 'Market'))
    },
    {
      name: 'Address - Near',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(nr|near)\.?\b/gi, 'Near'))
    },
    {
      name: 'Address - Opposite',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(opp|oppo)\.?\b/gi, 'Opposite'))
    },
    {
      name: 'Address - Post Office',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(p\.?o\.?|post office)\b/gi, 'Post Office'))
    },
    {
      name: 'Address - Road',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(r\.?d\.?|road)\b/gi, 'Road'))
    },
    {
      name: 'Address - Street',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(str|st)\.?\b/gi, 'Street'))
    },
    {
      name: 'Address - Town',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(twn|tn)\.?\b/gi, 'Town'))
    },
    {
      name: 'Address - Village',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(vill|vlg)\.?\b/gi, 'Village'))
    },
    {
      name: 'Address - House Number',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\b(h\.?\s*no\.?|h\/?no\.?|house no\.?\b)/gi, 'House Number'))
    },
    {
      name: 'Address - Colony',
      columns: ['Address'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\bcol\.?\b/gi, 'Colony'))
    },
    {
      name: 'Address - Proper case',
      columns: ['Address'],
      action: (value) => properCase(value)
    }
  ];

  const locationCorrections = [
    {
      name: 'Location - Keep Only Last Value After Final Comma',
      columns: ['Location'],
      action: (value) => {
        const text = normalizeText(value);
        const parts = text.split(',').map((part) => part.trim()).filter(Boolean);
        return parts.length > 1 ? parts[parts.length - 1] : text;
      }
    },
    {
      name: 'Location - Remove Brackets & Content',
      columns: ['Location'],
      action: (value) => normalizeText(toStringSafe(value).replace(/[\(\[\{][^\)\]\}]*[\)\]\}]/g, ' '))
    },
    {
      name: 'Location - Remove Digits',
      columns: ['Location'],
      action: (value) => normalizeText(toStringSafe(value).replace(/\d+/g, ''))
    },
    {
      name: 'Location - Clean Leading/Trailing Dots & Commas',
      columns: ['Location'],
      action: (value) => stripTerminalDotsAndCommas(value)
    },
    {
      name: 'Location - Proper case',
      columns: ['Location'],
      action: (value) => properCase(value)
    }
  ];

  const formattingRules = [
    {
      name: 'Pin Code - Missing',
      columns: ['Pin Code'],
      action: (value) => value,
      validateMissing: true
    },
    {
      name: 'Job Title - Missing',
      columns: ['Job Title'],
      action: (value) => value,
      validateMissing: true
    },
    {
      name: 'Department - Missing',
      columns: ['Department'],
      action: (value) => value,
      validateMissing: true
    },
    {
      name: 'Mobile Number - Missing',
      columns: ['Mobile Number'],
      action: (value) => value,
      validateMissing: true
    },
    {
      name: 'Phone Number - Missing',
      columns: ['Phone Number'],
      action: (value) => value,
      validateMissing: true
    },
    {
      name: 'Fax Number - Missing',
      columns: ['Fax Number'],
      action: (value) => value,
      validateMissing: true
    },
    {
      name: 'Pin Code - 6 digits',
      columns: ['Pin Code'],
      action: (value) => cleanPinCode(value),
      valid: (value) => /^\d{6}$/.test(value)
    },
    {
      name: 'Phone Number - 11 digits',
      columns: ['Phone Number'],
      action: (value) => cleanPhone(value),
      valid: (value) => /^0\d{10}$/.test(value)
    },
    {
      name: 'Mobile Number - 10 digits',
      columns: ['Mobile Number'],
      action: (value) => cleanMobile(value),
      valid: (value) => /^\d{10}$/.test(value)
    },
    {
      name: 'Fax Number - Format/Length',
      columns: ['Fax Number'],
      action: (value) => cleanPhone(value),
      valid: (value) => /^0\d{10}$/.test(value)
    },
    {
      name: 'Email - lowercase',
      columns: ['Email ID'],
      action: (value) => cleanEmail(value),
      valid: (value) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)
    }
  ];

  const activeCorrections = [];
  if (mode === 'account') {
    activeCorrections.push(...accountCorrections, ...addressCorrections, ...locationCorrections, ...formattingRules);
  } else if (mode === 'contact') {
    activeCorrections.push(...contactCorrections, ...formattingRules);
  } else {
    activeCorrections.push(...accountCorrections, ...addressCorrections, ...locationCorrections, ...contactCorrections, ...formattingRules);
  }

  activeCorrections.forEach((rule) => {
    if (ruleFlags[rule.name] !== 'Y') {
      issueSummary[rule.name] = { count: 'Not Checked', corrections: 'Not Checked' };
      return;
    }
    const logical = rule.columns[0];
    if (!(logical in actualCols)) {
      issueSummary[rule.name] = { count: 'Not Checked', corrections: 'Not Checked' };
      return;
    }

    let issueCount = 0;
    let correctionCount = 0;
    const correctedKey = `${logical}_corrected`;

    df.forEach((row) => {
      const originalValue = toStringSafe(row[correctedKey]);
      const cleanedValue = rule.action(originalValue);
      const changed = cleanedValue !== originalValue;
      if (changed) {
        row[correctedKey] = cleanedValue;
        correctionCount += 1;
      }

      let flagged = false;
      if (rule.validateMissing) {
        if (!normalizeText(cleanedValue)) flagged = true;
      } else if (rule.valid) {
        if (!rule.valid(cleanedValue) && cleanedValue !== '') flagged = true;
      } else if (changed) {
        flagged = true;
      }

      if (flagged) {
        row.Issues = appendIssue(row.Issues, rule.name);
        issueCount += 1;
      }

      if ((rule.name === 'Account Name - Ent/ent to Enterprise' || rule.name === 'Account Name - Short form correction') && changed) {
        row._accountNameHighlight = true;
      }
    });

    issueSummary[rule.name] = { count: issueCount, corrections: correctionCount };
  });

  df.forEach((row) => {
    row.Issues = normalizeText(row.Issues);
    if (!row.Issues) row.Issues = null;
  });

  return { df, issueSummary };
}

/* =====================================================================
   DEDUPLICATION ENGINE
   - Account Record ID / Contact Record ID are treated as unique keys.
     They are never cleaned and are used only to reference matches.
   - Matching runs on the *_corrected fields so it benefits from the
     cleanup pass (proper case, punctuation, etc.) already applied above.
   ===================================================================== */

// Similarity threshold for fuzzy fields (Account Name, Address, Contact Name).
// Requested range was 65-70%; 0.68 sits in the middle of that band.
const DEDUPE_FUZZY_THRESHOLD = 0.68;

function levenshteinDistance(a, b) {
  if (a === b) return 0;
  const al = a.length;
  const bl = b.length;
  if (!al) return bl;
  if (!bl) return al;
  let prev = new Array(bl + 1);
  let curr = new Array(bl + 1);
  for (let j = 0; j <= bl; j += 1) prev[j] = j;
  for (let i = 1; i <= al; i += 1) {
    curr[0] = i;
    for (let j = 1; j <= bl; j += 1) {
      const cost = a.charCodeAt(i - 1) === b.charCodeAt(j - 1) ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    const swap = prev;
    prev = curr;
    curr = swap;
  }
  return prev[bl];
}

function fuzzyRatio(a, b) {
  const s1 = normalizeText(a).toLowerCase();
  const s2 = normalizeText(b).toLowerCase();
  if (!s1 && !s2) return 1;
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;
  const distance = levenshteinDistance(s1, s2);
  const maxLen = Math.max(s1.length, s2.length);
  return maxLen ? 1 - distance / maxLen : 1;
}

// Sorts words alphabetically before comparing, so "Sharma Traders" and
// "Traders Sharma" are recognised as the same name.
function tokenSortRatio(a, b) {
  const sortTokens = (value) => normalizeText(value).toLowerCase().split(/\s+/).filter(Boolean).sort().join(' ');
  return fuzzyRatio(sortTokens(a), sortTokens(b));
}

function isFuzzyMatch(a, b, threshold = DEDUPE_FUZZY_THRESHOLD) {
  const va = normalizeText(a);
  const vb = normalizeText(b);
  if (!va || !vb) return false;
  return tokenSortRatio(va, vb) >= threshold;
}

function isExactMatch(a, b) {
  const va = normalizeText(a).toLowerCase();
  const vb = normalizeText(b).toLowerCase();
  return Boolean(va) && Boolean(vb) && va === vb;
}

// Merges a newly found match into a Map<row, Map<otherId, status>>, keeping
// the strongest status ("Confirmed Duplicate" beats "Duplicate") if a pair
// is discovered more than once via different blocking buckets.
function recordMatch(matchMap, rowA, rowB, idA, idB, status) {
  [[rowA, idB], [rowB, idA]].forEach(([self, otherId]) => {
    if (!matchMap.has(self)) matchMap.set(self, new Map());
    const existing = matchMap.get(self);
    const currentStatus = existing.get(otherId);
    if (!currentStatus || (currentStatus !== 'Confirmed Duplicate' && status === 'Confirmed Duplicate')) {
      existing.set(otherId, status);
    }
  });
}

function applyMatchMap(records, matchMap, statusField, matchesField, idField) {
  records.forEach((row) => {
    const idStatusMap = matchMap.get(row);
    if (!idStatusMap || !idStatusMap.size) return;
    const ids = Array.from(idStatusMap.keys());
    const hasConfirmed = Array.from(idStatusMap.values()).includes('Confirmed Duplicate');
    row[statusField] = hasConfirmed ? 'Confirmed Duplicate' : 'Duplicate';
    row[matchesField] = ids.join(', ');
  });
}

/**
 * Account dedupe rule:
 *   Account Name (fuzzy) + Pin Code (exact) + Address (fuzzy)  -> "Duplicate"
 *   ...and if Phone Number also matches exactly               -> "Confirmed Duplicate"
 * Rows are blocked (grouped) by exact Pin Code first, since Pin Code must
 * match exactly anyway - this keeps the comparison fast on large files.
 */
function runAccountDeduplication(records, actualCols) {
  const nameKey = actualCols['Account Name'] ? 'Account Name_corrected' : null;
  const pinKey = actualCols['Pin Code'] ? 'Pin Code_corrected' : null;
  const addressKey = actualCols['Address'] ? 'Address_corrected' : null;
  const phoneKey = actualCols['Phone Number'] ? 'Phone Number_corrected' : null;
  const idKey = actualCols['Account Record ID'] || null;

  records.forEach((row) => {
    row.Account_Duplicate_Status = '';
    row.Account_Duplicate_Matches = '';
  });

  if (!nameKey || !pinKey || !addressKey) {
    records.forEach((row) => { row.Account_Duplicate_Status = 'Not Checked'; });
    return { ran: false, duplicateCount: 0, confirmedCount: 0, reason: 'Account Name, Pin Code, and Address columns are all required for account dedupe.' };
  }

  const idFor = (row, index) => (idKey ? normalizeText(row[idKey]) : '') || `Row ${index + 2}`;

  const buckets = new Map();
  records.forEach((row, index) => {
    const pin = normalizeText(row[pinKey]);
    if (!pin) return;
    if (!buckets.has(pin)) buckets.set(pin, []);
    buckets.get(pin).push({ row, index });
  });

  const matchMap = new Map();
  buckets.forEach((bucket) => {
    for (let i = 0; i < bucket.length; i += 1) {
      for (let j = i + 1; j < bucket.length; j += 1) {
        const { row: rowA, index: idxA } = bucket[i];
        const { row: rowB, index: idxB } = bucket[j];
        if (!isFuzzyMatch(rowA[nameKey], rowB[nameKey])) continue;
        if (!isFuzzyMatch(rowA[addressKey], rowB[addressKey])) continue;
        const phoneMatch = phoneKey ? isExactMatch(rowA[phoneKey], rowB[phoneKey]) : false;
        const status = phoneMatch ? 'Confirmed Duplicate' : 'Duplicate';
        recordMatch(matchMap, rowA, rowB, idFor(rowA, idxA), idFor(rowB, idxB), status);
      }
    }
  });

  applyMatchMap(records, matchMap, 'Account_Duplicate_Status', 'Account_Duplicate_Matches', idKey);

  const duplicateCount = records.filter((r) => r.Account_Duplicate_Status === 'Duplicate').length;
  const confirmedCount = records.filter((r) => r.Account_Duplicate_Status === 'Confirmed Duplicate').length;
  return { ran: true, duplicateCount, confirmedCount };
}

/**
 * Contact dedupe rule:
 *   Contact Name (fuzzy) + (Email OR Mobile exact match)       -> "Duplicate"
 *   ...and if BOTH Email and Mobile also match exactly         -> "Confirmed Duplicate"
 * Rows are blocked by exact Email, then exact Mobile, so only rows that
 * already share a contact detail are ever compared for name similarity.
 */
function runContactDeduplication(records, actualCols) {
  const nameKey = actualCols['Contact Name'] ? 'Contact Name_corrected' : null;
  const emailKey = actualCols['Email ID'] ? 'Email ID_corrected' : null;
  const mobileKey = actualCols['Mobile Number'] ? 'Mobile Number_corrected' : null;
  const idKey = actualCols['Contact Record ID'] || null;

  records.forEach((row) => {
    row.Contact_Duplicate_Status = '';
    row.Contact_Duplicate_Matches = '';
  });

  if (!nameKey || (!emailKey && !mobileKey)) {
    records.forEach((row) => { row.Contact_Duplicate_Status = 'Not Checked'; });
    return { ran: false, duplicateCount: 0, confirmedCount: 0, reason: 'Contact Name plus Email or Mobile Number are required for contact dedupe.' };
  }

  const idFor = (row, index) => (idKey ? normalizeText(row[idKey]) : '') || `Row ${index + 2}`;

  const buildBuckets = (key) => {
    const buckets = new Map();
    if (!key) return buckets;
    records.forEach((row, index) => {
      const value = normalizeText(row[key]).toLowerCase();
      if (!value) return;
      if (!buckets.has(value)) buckets.set(value, []);
      buckets.get(value).push({ row, index });
    });
    return buckets;
  };

  const matchMap = new Map();
  const considerBucket = (bucket) => {
    for (let i = 0; i < bucket.length; i += 1) {
      for (let j = i + 1; j < bucket.length; j += 1) {
        const { row: rowA, index: idxA } = bucket[i];
        const { row: rowB, index: idxB } = bucket[j];
        if (!isFuzzyMatch(rowA[nameKey], rowB[nameKey])) continue;
        const emailMatch = emailKey ? isExactMatch(rowA[emailKey], rowB[emailKey]) : false;
        const mobileMatch = mobileKey ? isExactMatch(rowA[mobileKey], rowB[mobileKey]) : false;
        const status = (emailMatch && mobileMatch) ? 'Confirmed Duplicate' : 'Duplicate';
        recordMatch(matchMap, rowA, rowB, idFor(rowA, idxA), idFor(rowB, idxB), status);
      }
    }
  };

  buildBuckets(emailKey).forEach(considerBucket);
  buildBuckets(mobileKey).forEach(considerBucket);

  applyMatchMap(records, matchMap, 'Contact_Duplicate_Status', 'Contact_Duplicate_Matches', idKey);

  const duplicateCount = records.filter((r) => r.Contact_Duplicate_Status === 'Duplicate').length;
  const confirmedCount = records.filter((r) => r.Contact_Duplicate_Status === 'Confirmed Duplicate').length;
  return { ran: true, duplicateCount, confirmedCount };
}

function runDeduplication(records, actualCols, mode) {
  const result = { account: null, contact: null };
  if (mode === 'account' || mode === 'combined') {
    result.account = runAccountDeduplication(records, actualCols);
  }
  if (mode === 'contact' || mode === 'combined') {
    result.contact = runContactDeduplication(records, actualCols);
  }
  return result;
}

const PROGRESS_STEPS = [
  { title: 'Uploading your file…', footnote: 'Reading bytes from your selected workbook.', fill: 10 },
  { title: 'Reading the workbook…', footnote: 'Parsing sheets, headers, and rows.', fill: 28 },
  { title: 'Cleaning & validating…', footnote: 'Matching each field against the cleanup rule set.', fill: 52 },
  { title: 'Checking for duplicates…', footnote: 'Comparing records against the dedupe rules.', fill: 76 },
  { title: 'Building your report…', footnote: 'Assembling summary, issues, and cleaned data sheets.', fill: 94 }
];

function showProgressOverlay() {
  if (!progressOverlay) return;
  progressOverlay.classList.remove('hidden');
  if (progressStepList) {
    Array.from(progressStepList.children).forEach((item) => item.classList.remove('active', 'done'));
  }
  if (progressFill) progressFill.style.width = '4%';
}

function hideProgressOverlay() {
  if (progressOverlay) progressOverlay.classList.add('hidden');
}

function setProcessStep(index) {
  const step = PROGRESS_STEPS[index];
  if (!step || !progressOverlay) return;
  if (progressTitle) progressTitle.textContent = step.title;
  if (progressFootnote) progressFootnote.textContent = step.footnote;
  if (progressFill) progressFill.style.width = `${step.fill}%`;
  if (progressStepList) {
    Array.from(progressStepList.children).forEach((item, position) => {
      item.classList.toggle('active', position === index);
      item.classList.toggle('done', position < index);
    });
  }
}

function completeProgressOverlay() {
  if (progressFill) progressFill.style.width = '100%';
  if (progressTitle) progressTitle.textContent = 'All done!';
  if (progressFootnote) progressFootnote.textContent = 'Insights are ready below.';
  if (progressStepList) {
    Array.from(progressStepList.children).forEach((item) => {
      item.classList.remove('active');
      item.classList.add('done');
    });
  }
  return new Promise((resolve) => setTimeout(() => {
    hideProgressOverlay();
    resolve();
  }, 450));
}

function openDetailOverlay(eyebrow, title, bodyHtml) {
  if (!detailOverlay) return;
  detailEyebrow.textContent = eyebrow;
  detailTitle.textContent = title;
  detailBody.innerHTML = bodyHtml;
  detailOverlay.classList.remove('hidden');
}

function closeDetailOverlay() {
  if (detailOverlay) detailOverlay.classList.add('hidden');
}

function escapeHtml(value) {
  return toStringSafe(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function findCorrectedColumnForIssue(issueName) {
  const prefix = issueName.split(' - ')[0];
  return `${prefix}_corrected`;
}

function buildIssueDetailHtml(issueName, entry) {
  if (!lastRunData) return '<p class="detail-empty">No data available.</p>';
  const { records } = lastRunData;
  const affected = records.filter((row) => normalizeText(row.Issues).split(',').map((s) => s.trim()).includes(issueName));

  const statsHtml = `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${typeof entry.count === 'number' ? entry.count : entry.count}</strong><span>Rows flagged</span></div>
      <div class="detail-stat"><strong>${typeof entry.corrections === 'number' ? entry.corrections : entry.corrections}</strong><span>Values corrected</span></div>
      <div class="detail-stat"><strong>${records.length}</strong><span>Total rows scanned</span></div>
    </div>`;

  if (!affected.length) {
    return `${statsHtml}<p class="detail-empty">No individual rows are currently flagged for this rule.</p>`;
  }

  const correctedKey = findCorrectedColumnForIssue(issueName);
  const nameKey = Object.keys(affected[0]).find((k) => /account name|contact name/i.test(k) && k.endsWith('_corrected')) || correctedKey;
  const preview = affected.slice(0, 25);

  const rows = preview.map((row, index) => {
    const identifier = row[nameKey] != null ? row[nameKey] : `Row ${index + 1}`;
    const value = row[correctedKey] != null ? row[correctedKey] : '';
    return `<tr><td>${escapeHtml(identifier)}</td><td>${escapeHtml(value)}</td><td>${escapeHtml(row.Issues)}</td></tr>`;
  }).join('');

  const moreNote = affected.length > preview.length
    ? `<p class="detail-more">Showing ${preview.length} of ${affected.length} affected rows. Download the full report to see them all.</p>`
    : '';

  return `
    ${statsHtml}
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>Record</th><th>Current value</th><th>All issues on row</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    ${moreNote}`;
}

function buildRowsWithIssuesDetailHtml() {
  if (!lastRunData) return '<p class="detail-empty">No data available.</p>';
  const { records } = lastRunData;
  const affected = records.filter((row) => row.Issues);
  const statsHtml = `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${affected.length}</strong><span>Rows with issues</span></div>
      <div class="detail-stat"><strong>${records.length}</strong><span>Total rows</span></div>
      <div class="detail-stat"><strong>${records.length ? Math.round((affected.length / records.length) * 100) : 0}%</strong><span>Flag rate</span></div>
    </div>`;
  if (!affected.length) {
    return `${statsHtml}<p class="detail-empty">No rows currently have issues. Nice and clean!</p>`;
  }
  const preview = affected.slice(0, 25);
  const rows = preview.map((row, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(row.Issues)}</td></tr>`).join('');
  const moreNote = affected.length > preview.length
    ? `<p class="detail-more">Showing ${preview.length} of ${affected.length} rows. Download the full report to see them all.</p>`
    : '';
  return `
    ${statsHtml}
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>#</th><th>Issues found</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    ${moreNote}`;
}

function buildTotalIssuesDetailHtml(analytics) {
  const rows = analytics.topIssues.map(([name, entry]) => `<tr><td>${escapeHtml(name)}</td><td>${entry.count}</td></tr>`).join('');
  return `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${analytics.totalIssues}</strong><span>Total issues</span></div>
      <div class="detail-stat"><strong>${analytics.topIssues.length}</strong><span>Distinct issue types shown</span></div>
    </div>
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>Issue</th><th>Count</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="2">No issues detected.</td></tr>'}</tbody>
      </table>
    </div>`;
}

function buildAnalytics(issueSummary, records, dedupeResults) {
  const summary = Object.entries(issueSummary)
    .filter(([_, entry]) => typeof entry.count === 'number')
    .sort((a, b) => b[1].count - a[1].count);
  const totalIssues = summary.reduce((sum, [_, entry]) => sum + entry.count, 0);
  const issueRows = records.filter((row) => row.Issues).length;
  return {
    totalRecords: records.length,
    totalIssues,
    issueRows,
    topIssues: summary.slice(0, 6),
    dedupe: dedupeResults || { account: null, contact: null }
  };
}

function buildDuplicateDetailHtml(recordType, statusFilter) {
  if (!lastRunData) return '<p class="detail-empty">No data available.</p>';
  const { records } = lastRunData;
  const statusField = recordType === 'account' ? 'Account_Duplicate_Status' : 'Contact_Duplicate_Status';
  const matchesField = recordType === 'account' ? 'Account_Duplicate_Matches' : 'Contact_Duplicate_Matches';
  const idField = recordType === 'account' ? 'Account Record ID' : 'Contact Record ID';
  const nameField = recordType === 'account' ? 'Account Name_corrected' : 'Contact Name_corrected';

  const affected = records.filter((row) => statusFilter.includes(row[statusField]));
  const total = affected.length;
  const confirmed = records.filter((row) => row[statusField] === 'Confirmed Duplicate').length;

  const statsHtml = `
    <div class="detail-stats">
      <div class="detail-stat"><strong>${total}</strong><span>Rows shown</span></div>
      <div class="detail-stat"><strong>${confirmed}</strong><span>Confirmed duplicates overall</span></div>
      <div class="detail-stat"><strong>${records.length}</strong><span>Total rows scanned</span></div>
    </div>`;

  if (!total) {
    return `${statsHtml}<p class="detail-empty">No duplicates of this type were found.</p>`;
  }

  const idKey = Object.keys(affected[0]).find((k) => normalizeHeaderValue(k) === normalizeHeaderValue(idField));
  const preview = affected.slice(0, 30);
  const rows = preview.map((row) => {
    const id = idKey ? row[idKey] : '—';
    const name = row[nameField] != null ? row[nameField] : '';
    const status = row[statusField];
    const matches = row[matchesField] || '';
    const badgeClass = status === 'Confirmed Duplicate' ? 'value-diff' : '';
    return `<tr><td>${escapeHtml(id)}</td><td>${escapeHtml(name)}</td><td>${escapeHtml(matches)}</td><td class="${badgeClass}"><span class="${status === 'Confirmed Duplicate' ? 'before' : ''}">${escapeHtml(status)}</span></td></tr>`;
  }).join('');

  const moreNote = affected.length > preview.length
    ? `<p class="detail-more">Showing ${preview.length} of ${affected.length} matching rows. Download the full report to see them all.</p>`
    : '';

  return `
    ${statsHtml}
    <div class="detail-table-wrap">
      <table>
        <thead><tr><th>Record ID</th><th>Name</th><th>Matches (Record IDs)</th><th>Status</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    ${moreNote}`;
}

function renderDedupeSummary(analytics, mode) {
  const section = document.getElementById('dedupeSection');
  const grid = document.getElementById('dedupeGrid');
  if (!section || !grid) return;
  grid.innerHTML = '';
  const { account, contact } = analytics.dedupe;
  const cards = [];

  if (mode === 'account' || mode === 'combined') {
    if (account && account.ran) {
      cards.push({
        label: 'Account duplicates',
        value: account.duplicateCount,
        onClick: () => openDetailOverlay('Duplicate detection', 'Account duplicates', buildDuplicateDetailHtml('account', ['Duplicate', 'Confirmed Duplicate']))
      });
      cards.push({
        label: 'Account confirmed duplicates',
        value: account.confirmedCount,
        statusClass: account.confirmedCount ? 'status-confirmed' : '',
        onClick: () => openDetailOverlay('Duplicate detection', 'Account confirmed duplicates', buildDuplicateDetailHtml('account', ['Confirmed Duplicate']))
      });
    } else if (account) {
      cards.push({ label: 'Account duplicates', value: 'Not checked', statusClass: 'status-muted', note: account.reason });
    }
  }

  if (mode === 'contact' || mode === 'combined') {
    if (contact && contact.ran) {
      cards.push({
        label: 'Contact duplicates',
        value: contact.duplicateCount,
        onClick: () => openDetailOverlay('Duplicate detection', 'Contact duplicates', buildDuplicateDetailHtml('contact', ['Duplicate', 'Confirmed Duplicate']))
      });
      cards.push({
        label: 'Contact confirmed duplicates',
        value: contact.confirmedCount,
        statusClass: contact.confirmedCount ? 'status-confirmed' : '',
        onClick: () => openDetailOverlay('Duplicate detection', 'Contact confirmed duplicates', buildDuplicateDetailHtml('contact', ['Confirmed Duplicate']))
      });
    } else if (contact) {
      cards.push({ label: 'Contact duplicates', value: 'Not checked', statusClass: 'status-muted', note: contact.reason });
    }
  }

  if (!cards.length) {
    section.classList.add('hidden');
    return;
  }
  section.classList.remove('hidden');

  cards.forEach((card, index) => {
    const block = document.createElement('div');
    block.className = `summary-card${card.statusClass ? ` ${card.statusClass}` : ''}`;
    block.style.animationDelay = `${index * 70}ms`;
    block.innerHTML = `<strong>${card.value}</strong><span>${card.label}</span>`;
    if (card.note) {
      block.title = card.note;
    }
    if (card.onClick) {
      block.classList.add('clickable');
      block.addEventListener('click', card.onClick);
    }
    grid.appendChild(block);
  });
}

function renderAnalytics(analytics, mode) {
  summaryGrid.innerHTML = '';
  const cards = [
    { label: 'Mode', value: mode === 'combined' ? 'Combined' : mode.charAt(0).toUpperCase() + mode.slice(1) },
    { label: 'Rows processed', value: analytics.totalRecords },
    {
      label: 'Total issues found',
      value: analytics.totalIssues,
      onClick: () => openDetailOverlay('Overview', 'Total issues found', buildTotalIssuesDetailHtml(analytics))
    },
    {
      label: 'Rows with issues',
      value: analytics.issueRows,
      onClick: () => openDetailOverlay('Overview', 'Rows with issues', buildRowsWithIssuesDetailHtml())
    }
  ];
  cards.forEach((card, index) => {
    const block = document.createElement('div');
    block.className = 'summary-card';
    block.style.animationDelay = `${index * 70}ms`;
    block.innerHTML = `<strong>${card.value}</strong><span>${card.label}</span>`;
    if (card.onClick) {
      block.classList.add('clickable');
      block.addEventListener('click', card.onClick);
    }
    summaryGrid.appendChild(block);
  });

  renderDedupeSummary(analytics, mode);

  issueTableBody.innerHTML = '';
  if (!analytics.topIssues.length) {
    issueTableBody.innerHTML = '<tr><td colspan="3">No issues were detected.</td></tr>';
    return;
  }
  analytics.topIssues.forEach(([name, entry], index) => {
    const row = document.createElement('tr');
    row.className = 'clickable-row';
    row.style.animationDelay = `${index * 60}ms`;
    row.innerHTML = `<td>${name}</td><td>${entry.count}</td><td class="row-action">View details ›</td>`;
    row.addEventListener('click', () => {
      openDetailOverlay('Issue detail', name, buildIssueDetailHtml(name, entry));
    });
    issueTableBody.appendChild(row);
  });

  analyticsCard.classList.add('reveal');
}

function buildDuplicateSheet(wb, sheetName, records, statusField, matchesField, columns, dedupeInfo) {
  const sheet = wb.addWorksheet(sheetName);
  if (!dedupeInfo || !dedupeInfo.ran) {
    sheet.addRow([(dedupeInfo && dedupeInfo.reason) || 'Not checked - required columns were not found in the uploaded file.']);
    autoSizeColumns(sheet);
    return;
  }

  const duplicateRows = records
    .map((row, index) => ({ row, rowNum: index + 2 }))
    .filter(({ row }) => row[statusField]);

  if (!duplicateRows.length) {
    sheet.addRow(['No duplicate records were found.']);
    autoSizeColumns(sheet);
    return;
  }

  const headers = ['Row', ...columns.map((c) => c.header), 'Status', 'Matches (Record IDs)'];
  sheet.addRow(headers);
  duplicateRows.forEach(({ row, rowNum }) => {
    const values = [rowNum, ...columns.map((c) => (row[c.key] != null ? row[c.key] : '')), row[statusField], row[matchesField] || ''];
    sheet.addRow(values);
  });
  styleHeaderRow(sheet);
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.autoFilter = { from: { row: 1, col: 1 }, to: { row: sheet.rowCount, col: headers.length } };

  const statusColIndex = headers.indexOf('Status') + 1;
  for (let i = 2; i <= sheet.rowCount; i += 1) {
    const cell = sheet.getRow(i).getCell(statusColIndex);
    if (cell.value === 'Confirmed Duplicate') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
      cell.font = { bold: true };
    } else if (cell.value === 'Duplicate') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    }
  }
  autoSizeColumns(sheet);
}

function buildReportWorkbook(records, issueSummary, mode, actualCols, dedupeResults) {
  const dedupe = dedupeResults || { account: null, contact: null };
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Account & Contact Clean Tool';
  wb.created = new Date();

  const executiveSummary = wb.addWorksheet('Executive Summary');
  executiveSummary.addRow(['Property', 'Value']);
  executiveSummary.addRow(['Mode', mode === 'combined' ? 'Combined' : `${mode.charAt(0).toUpperCase() + mode.slice(1)}`]);
  executiveSummary.addRow(['Rows processed', records.length]);
  executiveSummary.addRow(['Total issues found', Object.values(issueSummary).reduce((sum, item) => sum + (typeof item.count === 'number' ? item.count : 0), 0)]);
  if (dedupe.account) {
    executiveSummary.addRow(['Account duplicates found', dedupe.account.ran ? dedupe.account.duplicateCount : 'Not checked']);
    executiveSummary.addRow(['Account confirmed duplicates', dedupe.account.ran ? dedupe.account.confirmedCount : 'Not checked']);
  }
  if (dedupe.contact) {
    executiveSummary.addRow(['Contact duplicates found', dedupe.contact.ran ? dedupe.contact.duplicateCount : 'Not checked']);
    executiveSummary.addRow(['Contact confirmed duplicates', dedupe.contact.ran ? dedupe.contact.confirmedCount : 'Not checked']);
  }
  executiveSummary.addRow(['Generated at', new Date().toLocaleString()]);
  styleHeaderRow(executiveSummary);
  executiveSummary.views = [{ state: 'frozen', ySplit: 1 }];
  autoSizeColumns(executiveSummary);

  const issueSummaryRows = Object.entries(issueSummary).map(([name, item]) => ({
    'Issue name': name,
    'Issue count': item.count,
    'Corrections applied': item.corrections
  }));
  const sortedIssueSummaryRows = issueSummaryRows.sort((a, b) => {
    const aNotChecked = a['Issue count'] === 'Not Checked';
    const bNotChecked = b['Issue count'] === 'Not Checked';
    if (aNotChecked && !bNotChecked) return 1;
    if (!aNotChecked && bNotChecked) return -1;
    if (aNotChecked && bNotChecked) return 0;
    return b['Issue count'] - a['Issue count'];
  });
  const issueSummarySheet = wb.addWorksheet('Issue Summary');
  issueSummarySheet.addRow(['Issue name', 'Issue count', 'Corrections applied']);
  sortedIssueSummaryRows.forEach((row) => issueSummarySheet.addRow([row['Issue name'], row['Issue count'], row['Corrections applied']]));
  const totals = Object.values(issueSummary).reduce((sum, item) => sum + (typeof item.count === 'number' ? item.count : 0), 0);
  const totalsRow = issueSummarySheet.addRow(['Total', totals, '']);
  totalsRow.font = { bold: true };
  styleHeaderRow(issueSummarySheet);
  issueSummarySheet.views = [{ state: 'frozen', ySplit: 1 }];
  issueSummarySheet.autoFilter = { from: { row: 1, col: 1 }, to: { row: issueSummarySheet.rowCount, col: 3 } };
  autoSizeColumns(issueSummarySheet);

  const issueRecords = records
    .map((row, index) => ({ ...row, Row: index + 2 }))
    .filter((row) => row.Issues);
  const issueRecordSheet = wb.addWorksheet('Issue Records');
  if (issueRecords.length) {
    const issueRecordHeaders = Array.from(new Set(issueRecords.flatMap((row) => Object.keys(row)))).filter((key) => key !== '_accountNameHighlight');
    issueRecordSheet.addRow(issueRecordHeaders);
    issueRecords.forEach((row) => {
      issueRecordSheet.addRow(issueRecordHeaders.map((header) => row[header] != null ? row[header] : ''));
    });
    styleHeaderRow(issueRecordSheet);
    issueRecordSheet.views = [{ state: 'frozen', ySplit: 1 }];
    issueRecordSheet.autoFilter = { from: { row: 1, col: 1 }, to: { row: issueRecordSheet.rowCount, col: issueRecordHeaders.length } };
    autoSizeColumns(issueRecordSheet);
  } else {
    issueRecordSheet.addRow(['No issue records found']);
  }

  if (mode === 'account' || mode === 'combined') {
    const idKey = actualCols['Account Record ID'] || null;
    const columns = [
      idKey ? { header: 'Account Record ID', key: idKey } : null,
      { header: 'Account Name', key: 'Account Name_corrected' },
      { header: 'Pin Code', key: 'Pin Code_corrected' },
      { header: 'Address', key: 'Address_corrected' },
      actualCols['Phone Number'] ? { header: 'Phone Number', key: 'Phone Number_corrected' } : null
    ].filter(Boolean);
    buildDuplicateSheet(wb, 'Account Duplicates', records, 'Account_Duplicate_Status', 'Account_Duplicate_Matches', columns, dedupe.account);
  }

  if (mode === 'contact' || mode === 'combined') {
    const idKey = actualCols['Contact Record ID'] || null;
    const columns = [
      idKey ? { header: 'Contact Record ID', key: idKey } : null,
      { header: 'Contact Name', key: 'Contact Name_corrected' },
      actualCols['Email ID'] ? { header: 'Email ID', key: 'Email ID_corrected' } : null,
      actualCols['Mobile Number'] ? { header: 'Mobile Number', key: 'Mobile Number_corrected' } : null
    ].filter(Boolean);
    buildDuplicateSheet(wb, 'Contact Duplicates', records, 'Contact_Duplicate_Status', 'Contact_Duplicate_Matches', columns, dedupe.contact);
  }

  const fullHeaders = Array.from(new Set(records.flatMap((row) => Object.keys(row)))).filter((key) => key !== '_accountNameHighlight');
  const fullCleanedSheet = wb.addWorksheet('Full Cleaned Data');
  fullCleanedSheet.addRow(fullHeaders);
  records.forEach((row) => {
    fullCleanedSheet.addRow(fullHeaders.map((header) => (row[header] != null ? row[header] : '')));
  });
  styleHeaderRow(fullCleanedSheet);
  fullCleanedSheet.views = [{ state: 'frozen', ySplit: 1 }];
  fullCleanedSheet.autoFilter = { from: { row: 1, col: 1 }, to: { row: fullCleanedSheet.rowCount, col: fullHeaders.length } };

  const accountNameColIndex = fullHeaders.indexOf('Account Name_corrected') + 1;
  for (let i = 2; i <= fullCleanedSheet.rowCount; i += 1) {
    const row = fullCleanedSheet.getRow(i);
    if (i % 2 === 0) {
      row.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2F2F2' }
      };
    }
    const record = records[i - 2];
    if (record && record._accountNameHighlight && accountNameColIndex > 0) {
      const cell = row.getCell(accountNameColIndex);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFC7CE' }
      };
    }
  }
  autoSizeColumns(fullCleanedSheet);

  return wb;
}

function getSelectedMode() {
  const selected = document.querySelector('input[name="fileMode"]:checked');
  return selected ? selected.value : 'account';
}

function updateFileControls() {
  const mode = getSelectedMode();
  accountUpload.classList.toggle('hidden', mode !== 'account');
  contactUpload.classList.toggle('hidden', mode !== 'contact');
  combinedUpload.classList.toggle('hidden', mode !== 'combined');
  document.querySelectorAll('.radio-card').forEach((card) => {
    const input = card.querySelector('input');
    card.classList.toggle('selected', input.checked);
  });
}

function getSelectedFile(mode) {
  if (mode === 'account') return document.getElementById('accountFile').files[0];
  if (mode === 'contact') return document.getElementById('contactFile').files[0];
  return document.getElementById('combinedFile').files[0];
}

function pause(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function runComparison(event) {
  event.preventDefault();
  const mode = getSelectedMode();
  const file = getSelectedFile(mode);
  if (!file) {
    setStatus('Upload the file required for the selected mode.');
    return;
  }
  analyticsCard.classList.add('hidden');
  analyticsCard.classList.remove('reveal');
  setStatus('Reading workbook...');
  showProgressOverlay();
  setProcessStep(0);
  try {
    await pause(250);
    setProcessStep(1);
    const workbook = await readFileAsWorkbook(file);
    const records = workbookToJsonArray(workbook);
    setStatus('Cleaning data and checking issues...');
    await pause(200);
    setProcessStep(2);
    const columnMapping = buildColumnMapping();
    const ruleFlags = buildRuleFlags();
    const { df, issueSummary } = cleanData(records, ruleFlags, columnMapping, mode);
    setStatus('Checking for duplicate records...');
    await pause(200);
    setProcessStep(3);
    const actualCols = buildActualColumns(df, columnMapping);
    const dedupeResults = runDeduplication(df, actualCols, mode);
    const analytics = buildAnalytics(issueSummary, df, dedupeResults);
    setStatus('Building the management report...');
    await pause(200);
    setProcessStep(4);
    const wb = buildReportWorkbook(df, issueSummary, mode, actualCols, dedupeResults);
    lastExportWorkbook = wb;
    lastExportFilename = `AC_Data_Report_${mode}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    await saveWorkbookToFile(wb, lastExportFilename);
    lastRunData = { records: df, issueSummary, mode, dedupeResults };
    await completeProgressOverlay();
    renderAnalytics(analytics, mode);
    analyticsCard.classList.remove('hidden');
    setStatus('Cleanup complete. Download the report or review analytics above.');
  } catch (error) {
    hideProgressOverlay();
    setStatus(`Error: ${error.message}`);
    console.error(error);
  }
}

async function downloadReport() {
  if (!lastExportWorkbook) {
    setStatus('Run the comparison first to generate the report.');
    return;
  }
  await saveWorkbookToFile(lastExportWorkbook, lastExportFilename);
}

compareForm.addEventListener('submit', runComparison);
document.querySelectorAll('input[name="fileMode"]').forEach((radio) => {
  radio.addEventListener('change', updateFileControls);
});
downloadReportButton.addEventListener('click', downloadReport);
if (detailClose) detailClose.addEventListener('click', closeDetailOverlay);
if (detailOverlay) {
  detailOverlay.addEventListener('click', (event) => {
    if (event.target === detailOverlay) closeDetailOverlay();
  });
}
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') closeDetailOverlay();
});
updateFileControls();
setStatus('Ready to run your cleanup.');
