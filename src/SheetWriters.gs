/** Output writers */

function SHEETWR_deleteExistingPerformerTabs_(ctx) {
  const ss = ctx.ss;
  const contributors = READ_readContributors_(ctx);
  for (const [uuid, c] of contributors.entries()) {
    const name = CORE_safeSheetName_(`Performer – ${c.name} (${uuid})`);
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  }
}

function SHEETWR_deleteOverview_(ctx) {
  const ss = ctx.ss;
  const name = ctx.config.SHEETS.OVERVIEW_OUT;
  const sh = ss.getSheetByName(name);
  if (sh) ss.deleteSheet(sh);
}

/** Create or clear a sheet */
function SHEETWR_getOrCreate_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear({ contentsOnly: true });
  return sh;
}

/** Map a currency code to a human-friendly symbol for number formats.
 *  If we don't know it, return '' so we fall back to plain #,##0.00.
 */
function SHEETWR_currencySymbol_(ccy) {
  const map = {
    USD: '$', GBP: '£', EUR: '€', JPY: '¥', CNY: '¥',
    AUD: 'A$', CAD: 'C$', NZD: 'NZ$', HKD: 'HK$', SGD: 'S$',
    CHF: 'CHF', SEK: 'kr', NOK: 'kr', DKK: 'kr',
    PLN: 'zł', CZK: 'Kč', HUF: 'Ft', RON: 'lei',
    ILS: '₪', INR: '₹', KRW: '₩', RUB: '₽', TRY: '₺',
    MXN: 'MX$', BRL: 'R$', ARS: 'ARS$', ZAR: 'R',
    AED: 'AED', SAR: 'SAR', COP: 'COP', CLP: 'CLP', PEN: 'S/', EGP: 'E£', NGN: '₦'
  };
  return map[(ccy || '').toUpperCase()] || '';
}

/** Number formats, auto-resize, and status-based conditional formatting (no filter) */
function SHEETWR_format_(sh, currencyCode, dataStartRow) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  const startRow = dataStartRow || 2; // 3 for performer tabs (because of totals), 2 otherwise

  // Header styling
  const headers = sh.getRange(1, 1, 1, lastCol);
  headers.setFontWeight('bold').setWrap(true);

  // Freeze header + (optionally) totals
  sh.setFrozenRows(startRow - 1);

  // How many data rows to format/colour
  const dataCount = Math.max(0, sh.getLastRow() - (startRow - 1));

  if (dataCount > 0) {
    // Plays columns (F,G = 6,7)
    [6, 7].forEach(c => sh.getRange(startRow, c, dataCount, 1).setNumberFormat('#,##0'));

    // Money columns (H,I,J = 8,9,10)
    const symbol = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(currencyCode) : '';
    const moneyFormat = symbol
      ? `${symbol}#,##0.00`
      : (currencyCode ? `[$${String(currencyCode || '').toUpperCase()}] #,##0.00` : '#,##0.00');
    [8, 9, 10].forEach(c => sh.getRange(startRow, c, dataCount, 1).setNumberFormat(moneyFormat));

    // Variance % (K = 11)
    sh.getRange(startRow, 11, dataCount, 1).setNumberFormat('0.0%');

    // Row highlighting by Status (L = 12) across the FULL data block
    if (typeof SHEETWR_applyStatusConditionalFormats_ === 'function') {
      SHEETWR_applyStatusConditionalFormats_(sh, 12 /* Status col */, dataCount, lastCol, startRow);
    }
  }

  // If a totals row exists (startRow==3), format its numeric cells too
  if (startRow === 3 && sh.getLastRow() >= 2) {
    const tRow = 2;
    sh.getRange(tRow, 6, 1, 2).setNumberFormat('#,##0'); // plays
    const symbolT = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(currencyCode) : '';
    const moneyFormatT = symbolT
      ? `${symbolT}#,##0.00`
      : (currencyCode ? `[$${String(currencyCode || '').toUpperCase()}] #,##0.00` : '#,##0.00');
    sh.getRange(tRow, 8, 1, 3).setNumberFormat(moneyFormatT);   // expected/income/variance abs
    sh.getRange(tRow, 11, 1, 1).setNumberFormat('0.0%');       // variance %
    // Make totals row bold with a bottom border
    sh.getRange(tRow, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Auto-fit
  sh.autoResizeColumns(1, lastCol);
}

/** Apply row-highlighting based on the text in the Status column */
function SHEETWR_applyStatusConditionalFormats_(sh, statusColIndex, dataRows, lastCol, startRow) {
  const range = sh.getRange(startRow, 1, dataRows, lastCol);
  const colLetter = String.fromCharCode('A'.charCodeAt(0) + statusColIndex - 1);

  const rules = [
    { status: 'Missing',         color: '#f4cccc' },
    { status: 'Underperforming', color: '#fff2cc' },
    { status: 'Overperforming',  color: '#b6d7a8' },
    { status: 'On Track',        color: '#d9ead3' },
    { status: 'Not Expected',    color: '#eeeeee' },
    { status: 'Non-Qualifying',  color: '#eeeeee' },
    { status: 'Data Mismatch',   color: '#cfe2f3' }
  ].map(x =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${colLetter}${startRow}="${x.status}"`)
      .setBackground(x.color)
      .setRanges([range])
      .build()
  );

  sh.setConditionalFormatRules(rules);
}


/** Write a single performer tab */
function SHEETWR_writePerformerSheet_(ctx, contributor, currency, rows) {
  const ss = ctx.ss;
  const name = CORE_safeSheetName_(`Performer – ${contributor.name} (${contributor.id})`);
  const sh = SHEETWR_getOrCreate_(ss, name);

  // Clear old content (keep sheet & column widths/CF)
  sh.clear({ contentsOnly: true });

  // 1) Headers (replace (CCY) placeholder)
  const headers = ctx.config.OUTPUT_HEADERS_BASE.map(h =>
    h.replace(/\(CCY\)/g, `(${currency || ''})`)
  );
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 2) Totals row (row 2)
  let totalPlays = 0, totalQualified = 0, totalExpected = 0, totalIncome = 0;
  for (const r of rows) {
    totalPlays      += Number(r.playsTotal)     || 0;
    totalQualified  += Number(r.playsQualified) || 0;
    totalExpected   += Number(r.expected)       || 0;
    totalIncome     += Number(r.income)         || 0;
  }
  const varianceAbs = totalIncome - totalExpected;
  const variancePct = totalExpected > 0 ? (totalIncome / totalExpected) - 1 : '';

  // Columns: Contributor, UUID, Society, Territory, Registration Status,
  //          Plays (Total), Plays (Qualified), Expected, Income, Variance Abs, Variance %, Status, Investigate?, Notes
  const totalsRow = [
    'TOTALS', '', '', '', '',
    totalPlays, totalQualified, totalExpected, totalIncome, varianceAbs, variancePct,
    '', '', ''
  ];
  sh.getRange(2, 1, 1, headers.length).setValues([totalsRow]);

  // 3) Data rows from row 3
  const values = rows.map(r => ([
    r.contributor, r.uuid, r.society, r.territory, CORE_properCase_(r.registrationStatus),
    r.playsTotal, r.playsQualified,
    r.expected, r.income, r.varianceAbs, r.variancePct,
    r.status, r.investigate ? 'Yes' : 'No', r.notes
  ]));
  if (values.length) sh.getRange(3, 1, values.length, headers.length).setValues(values);

  // 4) Freeze header + totals
  sh.setFrozenRows(2);

  // 5) Format (expects SHEETWR_format_ to accept optional dataStartRow; extra args are safe)
  if (typeof SHEETWR_format_ === 'function') {
    SHEETWR_format_(sh, currency, /*dataStartRow=*/3);
  } else {
    // Fallback: at least auto-size columns
    sh.autoResizeColumns(1, headers.length);
  }
}

/** Write all performer tabs */
function SHEETWR_writeAllPerformerTabs_(ctx, rowsByContributor) {
  for (const [uuid, pack] of rowsByContributor.entries()) {
    SHEETWR_writePerformerSheet_(ctx, pack.contributor, pack.currency, pack.rows);
  }
}

/** Normalize overview input (Array | Map | Object) → flat Array of row objects. */
function SHEETWR_flattenOverviewRows_(rowsLike) {
  if (!rowsLike) return [];
  // Already an array
  if (Array.isArray(rowsLike)) return rowsLike;

  const out = [];

  // Map<uuid, { contributor, currency, rows: [...] }>
  if (rowsLike instanceof Map) {
    rowsLike.forEach(pack => {
      if (!pack || !Array.isArray(pack.rows)) return;
      const ccy = (pack.currency || pack.currencyCode || '').toString();
      for (const r of pack.rows) {
        if (r && !r.currency && !r.currencyCode && ccy) r.currency = ccy;
        out.push(r);
      }
    });
    return out;
  }

  // Plain object keyed by uuid
  if (typeof rowsLike === 'object') {
    // If it's a `{ rows: [...] }` shape
    if (Array.isArray(rowsLike.rows)) return rowsLike.rows;

    // Otherwise assume { [uuid]: { currency, rows } }
    Object.keys(rowsLike).forEach(k => {
      const pack = rowsLike[k];
      if (!pack || !Array.isArray(pack.rows)) return;
      const ccy = (pack.currency || pack.currencyCode || '').toString();
      for (const r of pack.rows) {
        if (r && !r.currency && !r.currencyCode && ccy) r.currency = ccy;
        out.push(r);
      }
    });
    return out;
  }

  return out;
}

/** Write the consolidated "Performer Overview" with per-row currency symbols */
function SHEETWR_writeOverview_(ctx, rowsLike) {
  const ss = ctx.ss;
  const sheetName = ctx.config.SHEETS.OVERVIEW_OUT || 'Performer Overview';
  const sh = SHEETWR_getOrCreate_(ss, sheetName);

  // Flatten input to an array and ensure per-row currency exists
  const rows = SHEETWR_flattenOverviewRows_(rowsLike);

  // Clear old content, keep widths/CF
  sh.clear({ contentsOnly: true });

  // 1) Headers: drop (CCY) since this sheet mixes currencies
  const headers = ctx.config.OUTPUT_HEADERS_BASE.map(h => h.replace(/\s*\(CCY\)/g, ''));
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 2) Data rows
  if (rows.length) {
    const values = rows.map(r => ([
      r.contributor,
      r.uuid,
      r.society,
      r.territory,
      CORE_properCase_(r.registrationStatus),
      r.playsTotal,
      r.playsQualified,
      r.expected,
      r.income,
      r.varianceAbs,
      r.variancePct,
      r.status,
      r.investigate ? 'Yes' : 'No',
      r.notes
    ]));
    sh.getRange(2, 1, values.length, headers.length).setValues(values);
  }

  // 3) Freeze the header
  sh.setFrozenRows(1);

  // 4) Per-column formats + per-row currency formats
  const n = rows.length;
  if (n > 0) {
    // Plays totals/qualified (F,G)
    sh.getRange(2, 6, n, 2).setNumberFormat('#,##0');

    // Build per-row money formats for H (Expected), I (Income), J (Variance Abs)
    const moneyFormats = [];
    for (let i = 0; i < n; i++) {
      const ccy =
        (rows[i].currency || rows[i].currencyCode || '').toString().toUpperCase() ||
        (ctx.config.DEFAULT_CURRENCY || 'USD');

      const symbol = (typeof SHEETWR_currencySymbol_ === 'function')
        ? SHEETWR_currencySymbol_(ccy)
        : '';

      const fmt = symbol ? `${symbol}#,##0.00` : (ccy ? `[$${ccy}] #,##0.00` : '#,##0.00');
      moneyFormats.push([fmt, fmt, fmt]);
    }
    sh.getRange(2, 8, n, 3).setNumberFormats(moneyFormats);

    // Variance % (K)
    sh.getRange(2, 11, n, 1).setNumberFormat('0.0%');

    // Status banding across each row (L column drives the color)
    if (typeof SHEETWR_applyStatusConditionalFormats_ === 'function') {
      SHEETWR_applyStatusConditionalFormats_(sh, 12 /* Status col */, n, headers.length, /*startRow=*/2);
    }
  }

  // 5) Tidy
  sh.autoResizeColumns(1, headers.length);
}

/** Build the "Totals" row for a performer sheet */
function SHEETWR_buildTotalsRow_(rows, contributor, uuid, currencyCode) {
  let totalPlays = 0, totalQualified = 0, totalExpected = 0, totalIncome = 0;

  for (const r of rows) {
    totalPlays += Number(r.playsTotal) || 0;
    totalQualified += Number(r.playsQualified) || 0;
    totalExpected += Number(r.expected) || 0;
    totalIncome += Number(r.income) || 0;
  }

  const varianceAbs = totalIncome - totalExpected;
  const variancePct = totalExpected > 0 ? (totalIncome / totalExpected) - 1 : '';

  // Columns: Contributor, UUID, Society, Territory, Registration Status,
  // Plays (Total), Plays (Qualified), Expected, Income, Variance Abs, Variance %, Status, Investigate?, Notes
  return [
    contributor || '', uuid || '', 'ALL SOCIETIES', '', 'Totals',
    totalPlays, totalQualified,
    totalExpected, totalIncome, varianceAbs, variancePct,
    '', '', ''
  ];
}

function SHEETWR_deleteSocietyIssues_(ctx) {
  const name = 'Society Issues';
  const sh = ctx.ss.getSheetByName(name);
  if (!sh) return;

  // Avoid deleting the only sheet in the file (Apps Script restriction)
  if (ctx.ss.getSheets().length > 1) {
    ctx.ss.deleteSheet(sh);
  } else {
    sh.clear({ contentsOnly: true });
  }
}


