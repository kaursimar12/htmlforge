/**
 * Pure utilities for turning a flattened cell string into what a SpreadsheetML
 * cell needs: an A1-notation reference and a classification of its value
 * (a real number with an optional display format, or plain text). No XML/ZIP
 * concerns here — kept separate so the classification rules are easy to unit
 * test in isolation from the worksheet-building code that consumes them.
 */

/** 0-based column index → spreadsheet column letters (0→A, 25→Z, 26→AA, ...). */
export function colToLetter(colIndex0based) {
  let n = colIndex0based + 1;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** 0-based row/column → an A1-style cell reference, e.g. (0,0) → "A1". */
export function cellRef(row0based, col0based) {
  return colToLetter(col0based) + (row0based + 1);
}

const PLAIN_NUMBER = /^-?\d+(\.\d+)?$/;
const PERCENT = /^(-?\d+(\.\d+)?)%$/;
const CURRENCY = /^([$€£])(-?[\d,]+(\.\d+)?)$/;
const THOUSANDS = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/;

const CURRENCY_FORMAT_SYMBOL = { $: '"$"', '€': '"€"', '£': '"£"' };

function decimalPlaces(numericText) {
  const dot = numericText.indexOf('.');
  return dot === -1 ? 0 : numericText.length - dot - 1;
}

/**
 * Classify a trimmed cell string as a real number (with an optional display
 * format) or plain text — deliberately conservative: only whole/decimal
 * numbers, percentages, $/€/£ currency, and thousands-separated numbers are
 * recognized, always matched against the *entire* trimmed string (never a
 * substring), so things like "$(50.00)" (accounting negatives), "3 items",
 * non-$/€/£ currencies, or scientific notation all stay text rather than
 * risk silently misinterpreting a value Excel would then display wrong.
 *
 * Returns one of:
 *   { type: 'empty' }
 *   { type: 'number', value: string, numFmtCode?: string }
 *   { type: 'string' }
 */
export function classifyCellValue(text) {
  const trimmed = text.trim();
  if (trimmed === '') return { type: 'empty' };

  if (PLAIN_NUMBER.test(trimmed)) {
    return { type: 'number', value: trimmed };
  }

  const percentMatch = PERCENT.exec(trimmed);
  if (percentMatch) {
    const numericText = percentMatch[1];
    const value = String(Number(numericText) / 100);
    const decimals = decimalPlaces(numericText);
    const numFmtCode = decimals > 0 ? `0.${'0'.repeat(decimals)}%` : '0%';
    return { type: 'number', value, numFmtCode };
  }

  const currencyMatch = CURRENCY.exec(trimmed);
  if (currencyMatch) {
    const [, symbol, rawNumber] = currencyMatch;
    const value = rawNumber.replace(/,/g, '');
    if (!PLAIN_NUMBER.test(value)) return { type: 'string' };
    const decimals = decimalPlaces(value);
    const numFmtCode = decimals > 0
      ? `${CURRENCY_FORMAT_SYMBOL[symbol]}#,##0.${'0'.repeat(decimals)}`
      : `${CURRENCY_FORMAT_SYMBOL[symbol]}#,##0`;
    return { type: 'number', value, numFmtCode };
  }

  if (THOUSANDS.test(trimmed)) {
    const value = trimmed.replace(/,/g, '');
    const decimals = decimalPlaces(value);
    const numFmtCode = decimals > 0 ? `#,##0.${'0'.repeat(decimals)}` : '#,##0';
    return { type: 'number', value, numFmtCode };
  }

  return { type: 'string' };
}
