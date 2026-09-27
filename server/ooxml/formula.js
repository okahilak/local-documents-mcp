// Formula helpers: validation and relative-reference translation (A1 style, as stored in
// OOXML). Translation is used to promote a new master when a shared formula's master cell
// is overwritten.

const MAX_COL = 16384;
const MAX_ROW = 1048576;

export function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Strip a leading "=", and check length, quotes and parentheses. */
export function normalizeFormula(formula) {
  if (typeof formula !== "string") throw new Error("Formula must be a string.");
  const f = formula.trim().replace(/^=/, "").trim();
  if (!f) throw new Error("Formula is empty.");
  if (f.length > 8192) throw new Error("Formula is longer than Excel's 8192-character limit.");
  let depth = 0;
  let inString = false;
  let inQuote = false;
  for (const ch of f) {
    if (inString) {
      if (ch === '"') inString = false;
    } else if (inQuote) {
      if (ch === "'") inQuote = false;
    } else if (ch === '"') inString = true;
    else if (ch === "'") inQuote = true;
    else if (ch === "(") depth++;
    else if (ch === ")" && --depth < 0) throw new Error("Formula has unbalanced parentheses.");
  }
  if (inString) throw new Error("Formula has an unterminated string literal.");
  if (inQuote) throw new Error("Formula has an unterminated quoted sheet name.");
  if (depth !== 0) throw new Error("Formula has unbalanced parentheses.");
  return f;
}

const IDENT = /[A-Za-z0-9_.\\]/;

/**
 * Shift relative (non-$) references in `formula` by dRow/dCol. References that move off
 * the sheet become #REF!. Strings, quoted sheet names and [structured refs] are untouched.
 */
export function translateFormula(formula, dRow, dCol) {
  let out = "";
  let i = 0;
  const n = formula.length;
  const shiftCol = (abs, letters) => {
    if (abs) return abs + letters.toUpperCase();
    const c = colToNum(letters) + dCol;
    return c < 1 || c > MAX_COL ? null : numToCol(c);
  };
  const shiftRow = (abs, digits) => {
    if (abs) return abs + digits;
    const r = parseInt(digits, 10) + dRow;
    return r < 1 || r > MAX_ROW ? null : String(r);
  };
  while (i < n) {
    const ch = formula[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < n && !(formula[j] === '"' && formula[j + 1] !== '"')) j += formula[j] === '"' ? 2 : 1;
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      let j = i + 1;
      while (j < n && !(formula[j] === "'" && formula[j + 1] !== "'")) j += formula[j] === "'" ? 2 : 1;
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === "[") {
      let depth = 0;
      let j = i;
      for (; j < n; j++) {
        if (formula[j] === "[") depth++;
        else if (formula[j] === "]" && --depth === 0) break;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const prev = i > 0 ? formula[i - 1] : "";
    if (!IDENT.test(prev) || prev === "") {
      const rest = formula.slice(i);
      let m;
      // Column range A:C / $A:$C
      if ((m = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.!])/.exec(rest)) && colToNum(m[2]) <= MAX_COL && colToNum(m[4]) <= MAX_COL) {
        const a = shiftCol(m[1], m[2]);
        const b = shiftCol(m[3], m[4]);
        out += a && b ? `${a}:${b}` : "#REF!";
        i += m[0].length;
        continue;
      }
      // Row range 3:10 / $3:$10
      if ((m = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![0-9A-Za-z_(.])/.exec(rest))) {
        const a = shiftRow(m[1], m[2]);
        const b = shiftRow(m[3], m[4]);
        out += a && b ? `${a}:${b}` : "#REF!";
        i += m[0].length;
        continue;
      }
      // Cell A1 / $A$1 (not a function name like LOG10( and not a sheet name before !)
      if ((m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(.!])/.exec(rest)) && colToNum(m[2]) <= MAX_COL) {
        const c = shiftCol(m[1], m[2]);
        const r = shiftRow(m[3], m[4]);
        out += c && r ? `${c}${r}` : "#REF!";
        i += m[0].length;
        continue;
      }
      // Other identifiers (function names, defined names, sheet names): copy whole.
      if ((m = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(rest))) {
        out += m[0];
        i += m[0].length;
        continue;
      }
    }
    out += ch;
    i++;
  }
  return out;
}
