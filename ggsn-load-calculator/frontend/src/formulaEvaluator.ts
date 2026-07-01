// formulaEvaluator.ts
// Handles Excel formula evaluation on the client side

export interface CellInfo {
  header: string;
  value: any;
  formula: string;
  is_formula: boolean;
}

export interface GridRow {
  row_num: number;
  [col_letter: string]: any; // either CellInfo or row_num
}

export interface RefData {
  [sheet_name: string]: any[][];
}

export interface GlobalRefs {
  [cell_coord: string]: any;
}

// Convert column letter to 0-based index
export function colLetterToIndex(col: string): number {
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function evaluateGrid(
  rows: GridRow[],
  refData: RefData,
  globalRefs: GlobalRefs
): GridRow[] {
  // We'll clone rows to avoid modifying props
  const workingRows = JSON.parse(JSON.stringify(rows)) as GridRow[];
  
  // To resolve dependencies, we can do multiple passes or topological sort.
  // Given Excel sheets can have circular dependencies, we'll do simple iterative resolution
  // (max 5 passes) which is extremely robust for GGSN load-balancing sheet.
  const maxPasses = 5;
  for (let pass = 0; pass < maxPasses; pass++) {
    let changed = false;
    
    for (let i = 0; i < workingRows.length; i++) {
      const row = workingRows[i];
      const rowNum = row.row_num;
      
      for (const colLetter in row) {
        if (colLetter === "row_num") continue;
        const cell = row[colLetter] as CellInfo;
        if (!cell || !cell.is_formula || !cell.formula) continue;
        
        try {
          const newVal = evaluateFormula(cell.formula, rowNum, workingRows, refData, globalRefs);
          if (newVal !== cell.value) {
            cell.value = newVal;
            changed = true;
          }
        } catch (e) {
          // If evaluation fails, keep current or mark as error
          console.warn(`Formula failed: ${cell.formula} at row ${rowNum}, col ${colLetter}`, e);
        }
      }
    }
    
    if (!changed) break; // converged
  }
  
  return workingRows;
}

function evaluateFormula(
  formula: string,
  currentRowNum: number,
  rows: GridRow[],
  refData: RefData,
  globalRefs: GlobalRefs
): any {
  // Strip '='
  let expr = formula.substring(1).trim();
  
  // Pre-process some functions to make parsing easier
  // 1. IFERROR(a, b) -> we'll evaluate custom function
  // 2. IFNA(a, b) / _xlfn.IFNA(a, b)
  // 3. VLOOKUP
  // 4. SUMIF
  // 5. SUM
  
  // Let's implement a recursive parser or regex evaluator
  // We'll evaluate VLOOKUPs first as they return concrete values
  expr = resolveVlookups(expr, currentRowNum, rows, refData);
  
  // Resolve SUMIFs
  expr = resolveSumifs(expr, rows);
  
  // Resolve SUMs
  expr = resolveSums(expr, rows);

  // Resolve LEFT
  expr = resolveLeft(expr, currentRowNum, rows);

  // Replace cell references (e.g. L3, AK3, $T$82)
  expr = replaceCellRefs(expr, currentRowNum, rows, globalRefs);
  
  // Replace percentage (e.g. 50% -> 0.5)
  expr = expr.replace(/(\d+(\.\d+)?)%/g, "($1/100)");

  // Replace Excel style operators/functions
  // IFERROR(val, fallback) -> we'll map to JS try/catch logic
  // Since we've replaced all formulas, let's parse logical operators and simple formulas
  try {
    const result = evalExcelMath(expr);
    return isNaN(result) ? 0 : result;
  } catch (err) {
    return 0;
  }
}

// Evaluate mathematical/logical expression safely
function evalExcelMath(expr: string): any {
  // Handle IFERROR / IFNA / IF
  // A simple way is to convert Excel functions to JS functions or handle them directly
  
  // Let's handle IFERROR(x, y)
  if (expr.toUpperCase().startsWith("IFERROR(")) {
    const content = expr.substring(8, expr.length - 1);
    const parts = splitParams(content);
    try {
      const val = evalSafe(parts[0]);
      if (val === null || val === undefined || isNaN(val) || !isFinite(val)) {
        return evalSafe(parts[1]);
      }
      return val;
    } catch {
      return evalSafe(parts[1]);
    }
  }

  if (expr.toUpperCase().startsWith("IFNA(") || expr.toUpperCase().startsWith("_XLFN.IFNA(")) {
    const startIdx = expr.toUpperCase().startsWith("IFNA(") ? 5 : 11;
    const content = expr.substring(startIdx, expr.length - 1);
    const parts = splitParams(content);
    try {
      const val = evalSafe(parts[0]);
      if (val === null || val === undefined) {
        return evalSafe(parts[1]);
      }
      return val;
    } catch {
      return evalSafe(parts[1]);
    }
  }

  if (expr.toUpperCase().startsWith("IF(")) {
    const content = expr.substring(3, expr.length - 1);
    const parts = splitParams(content);
    const cond = evalSafe(parts[0]);
    return cond ? evalSafe(parts[1]) : evalSafe(parts[2]);
  }

  return evalSafe(expr);
}

function evalSafe(str: string): any {
  let s = str.trim();
  if (s === "") return 0;
  
  // Replace standard excel syntax with JS
  s = s.replace(/<>/g, "!=");
  s = s.replace(/=/g, "==");
  
  // If it's a string literal like "NO" or "YES", return it
  if (s.startsWith('"') && s.endsWith('"')) {
    return s.substring(1, s.length - 1);
  }
  if (s.startsWith("'") && s.endsWith("'")) {
    return s.substring(1, s.length - 1);
  }

  if (s === "TRUE" || s === "true") return true;
  if (s === "FALSE" || s === "false") return false;

  try {
    // Only allow numbers and basic math operators
    if (/^[0-9\+\-\*\/\(\)\.\s!=<>]+$/.test(s)) {
      // eslint-disable-next-line no-eval
      return eval(s);
    }
    return s;
  } catch {
    return 0;
  }
}

function splitParams(str: string): string[] {
  const parts: string[] = [];
  let current = "";
  let depth = 0;
  let inQuotes = false;
  
  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    if (char === '"' || char === "'") {
      inQuotes = !inQuotes;
    }
    if (!inQuotes) {
      if (char === "(") depth++;
      if (char === ")") depth--;
      if (char === "," && depth === 0) {
        parts.push(current);
        current = "";
        continue;
      }
    }
    current += char;
  }
  parts.push(current);
  return parts;
}

// VLOOKUP(LEFT(B3,6),'GGSN-Lic'!A:AP,39,FALSE)
function resolveVlookups(expr: string, currentRowNum: number, rows: GridRow[], refData: RefData): string {
  // Regex to match VLOOKUPs: VLOOKUP(args)
  // Since args can contain commas, we must parse parameters carefully
  const vlookupRegex = /VLOOKUP\(/gi;
  let match;
  while ((match = vlookupRegex.exec(expr)) !== null) {
    const startIdx = match.index;
    let depth = 1;
    let endIdx = startIdx + 8;
    while (depth > 0 && endIdx < expr.length) {
      if (expr[endIdx] === '(') depth++;
      if (expr[endIdx] === ')') depth--;
      endIdx++;
    }
    
    const fullCall = expr.substring(startIdx, endIdx);
    const innerContent = expr.substring(startIdx + 8, endIdx - 1);
    const params = splitParams(innerContent);
    
    if (params.length >= 3) {
      const lookupKeyExpr = params[0].trim();
      const sheetRangeExpr = params[1].trim();
      const colIdxExpr = params[2].trim();
      
      // Evaluate lookupKey
      let lookupKey = evaluateFormula(`=${lookupKeyExpr}`, currentRowNum, rows, refData, {});
      if (typeof lookupKey === "string") lookupKey = lookupKey.trim();
      
      // Parse Sheet and Range
      // e.g. 'GGSN-Lic'!A:AP or Input_Lic_Rate_PS!K:P
      let sheetName = "";
      let rangeStr = "";
      if (sheetRangeExpr.includes("!")) {
        const parts = sheetRangeExpr.split("!");
        sheetName = parts[0].replace(/'/g, "").trim();
        rangeStr = parts[1].trim();
      } else {
        sheetName = sheetRangeExpr.replace(/'/g, "").trim();
      }
      
      const colIndex = parseInt(colIdxExpr, 10) - 1; // 1-based index to 0-based index
      
      const sheet = refData[sheetName];
      let foundValue: any = 0;
      
      if (sheet) {
        // Parse Column Range (e.g. A:AP, K:R)
        let startCol = 0;
        if (rangeStr && rangeStr.includes(":")) {
          const colParts = rangeStr.split(":");
          startCol = colLetterToIndex(colParts[0]);
        }
        
        // Find row in reference sheet where the cell at startCol matches lookupKey
        for (const row of sheet) {
          if (!row || row.length <= startCol) continue;
          let cellVal = row[startCol];
          if (typeof cellVal === "string") cellVal = cellVal.trim();
          
          if (String(cellVal) === String(lookupKey)) {
            // Found row! Return the value at colIndex
            // Note: colIndex in VLOOKUP is relative to the start of the range.
            // E.g., if range is K:P (K is 10th col), and colIndex is 4 (0-based 3), then we need column K + 3 = N.
            const targetCol = startCol + colIndex;
            if (row.length > targetCol) {
              foundValue = row[targetCol];
            }
            break;
          }
        }
      }
      
      // Replace in expr
      expr = expr.replace(fullCall, typeof foundValue === "number" ? String(foundValue) : `"${foundValue}"`);
      // Reset regex index
      vlookupRegex.lastIndex = 0;
    }
  }
  return expr;
}

// SUMIF($Q$3:$Q$70,Q3,$M$3:$M$70)
function resolveSumifs(expr: string, rows: GridRow[]): string {
  const sumifRegex = /SUMIF\(/gi;
  let match;
  while ((match = sumifRegex.exec(expr)) !== null) {
    const startIdx = match.index;
    let depth = 1;
    let endIdx = startIdx + 6;
    while (depth > 0 && endIdx < expr.length) {
      if (expr[endIdx] === '(') depth++;
      if (expr[endIdx] === ')') depth--;
      endIdx++;
    }
    
    const fullCall = expr.substring(startIdx, endIdx);
    const innerContent = expr.substring(startIdx + 6, endIdx - 1);
    const params = splitParams(innerContent);
    
    if (params.length >= 3) {
      const condRange = params[0].replace(/\$/g, "").trim(); // e.g. Q3:Q70
      const criteriaExpr = params[1].trim(); // e.g. Q3
      const sumRange = params[2].replace(/\$/g, "").trim(); // e.g. M3:M70
      
      // Extract rows and columns
      const condCol = condRange.match(/[A-Z]+/i)?.[0] || "";
      const condRows = condRange.match(/\d+/g)?.map(Number) || [3, 70];
      
      const sumCol = sumRange.match(/[A-Z]+/i)?.[0] || "";
      const sumRows = sumRange.match(/\d+/g)?.map(Number) || [3, 70];
      
      // Get criteria value
      const criteriaVal = evaluateFormula(`=${criteriaExpr}`, condRows[0], rows, {}, {});
      
      let sum = 0;
      const startR = condRows[0];
      const endR = condRows[1];
      
      for (let r = startR; r <= endR; r++) {
        const row = rows.find(x => x.row_num === r);
        if (!row) continue;
        const condCell = row[condCol] as CellInfo;
        const sumCell = row[sumCol] as CellInfo;
        
        if (condCell && String(condCell.value) === String(criteriaVal)) {
          if (sumCell && typeof sumCell.value === "number") {
            sum += sumCell.value;
          }
        }
      }
      
      expr = expr.replace(fullCall, String(sum));
      sumifRegex.lastIndex = 0;
    }
  }
  return expr;
}

// SUM($L$3:$L$31)
function resolveSums(expr: string, rows: GridRow[]): string {
  const sumRegex = /SUM\(/gi;
  let match;
  while ((match = sumRegex.exec(expr)) !== null) {
    const startIdx = match.index;
    let depth = 1;
    let endIdx = startIdx + 4;
    while (depth > 0 && endIdx < expr.length) {
      if (expr[endIdx] === '(') depth++;
      if (expr[endIdx] === ')') depth--;
      endIdx++;
    }
    
    const fullCall = expr.substring(startIdx, endIdx);
    const innerContent = expr.substring(startIdx + 4, endIdx - 1);
    
    // Check if range like L3:L31
    const cleanRange = innerContent.replace(/\$/g, "").trim();
    let sum = 0;
    
    if (cleanRange.includes(":")) {
      const parts = cleanRange.split(":");
      const col = parts[0].match(/[A-Z]+/i)?.[0] || "";
      const startRow = parseInt(parts[0].match(/\d+/)?.[0] || "3", 10);
      const endRow = parseInt(parts[1].match(/\d+/)?.[0] || "31", 10);
      
      for (let r = startRow; r <= endRow; r++) {
        const row = rows.find(x => x.row_num === r);
        if (row) {
          const cell = row[col] as CellInfo;
          if (cell && typeof cell.value === "number") {
            sum += cell.value;
          }
        }
      }
    } else {
      // Just split by comma
      const params = splitParams(cleanRange);
      for (const p of params) {
        const val = Number(p);
        if (!isNaN(val)) sum += val;
      }
    }
    
    expr = expr.replace(fullCall, String(sum));
    sumRegex.lastIndex = 0;
  }
  return expr;
}

// LEFT(B3, 6)
function resolveLeft(expr: string, currentRowNum: number, rows: GridRow[]): string {
  const leftRegex = /LEFT\(/gi;
  let match;
  while ((match = leftRegex.exec(expr)) !== null) {
    const startIdx = match.index;
    let depth = 1;
    let endIdx = startIdx + 5;
    while (depth > 0 && endIdx < expr.length) {
      if (expr[endIdx] === '(') depth++;
      if (expr[endIdx] === ')') depth--;
      endIdx++;
    }
    
    const fullCall = expr.substring(startIdx, endIdx);
    const innerContent = expr.substring(startIdx + 5, endIdx - 1);
    const params = splitParams(innerContent);
    
    if (params.length >= 2) {
      const cellExpr = params[0].trim();
      const numChars = parseInt(params[1].trim(), 10);
      
      const cellVal = String(evaluateFormula(`=${cellExpr}`, currentRowNum, rows, {}, {}));
      const leftVal = cellVal.substring(0, numChars);
      
      expr = expr.replace(fullCall, `"${leftVal}"`);
      leftRegex.lastIndex = 0;
    }
  }
  return expr;
}

// Replace L3, AK3, $T$82 with values
function replaceCellRefs(
  expr: string,
  currentRowNum: number,
  rows: GridRow[],
  globalRefs: GlobalRefs
): string {
  // Regex matches coordinate pattern: $?([A-Z]+)$?(\d+)
  // We exclude letters followed immediately by open parenthesis (functions like SUM, VLOOKUP)
  // and we also make sure not to replace sheet references in quotes or words.
  const refRegex = /\$?([A-Z]+)\$?(\d+)/g;
  return expr.replace(refRegex, (match, col, rowStr) => {
    const rowNum = parseInt(rowStr, 10);
    
    // Check if it is a global reference
    const coord = `${col}${rowNum}`;
    if (globalRefs[coord] !== undefined) {
      return String(globalRefs[coord]);
    }
    
    const targetRow = rows.find(x => x.row_num === rowNum);
    if (targetRow) {
      const cell = targetRow[col] as CellInfo;
      if (cell) {
        return typeof cell.value === "number" ? String(cell.value) : `"${cell.value}"`;
      }
    }
    return "0";
  });
}
