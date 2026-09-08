export interface SheetRangeInspection {
  normalizedRange: string;
  sheetName?: string;
  sheetType?: string;
  rowCount?: number;
  columnCount?: number;
  requiredRows?: number;
  requiredColumns?: number;
  missingRows: number;
  missingColumns: number;
}

function columnNumber(label: string): number {
  return label
    .toUpperCase()
    .split("")
    .reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function unquoteSheetName(value: string): string {
  const trimmed = value.trim();
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/g, "'");
  }
  return trimmed;
}

export function normalizeA1Range(range: string): string {
  const normalized = range.trim();
  if (!normalized) throw new Error("Sheet range cannot be empty");
  return normalized.replace(/\s*!\s*/, "!");
}

function parseA1Extent(range: string): { rows?: number; columns?: number } {
  const localRange = range.includes("!") ? range.slice(range.lastIndexOf("!") + 1) : range;
  const endpoints = localRange.split(":");
  const end = endpoints[endpoints.length - 1];
  const match = end.match(/^\$?([A-Za-z]+)?\$?(\d+)?$/);
  if (!match) return {};
  return {
    columns: match[1] ? columnNumber(match[1]) : undefined,
    rows: match[2] ? Number(match[2]) : undefined,
  };
}

function parseA1Start(range: string): { row: number; column: number } {
  const localRange = range.includes("!") ? range.slice(range.lastIndexOf("!") + 1) : range;
  const start = localRange.split(":")[0];
  const match = start.match(/^\$?([A-Za-z]+)?\$?(\d+)?$/);
  return {
    column: match?.[1] ? columnNumber(match[1]) : 1,
    row: match?.[2] ? Number(match[2]) : 1,
  };
}

export function inspectSheetRange(
  range: string,
  spreadsheetInfo: unknown,
  valueShape?: { rows: number; columns: number },
): SheetRangeInspection {
  const normalizedRange = normalizeA1Range(range);
  const separator = normalizedRange.lastIndexOf("!");
  const requestedSheet =
    separator >= 0 ? unquoteSheetName(normalizedRange.slice(0, separator)) : undefined;
  const infoText =
    typeof spreadsheetInfo === "string"
      ? spreadsheetInfo
      : JSON.stringify(spreadsheetInfo);
  const lines = infoText.split(/\r?\n/);
  const sheetPattern =
    /^\s*-\s*"((?:[^"]|"")*)"\s+\(ID:\s*[^)]+\)\s*\|\s*Type:\s*([A-Z_]+)\s*\|\s*Size:\s*(\d+)x(\d+)/;
  const sheets = lines.flatMap((line) => {
    const match = line.match(sheetPattern);
    return match
      ? [{
          name: match[1].replace(/""/g, '"'),
          type: match[2],
          rows: Number(match[3]),
          columns: Number(match[4]),
        }]
      : [];
  });
  const selected = requestedSheet
    ? sheets.find((sheet) => sheet.name === requestedSheet)
    : sheets[0];
  if (!selected) {
    throw new Error(
      requestedSheet
        ? `Sheet '${requestedSheet}' was not found in spreadsheet metadata`
        : "Spreadsheet metadata did not contain a usable sheet",
    );
  }
  const required = parseA1Extent(normalizedRange);
  if (valueShape && valueShape.rows > 0 && valueShape.columns > 0) {
    const start = parseA1Start(normalizedRange);
    required.rows = Math.max(
      required.rows ?? 0,
      start.row + valueShape.rows - 1,
    );
    required.columns = Math.max(
      required.columns ?? 0,
      start.column + valueShape.columns - 1,
    );
  }
  return {
    normalizedRange,
    sheetName: selected.name,
    sheetType: selected.type,
    rowCount: selected.rows,
    columnCount: selected.columns,
    requiredRows: required.rows,
    requiredColumns: required.columns,
    missingRows: Math.max(0, (required.rows ?? selected.rows) - selected.rows),
    missingColumns: Math.max(
      0,
      (required.columns ?? selected.columns) - selected.columns,
    ),
  };
}
