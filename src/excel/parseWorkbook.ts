import { randomUUID } from "node:crypto";
import ExcelJS from "exceljs";
import type {
  ExcelCellPrimitive,
  HeaderInfo,
  SheetRowSnapshot,
  WorkbookSnapshot
} from "../types";

const DEFAULT_TTL_MS = 60 * 60 * 1000;

interface ParseWorkbookOptions {
  id?: string;
  now?: Date;
  ttlMs?: number;
}

export async function parseWorkbook(
  buffer: Buffer,
  options: ParseWorkbookOptions = {}
): Promise<WorkbookSnapshot> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as unknown as Parameters<ExcelJS.Xlsx["load"]>[0]);

  const now = options.now ?? new Date();
  const ttlMs = options.ttlMs ?? DEFAULT_TTL_MS;

  return {
    id: options.id ?? randomUUID(),
    uploadedAt: now.toISOString(),
    expiresAt: new Date(now.getTime() + ttlMs).toISOString(),
    sheets: workbook.worksheets.map((worksheet) => {
      const columnCount = Math.max(
        worksheet.actualColumnCount,
        worksheet.columnCount,
        worksheet.getRow(1).cellCount
      );

      return {
        name: worksheet.name,
        rowCount: worksheet.actualRowCount,
        columnCount,
        headers: extractHeaders(worksheet, columnCount),
        rows: extractRows(worksheet, columnCount)
      };
    })
  };
}

function extractHeaders(worksheet: ExcelJS.Worksheet, columnCount: number): HeaderInfo[] {
  const headerRow = worksheet.getRow(1);
  const headers: HeaderInfo[] = [];

  for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
    const header = cellValueToString(normalizeCellValue(headerRow.getCell(columnIndex).value)).trim();

    if (!header) {
      continue;
    }

    headers.push({
      header,
      columnIndex,
      columnLetter: columnNumberToName(columnIndex)
    });
  }

  return headers;
}

function extractRows(worksheet: ExcelJS.Worksheet, columnCount: number): SheetRowSnapshot[] {
  const rows: SheetRowSnapshot[] = [];

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const valuesByColumnIndex: Record<number, ExcelCellPrimitive> = {};

    for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
      valuesByColumnIndex[columnIndex] = normalizeCellValue(row.getCell(columnIndex).value);
    }

    rows.push({ rowNumber, valuesByColumnIndex });
  }

  return rows;
}

export function normalizeCellValue(value: ExcelJS.CellValue): ExcelCellPrimitive {
  if (value === null || value === undefined) {
    return null;
  }

  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  const maybeObject = value as unknown as Record<string, unknown>;

  if ("result" in maybeObject) {
    return normalizeCellValue(maybeObject.result as ExcelJS.CellValue);
  }

  if (typeof maybeObject.text === "string") {
    return maybeObject.text;
  }

  if (Array.isArray(maybeObject.richText)) {
    return maybeObject.richText
      .map((part) => (typeof part === "object" && part !== null && "text" in part ? String(part.text) : ""))
      .join("");
  }

  if (typeof maybeObject.hyperlink === "string" && typeof maybeObject.text === "string") {
    return maybeObject.text;
  }

  if (typeof maybeObject.error === "string") {
    return maybeObject.error;
  }

  return String(value);
}

export function cellValueToString(value: ExcelCellPrimitive): string {
  if (value === null) {
    return "";
  }

  if (value instanceof Date) {
    return formatDateOnly(value);
  }

  if (typeof value === "number" && Number.isInteger(value)) {
    return value.toFixed(0);
  }

  return String(value);
}

export function columnNumberToName(columnNumber: number): string {
  let current = columnNumber;
  let columnName = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    current = Math.floor((current - 1) / 26);
  }

  return columnName;
}

export function formatDateOnly(date: Date): string {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
