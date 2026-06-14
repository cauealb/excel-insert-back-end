import { z } from "zod";
import { getEntityCatalog } from "../catalog/entities";
import { AppError, ValidationError } from "../errors";
import { buildInsertScript } from "./sqlGenerator";
import { cellValueToString, formatDateOnly } from "../excel/parseWorkbook";
import type {
  ExcelCellPrimitive,
  FieldCatalogEntry,
  GenerateSqlRequest,
  GenerateSqlResult,
  HeaderInfo,
  SheetRowSnapshot,
  SheetSnapshot,
  SqlValue,
  WorkbookSnapshot
} from "../types";

export const generateSqlRequestSchema = z
  .object({
    workbookId: z.string().uuid(),
    sheetName: z.string().min(1),
    entity: z.string().min(1),
    mapping: z.record(z.string().min(1), z.string().min(1))
  })
  .strict();

interface FieldColumnBinding {
  field: FieldCatalogEntry;
  header: HeaderInfo;
}

export function generateSqlFromWorkbook(
  workbook: WorkbookSnapshot,
  request: GenerateSqlRequest
): GenerateSqlResult {
  const entity = getEntityCatalog(request.entity);

  if (!entity) {
    throw new AppError(400, "UNKNOWN_ENTITY", `Entidade nao permitida: ${request.entity}`);
  }

  const sheet = workbook.sheets.find((candidate) => candidate.name === request.sheetName);

  if (!sheet) {
    throw new AppError(404, "SHEET_NOT_FOUND", `Aba nao encontrada: ${request.sheetName}`);
  }

  const bindings = buildBindings(sheet, entity.fields, request.mapping);
  const rows = validateRows(sheet, bindings);
  const columns = bindings.map(({ field }) => field.column);

  return {
    sql: buildInsertScript(entity.tableName, columns, rows),
    summary: {
      entity: entity.entity,
      table: entity.tableName,
      sheetName: sheet.name,
      insertedRows: rows.length,
      columns
    }
  };
}

function buildBindings(
  sheet: SheetSnapshot,
  fields: FieldCatalogEntry[],
  mapping: Record<string, string>
): FieldColumnBinding[] {
  const fieldsByName = new Map(fields.map((field) => [field.field, field]));

  if (Object.keys(mapping).length === 0) {
    throw new ValidationError({
      line: null,
      column: null,
      field: null,
      value: mapping,
      reason: "mapeamento vazio"
    });
  }

  for (const fieldName of Object.keys(mapping)) {
    if (!fieldsByName.has(fieldName)) {
      throw new ValidationError(
        {
          line: null,
          column: null,
          field: fieldName,
          value: mapping[fieldName],
          reason: "campo nao permitido no catalogo interno"
        },
        400
      );
    }
  }

  for (const field of fields) {
    if (field.required && !mapping[field.field]) {
      throw new ValidationError({
        line: null,
        column: null,
        field: field.field,
        value: null,
        reason: "campo obrigatorio nao foi mapeado"
      });
    }
  }

  const headersByName = groupHeadersByNormalizedName(sheet.headers);
  const bindings: FieldColumnBinding[] = [];
  const mappedHeaders = new Set<string>();

  for (const field of fields) {
    const requestedHeader = mapping[field.field];

    if (!requestedHeader) {
      continue;
    }

    const normalizedHeader = normalizeHeaderName(requestedHeader);
    const matches = headersByName.get(normalizedHeader) ?? [];

    if (matches.length === 0) {
      throw new ValidationError({
        line: 1,
        column: null,
        field: field.field,
        value: requestedHeader,
        reason: "coluna mapeada nao existe na aba selecionada"
      });
    }

    if (matches.length > 1) {
      throw new ValidationError({
        line: 1,
        column: matches[1].columnLetter,
        field: field.field,
        value: requestedHeader,
        reason: "coluna mapeada aparece mais de uma vez no cabecalho"
      });
    }

    const header = matches[0];

    if (mappedHeaders.has(normalizedHeader)) {
      throw new ValidationError({
        line: 1,
        column: header.columnLetter,
        field: field.field,
        value: requestedHeader,
        reason: "a mesma coluna Excel foi mapeada para mais de um campo"
      });
    }

    mappedHeaders.add(normalizedHeader);
    bindings.push({ field, header });
  }

  return bindings;
}

function validateRows(
  sheet: SheetSnapshot,
  bindings: FieldColumnBinding[]
): Array<Record<string, SqlValue>> {
  const uniqueTrackers = new Map<string, Map<string, number>>();

  for (const { field } of bindings) {
    if (field.unique) {
      uniqueTrackers.set(field.field, new Map<string, number>());
    }
  }

  const rows: Array<Record<string, SqlValue>> = [];

  for (const row of sheet.rows) {
    if (isCompletelyEmptyRow(row)) {
      continue;
    }

    const sqlRow: Record<string, SqlValue> = {};

    for (const binding of bindings) {
      const rawValue = row.valuesByColumnIndex[binding.header.columnIndex] ?? null;
      const value = validateCell(binding.field, rawValue, row, binding.header);

      if (binding.field.unique && value !== null) {
        const normalizedValue = normalizeUniqueValue(binding.field, value);
        const tracker = uniqueTrackers.get(binding.field.field);
        const previousLine = tracker?.get(normalizedValue);

        if (previousLine !== undefined) {
          throw new ValidationError({
            line: row.rowNumber,
            column: binding.header.columnLetter,
            field: binding.field.field,
            value: cellValueToString(rawValue),
            reason: `valor duplicado; ja apareceu na linha ${previousLine}`
          });
        }

        tracker?.set(normalizedValue, row.rowNumber);
      }

      sqlRow[binding.field.column] = value;
    }

    rows.push(sqlRow);
  }

  if (rows.length === 0) {
    throw new ValidationError({
      line: null,
      column: null,
      field: null,
      value: null,
      reason: "aba selecionada nao possui linhas de dados depois do cabecalho"
    });
  }

  return rows;
}

function validateCell(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo
): SqlValue {
  if (isEmptyValue(rawValue)) {
    if (field.required) {
      throw cellError(field, rawValue, row, header, "campo obrigatorio nao informado");
    }

    return null;
  }

  switch (field.type) {
    case "string":
      return validateString(field, rawValue, row, header);
    case "email":
      return validateEmail(field, rawValue, row, header);
    case "cpf":
      return validateCpf(field, rawValue, row, header);
    case "boolean":
      return validateBoolean(field, rawValue, row, header);
    case "date":
      return validateDate(field, rawValue, row, header);
    default:
      throw cellError(field, rawValue, row, header, "tipo de campo nao suportado");
  }
}

function validateString(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo
): string {
  const value = cellValueToString(rawValue).trim();

  if (field.maxLength && value.length > field.maxLength) {
    throw cellError(field, rawValue, row, header, `texto excede ${field.maxLength} caracteres`);
  }

  return value;
}

function validateEmail(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo
): string {
  const value = cellValueToString(rawValue).trim().toLowerCase();
  const result = z.string().email().safeParse(value);

  if (!result.success) {
    throw cellError(field, rawValue, row, header, "e-mail invalido");
  }

  if (field.maxLength && value.length > field.maxLength) {
    throw cellError(field, rawValue, row, header, `e-mail excede ${field.maxLength} caracteres`);
  }

  return value;
}

function validateCpf(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo
): string {
  const digits = cellValueToString(rawValue).replace(/\D/g, "");

  if (!isValidCpf(digits)) {
    throw cellError(field, rawValue, row, header, "CPF invalido");
  }

  return digits;
}

function validateBoolean(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo
): boolean {
  if (typeof rawValue === "boolean") {
    return rawValue;
  }

  if (typeof rawValue === "number") {
    if (rawValue === 1) {
      return true;
    }

    if (rawValue === 0) {
      return false;
    }
  }

  const normalized = cellValueToString(rawValue).trim().toLowerCase();
  const truthyValues = new Set(["1", "true", "sim", "s", "yes", "y", "ativo", "ativa"]);
  const falsyValues = new Set(["0", "false", "nao", "não", "n", "no", "inativo", "inativa"]);

  if (truthyValues.has(normalized)) {
    return true;
  }

  if (falsyValues.has(normalized)) {
    return false;
  }

  throw cellError(field, rawValue, row, header, "valor booleano invalido");
}

function validateDate(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo
): string {
  const date = parseDate(rawValue);

  if (!date) {
    throw cellError(field, rawValue, row, header, "data invalida");
  }

  return formatDateOnly(date);
}

function parseDate(value: ExcelCellPrimitive): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return excelSerialDateToDate(value);
  }

  const text = cellValueToString(value).trim();

  if (!text) {
    return null;
  }

  const brDate = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(text);

  if (brDate) {
    const [, day, month, year] = brDate;
    const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
    return isSameDateParts(date, Number(year), Number(month), Number(day)) ? date : null;
  }

  const isoDate = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);

  if (isoDate) {
    const [, year, month, day] = isoDate;
    const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
    return isSameDateParts(date, Number(year), Number(month), Number(day)) ? date : null;
  }

  return null;
}

function excelSerialDateToDate(serial: number): Date {
  const excelEpoch = Date.UTC(1899, 11, 30);
  return new Date(excelEpoch + serial * 24 * 60 * 60 * 1000);
}

function isSameDateParts(date: Date, year: number, month: number, day: number): boolean {
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() + 1 === month &&
    date.getUTCDate() === day
  );
}

function isValidCpf(cpf: string): boolean {
  if (!/^\d{11}$/.test(cpf) || /^(\d)\1{10}$/.test(cpf)) {
    return false;
  }

  const numbers = cpf.split("").map(Number);
  const firstCheckDigit = calculateCpfDigit(numbers.slice(0, 9), 10);
  const secondCheckDigit = calculateCpfDigit([...numbers.slice(0, 9), firstCheckDigit], 11);

  return numbers[9] === firstCheckDigit && numbers[10] === secondCheckDigit;
}

function calculateCpfDigit(numbers: number[], initialWeight: number): number {
  const sum = numbers.reduce((total, number, index) => total + number * (initialWeight - index), 0);
  const remainder = (sum * 10) % 11;
  return remainder === 10 ? 0 : remainder;
}

function cellError(
  field: FieldCatalogEntry,
  rawValue: ExcelCellPrimitive,
  row: SheetRowSnapshot,
  header: HeaderInfo,
  reason: string
): ValidationError {
  return new ValidationError({
    line: row.rowNumber,
    column: header.columnLetter,
    field: field.field,
    value: cellValueToString(rawValue),
    reason
  });
}

function isCompletelyEmptyRow(row: SheetRowSnapshot): boolean {
  return Object.values(row.valuesByColumnIndex).every(isEmptyValue);
}

function isEmptyValue(value: ExcelCellPrimitive): boolean {
  return value === null || (typeof value === "string" && value.trim() === "");
}

function normalizeUniqueValue(field: FieldCatalogEntry, value: SqlValue): string {
  if (value === null) {
    return "";
  }

  if (field.type === "email") {
    return String(value).trim().toLowerCase();
  }

  if (field.type === "cpf") {
    return String(value).replace(/\D/g, "");
  }

  return String(value).trim();
}

function groupHeadersByNormalizedName(headers: HeaderInfo[]): Map<string, HeaderInfo[]> {
  const result = new Map<string, HeaderInfo[]>();

  for (const header of headers) {
    const normalized = normalizeHeaderName(header.header);
    const current = result.get(normalized) ?? [];
    current.push(header);
    result.set(normalized, current);
  }

  return result;
}

function normalizeHeaderName(header: string): string {
  return header.trim().replace(/\s+/g, " ").toLowerCase();
}
