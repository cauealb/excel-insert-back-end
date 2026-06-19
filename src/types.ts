export type EntityName = string;

export type FieldType = "string" | "email" | "cpf" | "boolean" | "date" | "number";

export interface FieldCatalogEntry {
  field: string;
  column: string;
  label: string;
  required: boolean;
  type: FieldType;
  unique?: boolean;
  maxLength?: number;
}

export interface EntityCatalogEntry {
  entity: EntityName;
  tableName: string;
  label: string;
  fields: FieldCatalogEntry[];
}

export type TableColumnDefinition =
  | string
  | {
      name: string;
      field?: string;
      column?: string;
      label?: string;
      required?: boolean;
      type?: FieldType;
      unique?: boolean;
      maxLength?: number;
    };

export interface DynamicTableDefinition {
  name: string;
  label?: string;
  columns: TableColumnDefinition[];
}

export type ExcelCellPrimitive = string | number | boolean | Date | null;

export interface HeaderInfo {
  header: string;
  columnIndex: number;
  columnLetter: string;
}

export interface SheetRowSnapshot {
  rowNumber: number;
  valuesByColumnIndex: Record<number, ExcelCellPrimitive>;
}

export interface SheetSnapshot {
  name: string;
  rowCount: number;
  columnCount: number;
  headers: HeaderInfo[];
  rows: SheetRowSnapshot[];
}

export interface WorkbookSnapshot {
  id: string;
  uploadedAt: string;
  expiresAt: string;
  sheets: SheetSnapshot[];
}

export interface ValidationIssue {
  line: number | null;
  column: string | null;
  field: string | null;
  value: unknown;
  reason: string;
}

export interface GenerateSqlRequest {
  workbookId: string;
  sheetName: string;
  entity?: string;
  table?: DynamicTableDefinition;
  mapping: Record<string, string>;
}

export type SqlValue = string | number | boolean | null;

export interface GenerateSqlResult {
  sql: string;
  summary: {
    entity: string;
    table: string;
    sheetName: string;
    insertedRows: number;
    columns: string[];
  };
}
