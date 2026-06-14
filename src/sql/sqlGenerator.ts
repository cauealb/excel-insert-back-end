import type { SqlValue } from "../types";

export function buildInsertScript(
  tableName: string,
  columns: string[],
  rows: Array<Record<string, SqlValue>>
): string {
  const quotedTable = quoteIdentifier(tableName);
  const quotedColumns = columns.map(quoteIdentifier).join(", ");
  const statements = rows.map((row) => {
    const values = columns.map((column) => formatSqlValue(row[column])).join(", ");
    return `INSERT INTO ${quotedTable} (${quotedColumns}) VALUES (${values});`;
  });

  return ["BEGIN;", ...statements, "COMMIT;"].join("\n");
}

function quoteIdentifier(identifier: string): string {
  if (!/^[a-z_][a-z0-9_]*$/i.test(identifier)) {
    throw new Error(`Unsafe SQL identifier in internal catalog: ${identifier}`);
  }

  return `"${identifier.replace(/"/g, "\"\"")}"`;
}

function formatSqlValue(value: SqlValue): string {
  if (value === null) {
    return "NULL";
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw new Error("Cannot format non-finite number as SQL value");
    }

    return String(value);
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  return `'${value.replace(/'/g, "''")}'`;
}
