import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { ValidationError } from "../src/errors";
import { parseWorkbook } from "../src/excel/parseWorkbook";
import { generateSqlFromWorkbook } from "../src/sql/generateSqlFromWorkbook";

describe("SQL workflow", () => {
  it("validates a workbook against a dynamic table and generates escaped SQL", async () => {
    const buffer = await createWorkbook([
      ["Nome", "Email", "Idade", "Ignorar"],
      ["Ana O'Brien", "ANA@example.com", "31", "fora"],
      ["Bruno Silva", "bruno@example.com", 42, "fora"]
    ]);

    const workbook = await parseWorkbook(buffer, { id: "00000000-0000-4000-8000-000000000001" });
    const result = generateSqlFromWorkbook(workbook, {
      workbookId: workbook.id,
      sheetName: "Usuarios",
      table: {
        name: "public.customers",
        columns: [
          { name: "name", label: "Nome", required: true, type: "string", maxLength: 120 },
          { name: "email", label: "Email", required: true, type: "email", unique: true },
          { name: "age", label: "Idade", type: "number" },
          "notes"
        ]
      },
      mapping: {
        name: "Nome",
        email: "Email",
        age: "Idade"
      }
    });

    expect(result.summary.insertedRows).toBe(2);
    expect(result.summary.table).toBe("public.customers");
    expect(result.summary.columns).toEqual(["name", "email", "age"]);
    expect(result.sql).toContain("BEGIN;");
    expect(result.sql).toContain("COMMIT;");
    expect(result.sql).toContain('INSERT INTO "public"."customers" ("name", "email", "age")');
    expect(result.sql).toContain("'Ana O''Brien'");
    expect(result.sql).toContain("'ana@example.com'");
    expect(result.sql).toContain("31");
  });

  it("stops at the first duplicated unique value", async () => {
    const buffer = await createWorkbook([
      ["Nome", "Email", "CPF"],
      ["Ana", "ana@example.com", "529.982.247-25"],
      ["Outra Ana", "ANA@example.com", "111.444.777-35"]
    ]);

    const workbook = await parseWorkbook(buffer);

    expect(() =>
      generateSqlFromWorkbook(workbook, {
        workbookId: workbook.id,
        sheetName: "Usuarios",
        entity: "users",
        mapping: {
          name: "Nome",
          email: "Email",
          cpf: "CPF"
        }
      })
    ).toThrow(ValidationError);

    try {
      generateSqlFromWorkbook(workbook, {
        workbookId: workbook.id,
        sheetName: "Usuarios",
        entity: "users",
        mapping: {
          name: "Nome",
          email: "Email",
          cpf: "CPF"
        }
      });
    } catch (error) {
      expect(error).toBeInstanceOf(ValidationError);
      expect((error as ValidationError).details).toMatchObject({
        line: 3,
        column: "B",
        field: "email"
      });
    }
  });

  it("reports line, column, field, value and reason for invalid cells", async () => {
    const buffer = await createWorkbook([
      ["Nome", "Email", "CPF"],
      ["Ana", "ana@example.com", "123"]
    ]);

    const workbook = await parseWorkbook(buffer);

    try {
      generateSqlFromWorkbook(workbook, {
        workbookId: workbook.id,
        sheetName: "Usuarios",
        entity: "users",
        mapping: {
          name: "Nome",
          email: "Email",
          cpf: "CPF"
        }
      });

      throw new Error("Expected validation to fail");
    } catch (error) {
      expect(error).toBeInstanceOf(ValidationError);
      expect((error as ValidationError).details).toEqual({
        line: 2,
        column: "C",
        field: "cpf",
        value: "123",
        reason: "CPF invalido"
      });
    }
  });
});

async function createWorkbook(rows: unknown[][]): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Usuarios");

  for (const row of rows) {
    sheet.addRow(row);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
