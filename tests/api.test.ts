import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { buildApp } from "../src/app";
import { InMemoryWorkbookStore } from "../src/workbooks/InMemoryWorkbookStore";

describe("API workflow", () => {
  it("uploads an xlsx workbook and generates SQL through HTTP routes", async () => {
    const app = buildApp({ store: new InMemoryWorkbookStore(), logger: false });
    const workbookBuffer = await createWorkbook([
      ["Nome", "Email", "CPF", "Ativo"],
      ["Ana O'Brien", "ana@example.com", "529.982.247-25", "sim"]
    ]);

    const uploadResponse = await app.inject({
      method: "POST",
      url: "/api/workbooks/upload",
      headers: {
        "content-type": `multipart/form-data; boundary=${MULTIPART_BOUNDARY}`
      },
      payload: buildMultipartPayload(workbookBuffer)
    });

    expect(uploadResponse.statusCode).toBe(201);

    const uploadBody = uploadResponse.json<{
      workbookId: string;
      sheets: Array<{ name: string; headers: Array<{ header: string; columnLetter: string }> }>;
    }>();

    expect(uploadBody.sheets[0]).toMatchObject({
      name: "Usuarios",
      headers: [
        { header: "Nome", columnLetter: "A" },
        { header: "Email", columnLetter: "B" },
        { header: "CPF", columnLetter: "C" },
        { header: "Ativo", columnLetter: "D" }
      ]
    });

    const generateResponse = await app.inject({
      method: "POST",
      url: "/api/sql/generate",
      headers: {
        "content-type": "application/json"
      },
      payload: JSON.stringify({
        workbookId: uploadBody.workbookId,
        sheetName: "Usuarios",
        table: {
          name: "users",
          columns: [
            { name: "name", label: "Nome", required: true, type: "string" },
            { name: "email", label: "Email", required: true, type: "email", unique: true },
            { name: "cpf", label: "CPF", required: true, type: "cpf", unique: true },
            { name: "is_active", label: "Ativo", type: "boolean" }
          ]
        },
        mapping: {
          name: "Nome",
          email: "Email",
          cpf: "CPF",
          is_active: "Ativo"
        }
      })
    });

    expect(generateResponse.statusCode).toBe(200);
    expect(generateResponse.json<{ sql: string }>().sql).toContain(
      "INSERT INTO \"users\" (\"name\", \"email\", \"cpf\", \"is_active\")"
    );

    await app.close();
  }, 15000);
});

const MULTIPART_BOUNDARY = "----codex-api-test-boundary";

function buildMultipartPayload(fileBuffer: Buffer): Buffer {
  return Buffer.concat([
    Buffer.from(
      `--${MULTIPART_BOUNDARY}\r\n` +
        'Content-Disposition: form-data; name="file"; filename="users.xlsx"\r\n' +
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    ),
    fileBuffer,
    Buffer.from(`\r\n--${MULTIPART_BOUNDARY}--\r\n`)
  ]);
}

async function createWorkbook(rows: unknown[][]): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Usuarios");

  for (const row of rows) {
    sheet.addRow(row);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}
