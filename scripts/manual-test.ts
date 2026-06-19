import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { parseWorkbook } from "../src/excel/parseWorkbook";
import { generateSqlFromWorkbook } from "../src/sql/generateSqlFromWorkbook";

async function main(): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Usuarios");

  sheet.addRow(["Nome", "Email", "CPF", "Ativo"]);
  sheet.addRow(["Ana O'Brien", "ana@example.com", "529.982.247-25", "sim"]);
  sheet.addRow(["Bruno Silva", "bruno@example.com", "111.444.777-35", "nao"]);

  const outputDir = path.resolve("examples");
  await mkdir(outputDir, { recursive: true });

  const filePath = path.join(outputDir, "users-valid.xlsx");
  const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
  await writeFile(filePath, buffer);

  const parsedWorkbook = await parseWorkbook(buffer);
  const result = generateSqlFromWorkbook(parsedWorkbook, {
    workbookId: parsedWorkbook.id,
    sheetName: "Usuarios",
    table: {
      name: "users",
      columns: [
        { name: "name", label: "Nome", required: true, type: "string", maxLength: 120 },
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
  });

  console.log(`Arquivo de exemplo criado em: ${filePath}`);
  console.log("");
  console.log(result.sql);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
