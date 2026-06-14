import cors from "@fastify/cors";
import multipart from "@fastify/multipart";
import Fastify, { type FastifyInstance } from "fastify";
import { z } from "zod";
import { listCatalog } from "./catalog/entities";
import { AppError } from "./errors";
import { parseWorkbook } from "./excel/parseWorkbook";
import {
  generateSqlFromWorkbook,
  generateSqlRequestSchema
} from "./sql/generateSqlFromWorkbook";
import type { WorkbookSnapshot } from "./types";
import {
  InMemoryWorkbookStore,
  workbookStore as defaultWorkbookStore
} from "./workbooks/InMemoryWorkbookStore";

interface BuildAppOptions {
  store?: InMemoryWorkbookStore;
  logger?: boolean;
}

const workbookParamsSchema = z.object({
  workbookId: z.string().uuid()
});

const sheetParamsSchema = workbookParamsSchema.extend({
  sheetName: z.string().min(1)
});

export function buildApp(options: BuildAppOptions = {}): FastifyInstance {
  const store = options.store ?? defaultWorkbookStore;
  const app = Fastify({
    logger: options.logger ?? true
  });

  app.register(cors, {
    origin: true
  });

  app.register(multipart, {
    limits: {
      files: 1,
      fileSize: 15 * 1024 * 1024
    }
  });

  app.setErrorHandler((error, _request, reply) => {
    if (error instanceof AppError) {
      return reply.status(error.statusCode).send({
        error: {
          code: error.code,
          message: error.message,
          details: error.details
        }
      });
    }

    const requestError = error as { statusCode?: unknown; message?: unknown };
    const statusCode = typeof requestError.statusCode === "number" ? requestError.statusCode : 500;

    if (statusCode >= 400 && statusCode < 500) {
      return reply.status(statusCode).send({
        error: {
          code: "REQUEST_ERROR",
          message: typeof requestError.message === "string" ? requestError.message : "Requisicao invalida"
        }
      });
    }

    app.log.error(error);
    return reply.status(500).send({
      error: {
        code: "INTERNAL_SERVER_ERROR",
        message: "Erro interno inesperado"
      }
    });
  });

  app.get("/health", async () => ({
    status: "ok"
  }));

  app.get("/api/catalog", async () => ({
    entities: listCatalog()
  }));

  app.post("/api/workbooks/upload", async (request, reply) => {
    const file = await request.file();

    if (!file) {
      throw new AppError(400, "FILE_REQUIRED", "Envie um arquivo .xlsx no campo multipart file");
    }

    if (!file.filename.toLowerCase().endsWith(".xlsx")) {
      throw new AppError(400, "INVALID_FILE_TYPE", "O arquivo deve ter extensao .xlsx");
    }

    const chunks: Buffer[] = [];

    for await (const chunk of file.file) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }

    let workbook: WorkbookSnapshot;

    try {
      workbook = await parseWorkbook(Buffer.concat(chunks));
    } catch {
      throw new AppError(400, "INVALID_EXCEL", "Nao foi possivel ler o arquivo Excel enviado");
    }

    store.save(workbook);

    return reply.status(201).send(toWorkbookResponse(workbook));
  });

  app.get("/api/workbooks/:workbookId", async (request) => {
    const params = parseParams(workbookParamsSchema, request.params);
    const workbook = getWorkbookOrThrow(store, params.workbookId);
    return toWorkbookResponse(workbook);
  });

  app.get("/api/workbooks/:workbookId/sheets/:sheetName", async (request) => {
    const params = parseParams(sheetParamsSchema, request.params);
    const workbook = getWorkbookOrThrow(store, params.workbookId);
    const sheet = workbook.sheets.find((candidate) => candidate.name === params.sheetName);

    if (!sheet) {
      throw new AppError(404, "SHEET_NOT_FOUND", `Aba nao encontrada: ${params.sheetName}`);
    }

    return {
      workbookId: workbook.id,
      sheet
    };
  });

  app.post("/api/sql/generate", async (request) => {
    const result = generateSqlRequestSchema.safeParse(request.body);

    if (!result.success) {
      throw new AppError(400, "INVALID_BODY", "Corpo da requisicao invalido", result.error.flatten());
    }

    const workbook = getWorkbookOrThrow(store, result.data.workbookId);
    return generateSqlFromWorkbook(workbook, result.data);
  });

  return app;
}

function getWorkbookOrThrow(store: InMemoryWorkbookStore, workbookId: string): WorkbookSnapshot {
  const workbook = store.get(workbookId);

  if (!workbook) {
    throw new AppError(404, "WORKBOOK_NOT_FOUND", "Planilha nao encontrada ou expirada");
  }

  return workbook;
}

function parseParams<T>(schema: z.Schema<T>, params: unknown): T {
  const result = schema.safeParse(params);

  if (!result.success) {
    throw new AppError(400, "INVALID_PARAMS", "Parametros invalidos", result.error.flatten());
  }

  return result.data;
}

function toWorkbookResponse(workbook: WorkbookSnapshot) {
  return {
    workbookId: workbook.id,
    uploadedAt: workbook.uploadedAt,
    expiresAt: workbook.expiresAt,
    sheets: workbook.sheets.map((sheet) => ({
      name: sheet.name,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      headers: sheet.headers
    }))
  };
}
