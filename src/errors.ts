import type { ValidationIssue } from "./types";

export class AppError extends Error {
  constructor(
    public readonly statusCode: number,
    public readonly code: string,
    message: string,
    public readonly details?: unknown
  ) {
    super(message);
  }
}

export class ValidationError extends AppError {
  constructor(issue: ValidationIssue, statusCode = 422) {
    super(statusCode, "VALIDATION_ERROR", formatValidationMessage(issue), issue);
  }
}

export function formatValidationMessage(issue: ValidationIssue): string {
  const location =
    issue.line === null
      ? "Erro de validacao"
      : `Erro na linha ${issue.line}${issue.column ? `, coluna ${issue.column}` : ""}`;
  const field = issue.field ? ` (${issue.field})` : "";
  return `${location}${field}: ${issue.reason}`;
}
