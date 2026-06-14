import type { WorkbookSnapshot } from "../types";

export class InMemoryWorkbookStore {
  private readonly workbooks = new Map<string, WorkbookSnapshot>();

  save(workbook: WorkbookSnapshot): WorkbookSnapshot {
    this.pruneExpired();
    this.workbooks.set(workbook.id, workbook);
    return workbook;
  }

  get(workbookId: string): WorkbookSnapshot | null {
    this.pruneExpired();
    return this.workbooks.get(workbookId) ?? null;
  }

  private pruneExpired(): void {
    const now = Date.now();

    for (const [id, workbook] of this.workbooks.entries()) {
      if (Date.parse(workbook.expiresAt) <= now) {
        this.workbooks.delete(id);
      }
    }
  }
}

export const workbookStore = new InMemoryWorkbookStore();
