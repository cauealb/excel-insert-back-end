import type { EntityCatalogEntry } from "../types";

export const entityCatalog = {
  users: {
    entity: "users",
    tableName: "users",
    label: "Usuarios",
    fields: [
      {
        field: "external_id",
        column: "external_id",
        label: "ID externo",
        required: false,
        type: "string",
        maxLength: 64
      },
      {
        field: "name",
        column: "name",
        label: "Nome",
        required: true,
        type: "string",
        maxLength: 120
      },
      {
        field: "email",
        column: "email",
        label: "E-mail",
        required: true,
        type: "email",
        unique: true,
        maxLength: 254
      },
      {
        field: "cpf",
        column: "cpf",
        label: "CPF",
        required: true,
        type: "cpf",
        unique: true
      },
      {
        field: "phone",
        column: "phone",
        label: "Telefone",
        required: false,
        type: "string",
        maxLength: 32
      },
      {
        field: "birth_date",
        column: "birth_date",
        label: "Data de nascimento",
        required: false,
        type: "date"
      },
      {
        field: "is_active",
        column: "is_active",
        label: "Ativo",
        required: false,
        type: "boolean"
      }
    ]
  }
} satisfies Record<string, EntityCatalogEntry>;

export function getEntityCatalog(entity: string): EntityCatalogEntry | undefined {
  return entityCatalog[entity as keyof typeof entityCatalog];
}

export function listCatalog(): EntityCatalogEntry[] {
  return Object.values(entityCatalog);
}
