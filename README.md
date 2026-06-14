# Internal Excel SQL API

API interna em Node.js + TypeScript para validar planilhas Excel e gerar scripts SQL revisaveis. O MVP nao executa SQL no banco, nao recebe SQL bruto do usuario e so permite entidades/campos cadastrados no catalogo interno.

## Fluxo pensado para o front React

1. `GET /api/catalog`
   - Lista entidades e campos permitidos.
2. `POST /api/workbooks/upload`
   - Envia um `.xlsx` por multipart/form-data no campo `file`.
   - Retorna `workbookId`, abas e cabecalhos encontrados.
3. `GET /api/workbooks/:workbookId/sheets/:sheetName`
   - Retorna os cabecalhos de uma aba especifica para montar a tela de mapeamento.
4. `POST /api/sql/generate`
   - Recebe a entidade, a aba e o mapeamento entre campos permitidos e colunas Excel.
   - Valida linha a linha, para no primeiro erro e retorna linha, coluna, campo, valor e motivo.
   - Se tudo estiver valido, retorna um script com `BEGIN`, `INSERTs` e `COMMIT`.

## Rodando

```bash
npm install
npm run dev
```

Servidor padrao: `http://localhost:3333`.

## Testes

```bash
npm test
npm run manual:test
```

O teste manual cria `examples/users-valid.xlsx`, valida a planilha e imprime um SQL gerado.

## Entidade inicial: users

Campos permitidos:

| Campo API | Coluna SQL | Obrigatorio | Unico | Tipo |
| --- | --- | --- | --- | --- |
| `external_id` | `external_id` | nao | nao | string |
| `name` | `name` | sim | nao | string |
| `email` | `email` | sim | sim | email |
| `cpf` | `cpf` | sim | sim | cpf |
| `phone` | `phone` | nao | nao | string |
| `birth_date` | `birth_date` | nao | nao | date |
| `is_active` | `is_active` | nao | nao | boolean |

## Exemplo de geracao

```json
{
  "workbookId": "uuid-retornado-no-upload",
  "sheetName": "Usuarios",
  "entity": "users",
  "mapping": {
    "name": "Nome",
    "email": "Email",
    "cpf": "CPF",
    "is_active": "Ativo"
  }
}
```

Resposta de sucesso:

```json
{
  "sql": "BEGIN;\nINSERT INTO \"users\" ...\nCOMMIT;",
  "summary": {
    "entity": "users",
    "table": "users",
    "sheetName": "Usuarios",
    "insertedRows": 2,
    "columns": ["name", "email", "cpf", "is_active"]
  }
}
```

Resposta de erro de validacao:

```json
{
  "error": {
    "code": "VALIDATION_ERROR",
    "message": "Erro na linha 3, coluna B (email): valor duplicado; ja apareceu na linha 2",
    "details": {
      "line": 3,
      "column": "B",
      "field": "email",
      "value": "ana@example.com",
      "reason": "valor duplicado; ja apareceu na linha 2"
    }
  }
}
```
