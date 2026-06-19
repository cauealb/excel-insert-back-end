# Internal Excel SQL API

API interna em Node.js + TypeScript para validar planilhas Excel e gerar scripts SQL revisaveis. O MVP nao executa SQL no banco e nao recebe SQL bruto do usuario. A tabela e as colunas podem ser informadas dinamicamente pelo cliente na hora de gerar o script.

## Fluxo pensado para o front React

1. `GET /api/catalog`
   - Lista a entidade demo legada e os tipos de campo aceitos.
2. `POST /api/workbooks/upload`
   - Envia um `.xlsx` por multipart/form-data no campo `file`.
   - Retorna `workbookId`, abas e cabecalhos encontrados.
3. `GET /api/workbooks/:workbookId/sheets/:sheetName`
   - Retorna os cabecalhos de uma aba especifica para montar a tela de mapeamento.
4. `POST /api/sql/generate`
   - Recebe a tabela, as colunas da tabela, a aba e o mapeamento entre colunas SQL e colunas Excel.
   - Campos sem mapeamento sao ignorados, exceto quando forem marcados como obrigatorios.
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

O teste manual cria `examples/users-valid.xlsx`, valida a planilha com uma tabela dinamica e imprime um SQL gerado.

## Definicao dinamica da tabela

No endpoint `POST /api/sql/generate`, envie `table` com o nome da tabela e as colunas permitidas. Cada coluna pode ser uma string simples ou um objeto com regras de validacao.

Tipos aceitos:

- `string`
- `email`
- `cpf`
- `boolean`
- `date`
- `number`

Exemplo de coluna completa:

```json
{
  "name": "email",
  "label": "E-mail",
  "required": true,
  "type": "email",
  "unique": true,
  "maxLength": 254
}
```

Por padrao, uma coluna string simples como `"name"` vira um campo opcional do tipo `string`, usando `name` tanto no mapeamento quanto no SQL.

Quando precisar separar o nome usado no mapeamento do nome SQL, use `field` e `column`. Exemplo: `{ "name": "nome_cliente", "field": "customerName", "column": "nome_cliente" }`.

## Entidade demo legada: users

O catalogo legado continua disponivel para compatibilidade, mas o fluxo principal deve usar `table`.

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
  "table": {
    "name": "users",
    "columns": [
      { "name": "name", "label": "Nome", "required": true, "type": "string", "maxLength": 120 },
      { "name": "email", "label": "E-mail", "required": true, "type": "email", "unique": true },
      { "name": "cpf", "label": "CPF", "required": true, "type": "cpf", "unique": true },
      { "name": "is_active", "label": "Ativo", "type": "boolean" }
    ]
  },
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
